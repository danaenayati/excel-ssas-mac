"""A real Analysis Services instance behind the Backend interface.

Metadata comes from the standard MDSCHEMA rowsets; member lists and pivot results
come from MDX. One session is kept open and reused. If the server has dropped it
(idle timeout, VPN reconnect), the request is retried once on a fresh connection.
"""

from __future__ import annotations

import threading
from collections.abc import Callable

from ..errors import CubeError
from ..pivot import mdx
from ..pivot.model import CellSet, Layout
from ..pivot.names import validate_unique_name
from ..ssas import resolve
from ..ssas.session import Session


def _is_true(value: str | None) -> bool:
    return (value or "true").strip().lower() in ("true", "1")


def _valid(name: str | None) -> str | None:
    try:
        return validate_unique_name(name)
    except CubeError:
        return None


class SsasBackend:
    def __init__(
        self,
        server: str,
        username: str,
        password: str,
        *,
        session_factory: Callable[..., Session] = Session,
        lookup: Callable[[str, str], int] | None = None,
    ) -> None:
        self._target = resolve.parse_server(server)
        self._username = username
        self._password = password
        self._session_factory = session_factory
        self._lookup = lookup
        self._endpoint: tuple[str, int] | None = None
        self._session: Session | None = None
        self._lock = threading.Lock()
        # (catalog, cube) -> {hierarchy unique name: All member or None}
        self._all_members: dict[tuple[str, str], dict[str, str | None]] = {}

    # -- connection --------------------------------------------------------------
    def _open_session(self) -> tuple[Session, bool]:
        """Return a usable session and whether it had to be opened just now."""
        with self._lock:
            if self._session is not None and self._session.is_open:
                return self._session, False
            if self._endpoint is None:
                self._endpoint = resolve.endpoint(self._target, self._lookup)
            session = self._session_factory(self._endpoint[0], self._endpoint[1], self._username, self._password)
            session.open()
            self._session = session
            return session, True

    def _call(self, request: Callable[[Session], object]):
        session, fresh = self._open_session()
        try:
            return request(session)
        except CubeError as exc:
            if exc.category != "unreachable" or fresh:
                raise
        # The connection we had been holding was dead. Try once more on a new one.
        session, _ = self._open_session()
        return request(session)

    def cancel(self) -> None:
        session = self._session
        if session is not None:
            session.abort()

    def close(self) -> None:
        with self._lock:
            if self._session is not None:
                self._session.close()
            self._session = None

    # -- metadata ----------------------------------------------------------------
    def catalogs(self) -> list[dict]:
        rows = self._call(lambda s: s.discover("DBSCHEMA_CATALOGS"))
        found = [
            {"name": r["CATALOG_NAME"], "description": r.get("DESCRIPTION", "")}
            for r in rows
            if r.get("CATALOG_NAME")
        ]
        return sorted(found, key=lambda c: c["name"].lower())

    def cubes(self, catalog: str) -> list[dict]:
        rows = self._call(lambda s: s.discover("MDSCHEMA_CUBES", {"CATALOG_NAME": catalog}, catalog))
        found = []
        for r in rows:
            name = r.get("CUBE_NAME", "")
            # "$Name" entries are dimensions exposed as cubes, not something to browse.
            if not name or name.startswith("$") or r.get("CUBE_SOURCE") == "2":
                continue
            found.append({"name": name, "caption": r.get("CUBE_CAPTION") or name})
        return sorted(found, key=lambda c: c["caption"].lower())

    def fields(self, catalog: str, cube: str) -> dict:
        restrictions = {"CATALOG_NAME": catalog, "CUBE_NAME": cube}

        def load(session: Session):
            return (
                session.discover("MDSCHEMA_MEASURES", restrictions, catalog),
                session.discover("MDSCHEMA_DIMENSIONS", restrictions, catalog),
                session.discover("MDSCHEMA_HIERARCHIES", restrictions, catalog),
            )

        measure_rows, dimension_rows, hierarchy_rows = self._call(load)

        measures = []
        for r in measure_rows:
            unique = _valid(r.get("MEASURE_UNIQUE_NAME"))
            if unique is None or not _is_true(r.get("MEASURE_IS_VISIBLE")):
                continue
            measures.append(
                {
                    "uniqueName": unique,
                    "caption": r.get("MEASURE_CAPTION") or r.get("MEASURE_NAME") or unique,
                    "group": r.get("MEASUREGROUP_NAME", ""),
                    "folder": r.get("MEASURE_DISPLAY_FOLDER", ""),
                }
            )
        measures.sort(key=lambda m: (m["group"].lower(), m["folder"].lower(), m["caption"].lower()))

        dimensions: dict[str, dict] = {}
        for r in dimension_rows:
            unique = _valid(r.get("DIMENSION_UNIQUE_NAME"))
            if unique is None or r.get("DIMENSION_TYPE") == "2" or not _is_true(r.get("DIMENSION_IS_VISIBLE")):
                continue
            dimensions[unique] = {
                "uniqueName": unique,
                "caption": r.get("DIMENSION_CAPTION") or unique.strip("[]"),
                "hierarchies": [],
            }

        all_members: dict[str, str | None] = {}
        for r in hierarchy_rows:
            unique = _valid(r.get("HIERARCHY_UNIQUE_NAME"))
            dimension = dimensions.get(r.get("DIMENSION_UNIQUE_NAME", ""))
            if unique is None or dimension is None or not _is_true(r.get("HIERARCHY_IS_VISIBLE")):
                continue
            try:
                origin = int(r.get("HIERARCHY_ORIGIN") or 0)
            except ValueError:
                origin = 0
            all_member = _valid(r.get("ALL_MEMBER")) if r.get("ALL_MEMBER") else None
            all_members[unique] = all_member
            dimension["hierarchies"].append(
                {
                    "uniqueName": unique,
                    "caption": r.get("HIERARCHY_CAPTION") or unique,
                    "folder": r.get("HIERARCHY_DISPLAY_FOLDER", ""),
                    # Bit 1 marks a hierarchy the cube designer built from several levels.
                    "kind": "user" if origin & 1 else "attribute",
                    "allMember": all_member,
                }
            )
        self._all_members[(catalog, cube)] = all_members

        shown = [d for d in dimensions.values() if d["hierarchies"]]
        for d in shown:
            d["hierarchies"].sort(key=lambda h: h["caption"].lower())
        shown.sort(key=lambda d: d["caption"].lower())
        return {"measures": measures, "dimensions": shown}

    # -- members and queries -----------------------------------------------------
    def members(self, catalog: str, cube: str, hierarchy: str, parent: str | None, search: str | None, limit: int) -> dict:
        if (catalog, cube) not in self._all_members:
            self.fields(catalog, cube)
        known = self._all_members[(catalog, cube)]
        if hierarchy not in known:
            raise CubeError("not-found", "That field is no longer in the cube. Reload the field list.", hierarchy)
        all_member = known[hierarchy]
        text = mdx.members_query(cube, hierarchy, all_member, parent, search, limit)
        cellset = self._call(lambda s: s.execute(text, catalog))
        tuples = cellset.axes[1] if len(cellset.axes) > 1 else []
        found = [
            {
                "uniqueName": t[0].unique_name,
                "caption": t[0].caption,
                "level": t[0].level,
                "hasChildren": t[0].child_count > 0,
            }
            for t in tuples
            if t and t[0].unique_name != all_member
        ]
        return {"members": found[:limit], "truncated": len(found) > limit}

    def query(self, catalog: str, cube: str, layout: Layout) -> tuple[CellSet, str]:
        text = mdx.build_query(cube, layout)
        return self._call(lambda s: s.execute(text, catalog)), text
