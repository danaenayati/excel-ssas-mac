"""A real Analysis Services instance behind the Backend interface.

Metadata comes from the standard MDSCHEMA rowsets; member lists and pivot results
come from MDX. One session is kept open and reused. If the server has dropped it
(idle timeout, VPN reconnect), the request is retried once on a fresh connection.

Cancelling is aimed at queries only. It shuts the shared connection down, so a
metadata request that happened to be on it is quietly repeated, while the query
reports "cancelled". A cancel that arrives when no query is running does nothing.
"""

from __future__ import annotations

import threading
from collections.abc import Callable

from ..errors import CubeError
from ..pivot import mdx
from ..pivot.grid import clean_value
from ..pivot.model import CellSet, Layout
from ..pivot.names import validate_unique_name
from ..ssas import resolve, xmla
from ..ssas.session import Session
from .base import plain_column_names


# OLE DB types of a level's key that hold dates (DATE, DBDATE, DBTIMESTAMP), and the
# level type of a day level in a time dimension.
_DATE_DBTYPES = frozenset({"7", "133", "135"})
_LEVEL_TYPE_ALL = 1
_LEVEL_TYPE_DAYS = 0x204


def _optional(load: Callable[[], list[dict[str, str]]]) -> list[dict[str, str]]:
    """A schema rowset that only adds to the field list (levels, KPIs, sets): when the
    server refuses it, the field list goes without what it would add."""
    try:
        return load()
    except CubeError as exc:
        if exc.category in ("server", "bad-request", "not-found"):
            return []
        raise


def _int(value: str | None) -> int:
    try:
        return int(value or 0)
    except ValueError:
        return 0


def _date_hierarchies(level_rows: list[dict[str, str]]) -> set[str]:
    """Hierarchies whose first level below All holds dates: Date Filters apply to them."""
    first: dict[str, dict[str, str]] = {}
    for r in level_rows:
        hierarchy = r.get("HIERARCHY_UNIQUE_NAME")
        if not hierarchy or _int(r.get("LEVEL_TYPE")) == _LEVEL_TYPE_ALL:
            continue
        if hierarchy not in first or _int(r.get("LEVEL_NUMBER")) < _int(first[hierarchy].get("LEVEL_NUMBER")):
            first[hierarchy] = r
    return {
        h
        for h, r in first.items()
        if r.get("LEVEL_DBTYPE") in _DATE_DBTYPES or _int(r.get("LEVEL_TYPE")) == _LEVEL_TYPE_DAYS
    }


def _kpis(rows: list[dict[str, str]]) -> list[dict]:
    out = []
    for r in rows:
        name, measure = r.get("KPI_NAME"), _valid(r.get("KPI_VALUE"))
        if not name or measure is None:
            continue
        parts = ["value"] + [part for part in ("goal", "status", "trend") if r.get(f"KPI_{part.upper()}")]
        out.append(
            {
                "name": name,
                "caption": r.get("KPI_CAPTION") or name,
                "folder": r.get("KPI_DISPLAY_FOLDER", ""),
                "measure": measure,
                "parts": parts,
            }
        )
    return sorted(out, key=lambda k: k["caption"].lower())


def _sets(rows: list[dict[str, str]], dimensions: set[str], all_members: dict[str, str | None]) -> list[dict]:
    out = []
    for r in rows:
        name = r.get("SET_NAME")
        if not name or any(ord(ch) < 0x20 for ch in name):
            continue
        hierarchies = [h for h in (_valid(part.strip()) for part in _split_names(r.get("DIMENSIONS", ""))) if h]
        dimension = hierarchies[0].split("].[")[0] + "]" if hierarchies else ""
        hierarchy = hierarchies[0] if len(hierarchies) == 1 and hierarchies[0].count("].[") == 1 else None
        out.append(
            {
                "uniqueName": "[" + name.replace("]", "]]") + "]",
                "caption": r.get("SET_CAPTION") or name,
                "folder": r.get("SET_DISPLAY_FOLDER", ""),
                "dimension": dimension if dimension in dimensions else "",
                "hierarchy": hierarchy,
                "allMember": all_members.get(hierarchy) if hierarchy else None,
            }
        )
    return sorted(out, key=lambda s: s["caption"].lower())


def _split_names(text: str) -> list[str]:
    """Split "[A].[B],[C].[D]" on the commas outside brackets."""
    parts, current, depth = [], [], 0
    for ch in text:
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth = max(0, depth - 1)
        if ch == "," and depth == 0:
            parts.append("".join(current))
            current = []
        else:
            current.append(ch)
    if current:
        parts.append("".join(current))
    return [p for p in parts if p.strip()]


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
        self._pending: Session | None = None  # being opened; cancel and close can reach it
        self._closed = False
        self._lock = threading.Lock()  # guards opening and replacing the session
        self._state = threading.Lock()  # guards the two counters below
        self._active_queries = 0
        self._cancel_epoch = 0  # bumped by every cancel that had a query to stop
        # (catalog, cube) -> {hierarchy unique name: All member or None}
        self._all_members: dict[tuple[str, str], dict[str, str | None]] = {}
        # (catalog, cube) -> {measure unique name: its default FORMAT_STRING}
        self._measure_formats: dict[tuple[str, str], dict[str, str]] = {}

    # -- connection --------------------------------------------------------------
    def _open_session(self) -> tuple[Session, bool]:
        """Return a usable session and whether it had to be opened just now."""
        with self._lock:
            if self._closed:
                raise CubeError("unreachable", "This connection has been removed.")
            if self._session is not None and self._session.is_open:
                return self._session, False
            if self._endpoint is None:
                self._endpoint = resolve.endpoint(self._target, self._lookup)
            session = self._session_factory(self._endpoint[0], self._endpoint[1], self._username, self._password)
            self._pending = session
            try:
                session.open()
            except CubeError as exc:
                if exc.category == "unreachable" and self._target.instance:
                    # A named instance gets a new port when the service restarts.
                    self._endpoint = None
                raise
            finally:
                self._pending = None
            self._session = session
            return session, True

    def _call(self, request: Callable[[Session], object], *, epoch: int | None = None):
        """Run one request. `epoch` is given for queries: the cancel count at their start."""
        retried = False
        while True:
            fresh = True
            try:
                session, fresh = self._open_session()
                if epoch is not None and self._cancel_epoch != epoch:
                    raise CubeError("cancelled", "The query was cancelled.")
                return request(session)
            except CubeError as exc:
                if epoch is not None and self._cancel_epoch != epoch:
                    raise CubeError("cancelled", "The query was cancelled.") from exc
                # Repeat once when the connection we had been holding was dead, or when
                # a cancel aimed at a query shut it down under a metadata request.
                # Never after a timeout: the server may still be working on it.
                stale_connection = exc.retry and not fresh
                collateral_cancel = exc.category == "cancelled" and epoch is None
                if retried or not (stale_connection or collateral_cancel):
                    raise
                retried = True

    def cancel(self) -> None:
        with self._state:
            if self._active_queries == 0:
                return
            self._cancel_epoch += 1
        for session in (self._session, self._pending):
            if session is not None:
                session.abort()

    def close(self) -> None:
        self._closed = True
        # Abort first: the session lock is held for a whole round trip, and waiting
        # for a long query to finish is not what "remove this connection" means.
        for session in (self._session, self._pending):
            if session is not None:
                session.abort()
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
                _optional(lambda: session.discover("MDSCHEMA_LEVELS", restrictions, catalog)),
                _optional(lambda: session.discover("MDSCHEMA_KPIS", restrictions, catalog)),
                _optional(lambda: session.discover("MDSCHEMA_SETS", restrictions, catalog)),
            )

        measure_rows, dimension_rows, hierarchy_rows, level_rows, kpi_rows, set_rows = self._call(load)
        date_hierarchies = _date_hierarchies(level_rows)

        measures = []
        self._measure_formats[(catalog, cube)] = {
            r["MEASURE_UNIQUE_NAME"]: r.get("DEFAULT_FORMAT_STRING", "")
            for r in measure_rows
            if r.get("MEASURE_UNIQUE_NAME")
        }
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
            entry = {
                "uniqueName": unique,
                "caption": r.get("HIERARCHY_CAPTION") or unique,
                "folder": r.get("HIERARCHY_DISPLAY_FOLDER", ""),
                # Bit 1 marks a hierarchy the cube designer built from several levels.
                "kind": "user" if origin & 1 else "attribute",
                "allMember": all_member,
            }
            if unique in date_hierarchies:
                entry["isDate"] = True
            dimension["hierarchies"].append(entry)
        self._all_members[(catalog, cube)] = all_members

        shown = [d for d in dimensions.values() if d["hierarchies"]]
        for d in shown:
            d["hierarchies"].sort(key=lambda h: h["caption"].lower())
        shown.sort(key=lambda d: d["caption"].lower())
        return {
            "measures": measures,
            "dimensions": shown,
            "kpis": _kpis(kpi_rows),
            "sets": _sets(set_rows, {d["uniqueName"] for d in shown}, all_members),
        }

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

    def details(
        self, catalog: str, cube: str, layout: Layout, members: list[str], measure: str, max_rows: int
    ) -> dict:
        text = mdx.details_query(cube, layout, members, measure, max_rows)
        with self._state:
            self._active_queries += 1
            epoch = self._cancel_epoch
        try:
            parsed = self._call(lambda s: s.execute_rowset(text, catalog), epoch=epoch)
        finally:
            with self._state:
                self._active_queries -= 1
        columns = parsed.columns or list(dict.fromkeys(name for row in parsed.rows for name in row))
        rows = [[clean_value(v) for v in line] for line in xmla.typed_rows(parsed)]
        return {"columns": plain_column_names(columns), "rows": rows[:max_rows], "truncated": len(rows) > max_rows}

    def query(self, catalog: str, cube: str, layout: Layout) -> tuple[CellSet, str]:
        formats = None
        plain = [v.measure.lower() for v in layout.values if v.show_as == "normal" and not v.kpi_part]
        repeated = len(plain) != len(set(plain))
        if repeated or any(v.show_as in ("differenceFrom", "runningTotalIn") for v in layout.values):
            if (catalog, cube) not in self._measure_formats:
                self.fields(catalog, cube)
            formats = self._measure_formats.get((catalog, cube))
        text = mdx.build_query(cube, layout, formats)
        with self._state:
            self._active_queries += 1
            epoch = self._cancel_epoch
        try:
            return self._call(lambda s: s.execute(text, catalog), epoch=epoch), text
        finally:
            with self._state:
                self._active_queries -= 1
