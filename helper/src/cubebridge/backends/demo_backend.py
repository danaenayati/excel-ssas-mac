"""A small built-in cube, so the add-in can be tried and tested with no server.

It answers the same `Layout` a real cube would and returns the same `CellSet` shape,
following the rules in pivot/mdx.py (detail members, drill-down, All-member totals,
selections as sub-selects, NON EMPTY). It does not parse MDX: it evaluates the layout
directly over a few thousand generated fact rows.
"""

from __future__ import annotations

import itertools
import random
from dataclasses import dataclass

from ..errors import CubeError
from ..pivot import mdx
from ..pivot.model import AxisMember, Cell, CellSet, Layout, PlacedField

CATALOG = "Sample"
CUBE = "Ride Sales"

_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
_CITIES = {
    "Iran": {"Tehran": 9.0, "Mashhad": 3.2, "Isfahan": 2.4, "Shiraz": 2.0, "Tabriz": 1.6},
    "Turkey": {"Istanbul": 6.5, "Ankara": 2.2},
    "UAE": {"Dubai": 3.0, "Abu Dhabi": 1.4},
}
_SERVICES = {"Economy": (1.0, 4.2), "Comfort": (0.45, 7.5), "Bike": (0.6, 1.9), "Delivery": (0.35, 3.1)}
_SEGMENTS = {"New": 0.22, "Returning": 0.6, "Business": 0.18}
_NO_BIKE = {"Dubai", "Abu Dhabi", "Ankara"}


@dataclass(frozen=True)
class _Hierarchy:
    unique_name: str
    caption: str
    dimension: str
    dimension_caption: str
    levels: tuple[tuple[str, str], ...]  # (level name, fact attribute)
    kind: str

    @property
    def all_member(self) -> str:
        return f"{self.unique_name}.[All]"


_HIERARCHIES = (
    _Hierarchy("[Date].[Calendar]", "Calendar", "[Date]", "Date", (("Year", "year"), ("Quarter", "quarter"), ("Month", "month")), "user"),
    _Hierarchy("[Geography].[Geography]", "Geography", "[Geography]", "Geography", (("Country", "country"), ("City", "city")), "user"),
    _Hierarchy("[Geography].[City]", "City", "[Geography]", "Geography", (("City", "city"),), "attribute"),
    _Hierarchy("[Service].[Service Type]", "Service Type", "[Service]", "Service", (("Service Type", "service"),), "attribute"),
    _Hierarchy("[Customer].[Segment]", "Segment", "[Customer]", "Customer", (("Segment", "segment"),), "attribute"),
)

# unique name -> (caption, group, format, kind)
_MEASURES = {
    "[Measures].[Rides]": ("Rides", "Rides", "#,##0"),
    "[Measures].[Cancelled Rides]": ("Cancelled Rides", "Rides", "#,##0"),
    "[Measures].[Cancel Rate]": ("Cancel Rate", "Rides", "Percent"),
    "[Measures].[Revenue]": ("Revenue", "Revenue", "Currency"),
    "[Measures].[Avg Fare]": ("Avg Fare", "Revenue", "Currency"),
}
_DEFAULT_MEASURE = "[Measures].[Rides]"

_Key = tuple[str, tuple]  # (hierarchy unique name, path); () is the All member


def _generate_facts() -> list[dict]:
    rng = random.Random(20261013)
    facts = []
    for year, month in [(2025, m) for m in range(1, 13)] + [(2026, m) for m in range(1, 10)]:
        season = 1.0 + 0.18 * (1 if month in (3, 9, 12) else 0) - 0.12 * (1 if month in (1, 7) else 0)
        growth = 1.0 + 0.02 * ((year - 2025) * 12 + month - 1)
        for country, cities in _CITIES.items():
            for city, size in cities.items():
                for service, (share, fare) in _SERVICES.items():
                    if service == "Bike" and city in _NO_BIKE:
                        continue
                    for segment, weight in _SEGMENTS.items():
                        if segment == "Business" and service == "Bike":
                            continue
                        rides = int(900 * size * share * weight * season * growth * rng.uniform(0.85, 1.15))
                        if rides <= 0:
                            continue
                        price = fare * (1.6 if country == "UAE" else 1.0) * (1.25 if segment == "Business" else 1.0)
                        facts.append(
                            {
                                "year": year,
                                "quarter": f"Q{(month - 1) // 3 + 1}",
                                "month": month,
                                "country": country,
                                "city": city,
                                "service": service,
                                "segment": segment,
                                "rides": rides,
                                "revenue": round(rides * price * rng.uniform(0.92, 1.08), 2),
                                "cancelled": int(rides * rng.uniform(0.03, 0.11)),
                            }
                        )
    return facts


class DemoBackend:
    def __init__(self) -> None:
        self._facts = _generate_facts()
        self._hierarchies = {h.unique_name: h for h in _HIERARCHIES}
        self._children: dict[_Key, list[tuple]] = {}
        self._by_name: dict[str, _Key] = {}
        self._name_of: dict[_Key, str] = {}
        for h in _HIERARCHIES:
            self._index(h)

    # -- metadata ----------------------------------------------------------------
    def catalogs(self) -> list[dict]:
        return [{"name": CATALOG, "description": "Built-in sample data. No server needed."}]

    def cubes(self, catalog: str) -> list[dict]:
        self._check(catalog)
        return [{"name": CUBE, "caption": CUBE}]

    def fields(self, catalog: str, cube: str) -> dict:
        self._check(catalog, cube)
        dimensions: dict[str, dict] = {}
        for h in _HIERARCHIES:
            dimension = dimensions.setdefault(
                h.dimension, {"uniqueName": h.dimension, "caption": h.dimension_caption, "hierarchies": []}
            )
            dimension["hierarchies"].append(
                {"uniqueName": h.unique_name, "caption": h.caption, "folder": "", "kind": h.kind, "allMember": h.all_member}
            )
        return {
            "measures": [
                {"uniqueName": name, "caption": caption, "group": group, "folder": ""}
                for name, (caption, group, _fmt) in _MEASURES.items()
            ],
            "dimensions": list(dimensions.values()),
        }

    def members(self, catalog: str, cube: str, hierarchy: str, parent: str | None, search: str | None, limit: int) -> dict:
        self._check(catalog, cube)
        h = self._hierarchy(hierarchy)
        if search:
            needle = search.lower()
            keys = [k for k in self._walk(h, ()) if needle in self._caption(k).lower()]
        else:
            path = ()
            if parent:
                key = self._by_name.get(parent)
                if key is None or key[0] != h.unique_name:
                    raise CubeError("not-found", "That member is not in the sample cube.")
                path = key[1]
            keys = [(h.unique_name, child) for child in self._children.get((h.unique_name, path), [])]
        return {
            "members": [
                {
                    "uniqueName": self._name_of[key],
                    "caption": self._caption(key),
                    "level": len(key[1]),
                    "hasChildren": bool(self._children.get(key)),
                }
                for key in keys[:limit]
            ],
            "truncated": len(keys) > limit,
        }

    # -- query -------------------------------------------------------------------
    def query(self, catalog: str, cube: str, layout: Layout) -> tuple[CellSet, str]:
        self._check(catalog, cube)
        text = mdx.build_query(cube, layout)
        facts = self._filtered_facts(layout)
        row_tuples = self._axis_tuples(layout.rows, layout) if layout.rows else [()]
        col_tuples = self._axis_tuples(layout.columns, layout) if layout.columns else [()]
        row_index = {t: i for i, t in enumerate(row_tuples)}
        col_index = {t: i for i, t in enumerate(col_tuples)}
        row_shown = [{t[i] for t in row_tuples} for i in range(len(layout.rows))]
        col_shown = [{t[i] for t in col_tuples} for i in range(len(layout.columns))]

        totals: dict[tuple[int, int], list[float]] = {}
        for fact in facts:
            rows_hit = [row_index[t] for t in self._matching(fact, layout.rows, row_shown) if t in row_index]
            cols_hit = [col_index[t] for t in self._matching(fact, layout.columns, col_shown) if t in col_index]
            for r in rows_hit:
                for c in cols_hit:
                    slot = totals.setdefault((r, c), [0, 0.0, 0])
                    slot[0] += fact["rides"]
                    slot[1] += fact["revenue"]
                    slot[2] += fact["cancelled"]

        if layout.options.non_empty:
            keep_rows = sorted({r for r, _ in totals})
            keep_cols = sorted({c for _, c in totals})
        else:
            keep_rows = list(range(len(row_tuples)))
            keep_cols = list(range(len(col_tuples)))

        measures = [v.measure for v in layout.values] or [_DEFAULT_MEASURE]
        for name in measures:
            if name not in _MEASURES:
                raise CubeError("server", "The cube rejected the request.", f"The measure {name} was not found in the cube.")

        cellset = CellSet()
        columns_axis = []
        for c in keep_cols:
            for name in measures:
                columns_axis.append(
                    tuple(self._axis_member(k) for k in col_tuples[c])
                    + (AxisMember(name, _MEASURES[name][0], "[Measures]", 0, 0),)
                )
        cellset.axes.append(columns_axis)
        if layout.rows:
            cellset.axes.append([tuple(self._axis_member(k) for k in row_tuples[r]) for r in keep_rows])

        width = len(columns_axis)
        for out_r, r in enumerate(keep_rows):
            for out_c, c in enumerate(keep_cols):
                slot = totals.get((r, c))
                if slot is None:
                    continue
                for m_index, name in enumerate(measures):
                    value = _measure_value(name, slot)
                    if value is not None:
                        ordinal = out_r * width + out_c * len(measures) + m_index
                        cellset.cells[ordinal] = Cell(value, _MEASURES[name][2], "")
        return cellset, text

    def cancel(self) -> None:
        pass

    def close(self) -> None:
        pass

    # -- internals ---------------------------------------------------------------
    def _check(self, catalog: str, cube: str | None = None) -> None:
        if catalog != CATALOG:
            raise CubeError("not-found", f"The sample connection has no database called '{catalog}'.")
        if cube is not None and cube != CUBE:
            raise CubeError("not-found", f"The sample database has no cube called '{cube}'.")

    def _hierarchy(self, unique_name: str) -> _Hierarchy:
        h = self._hierarchies.get(unique_name)
        if h is None:
            raise CubeError("not-found", "That field is not in the sample cube.", unique_name)
        return h

    def _index(self, h: _Hierarchy) -> None:
        paths = {()}
        for fact in self._facts:
            values = tuple(fact[attr] for _, attr in h.levels)
            for depth in range(1, len(values) + 1):
                paths.add(values[:depth])
        for path in paths:
            key = (h.unique_name, path)
            if path:
                level_name = h.levels[len(path) - 1][0]
                name = f"{h.unique_name}.[{level_name}]." + "".join(f"&[{v}]" for v in path)
                self._children.setdefault((h.unique_name, path[:-1]), []).append(path)
            else:
                name = h.all_member
            self._by_name[name] = key
            self._name_of[key] = name
        for siblings in self._children.values():
            siblings.sort(key=lambda path: _sort_key(path[-1]))

    def _caption(self, key: _Key) -> str:
        path = key[1]
        if not path:
            return "All"
        h = self._hierarchies[key[0]]
        if h.levels[len(path) - 1][1] == "month":
            return _MONTHS[path[-1] - 1]
        return str(path[-1])

    def _walk(self, h: _Hierarchy, path: tuple):
        for child in self._children.get((h.unique_name, path), []):
            yield (h.unique_name, child)
            yield from self._walk(h, child)

    def _axis_member(self, key: _Key) -> AxisMember:
        return AxisMember(
            unique_name=self._name_of[key],
            caption=self._caption(key),
            hierarchy=key[0],
            level=len(key[1]),
            child_count=len(self._children.get(key, [])),
        )

    def _selection(self, placed: PlacedField) -> list[tuple] | None:
        """Selected member paths for a field, or None when nothing is selected."""
        if not placed.members:
            return None
        self._hierarchy(placed.hierarchy)
        return [key[1] for key in map(self._by_name.get, placed.members) if key and key[0] == placed.hierarchy]

    def _filtered_facts(self, layout: Layout) -> list[dict]:
        facts = self._facts
        for placed in (*layout.rows, *layout.columns, *layout.filters):
            selected = self._selection(placed)
            if selected is None:
                continue
            h = self._hierarchies[placed.hierarchy]
            attrs = [attr for _, attr in h.levels]
            facts = [
                f for f in facts if any(tuple(f[a] for a in attrs[: len(path)]) == path for path in selected)
            ]
        return facts

    def _in_space(self, path: tuple, selected: list[tuple] | None) -> bool:
        if selected is None:
            return True
        return any(path[: len(s)] == s or s[: len(path)] == path for s in selected)

    def _detail(self, placed: PlacedField) -> list[_Key]:
        h = self._hierarchy(placed.hierarchy)
        expanded = {self._by_name[name][1] for name in placed.expanded if name in self._by_name}
        selected = self._selection(placed)
        out: list[_Key] = []

        def visit(parent: tuple) -> None:
            for child in self._children.get((h.unique_name, parent), []):
                if not self._in_space(child, selected):
                    continue
                out.append((h.unique_name, child))
                if child in expanded:
                    visit(child)

        visit(())
        return out

    def _axis_tuples(self, fields: tuple[PlacedField, ...], layout: Layout) -> list[tuple[_Key, ...]]:
        details = [self._detail(f) for f in fields]
        alls = [(f.hierarchy, ()) for f in fields]

        def build(depth: int) -> list[tuple[_Key, ...]]:
            if depth == len(fields) - 1:
                body = [(k,) for k in details[depth]]
            else:
                inner = build(depth + 1)
                body = [(k,) + rest for k in details[depth] for rest in inner]
            wanted = layout.options.grand_totals if depth == 0 else layout.options.subtotals
            if wanted:
                body.append(tuple(alls[depth:]))
            return body

        return build(0)

    def _matching(self, fact: dict, fields: tuple[PlacedField, ...], shown: list[set[_Key]]):
        """Every axis tuple this fact contributes to: itself, its ancestors, and totals."""
        if not fields:
            return [()]
        per_field = []
        for placed, visible in zip(fields, shown):
            h = self._hierarchies[placed.hierarchy]
            values = tuple(fact[attr] for _, attr in h.levels)
            candidates = [(h.unique_name, values[:depth]) for depth in range(len(values) + 1)]
            per_field.append([k for k in candidates if k in visible])
        return itertools.product(*per_field)


def _sort_key(value: object) -> tuple:
    """Numbers ascending, services in their defined order, everything else by name."""
    if isinstance(value, int):
        return (0, value, "")
    if value in _SERVICES:
        return (0, list(_SERVICES).index(value), "")
    return (1, 0, str(value))


def _measure_value(name: str, slot: list[float]) -> object:
    rides, revenue, cancelled = slot
    if name == "[Measures].[Rides]":
        return int(rides)
    if name == "[Measures].[Cancelled Rides]":
        return int(cancelled)
    if name == "[Measures].[Revenue]":
        return round(revenue, 2)
    if name == "[Measures].[Avg Fare]":
        return revenue / rides if rides else None
    if name == "[Measures].[Cancel Rate]":
        return cancelled / (rides + cancelled) if rides + cancelled else None
    return None
