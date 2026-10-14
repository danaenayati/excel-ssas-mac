"""A small built-in cube, so the add-in can be tried and tested with no server.

It answers the same `Layout` a real cube would and returns the same `CellSet` shape,
following the rules in pivot/mdx.py (detail members, drill-down, sorting, All-member
totals, selections and label/value filters as sub-selects with visual totals,
"Show Values As", NON EMPTY, and DRILLTHROUGH refusals). It does not parse MDX: it
evaluates the layout directly over a few thousand generated fact rows.
"""

from __future__ import annotations

import itertools
import random
from dataclasses import dataclass

from ..errors import CubeError
from ..pivot import mdx
from ..pivot.names import quote_identifier
from ..pivot.model import AxisMember, Cell, CellSet, LabelFilter, Layout, PlacedField, ValueField, ValueFilter

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
        text = mdx.build_query(cube, layout, {name: fmt for name, (_c, _g, fmt) in _MEASURES.items()})
        values = list(layout.values) or [ValueField(_DEFAULT_MEASURE, _MEASURES[_DEFAULT_MEASURE][0])]
        used = [v.measure for v in values] + [
            m
            for placed in (*layout.rows, *layout.columns)
            for m in (placed.sort.measure if placed.sort else None, placed.value_filter.measure if placed.value_filter else None)
            if m
        ]
        for name in used:
            if name not in _MEASURES:
                raise CubeError("server", "The cube rejected the request.", f"The measure {name} was not found in the cube.")

        facts = self._filtered_facts(layout)
        # Label and value filters, each evaluated in the space the earlier ones left,
        # as the nested sub-selects of the real query are. A value filter on a field
        # inside others keeps, for each combination of outer items, its own winners.
        allowed: dict[str, set[tuple]] = {}
        for fields in (layout.rows, layout.columns):
            for depth, placed in enumerate(fields):
                h = placed.hierarchy
                if placed.label_filter is not None:
                    keep = {
                        p for p in self._items(placed, allowed)
                        if _label_match(self._caption((h, p)), placed.label_filter)
                    }
                    allowed[h] = keep
                    facts = [f for f in facts if self._top_path(placed, f) in keep]
                if placed.value_filter is None:
                    continue
                outer = fields[:depth]
                groups: dict[tuple, list[dict]] = {}
                for fact in facts:
                    groups.setdefault(tuple(self._top_path(o, fact) for o in outer), []).append(fact)
                pairs = {
                    (key, p)
                    for key, group in groups.items()
                    for p in self._value_items(placed, group, allowed)
                }
                allowed[h] = {p for _, p in pairs}
                facts = [
                    f for f in facts
                    if (tuple(self._top_path(o, f) for o in outer), self._top_path(placed, f)) in pairs
                ]

        options = layout.options
        row_tuples = (
            self._axis_tuples(layout.rows, layout, facts, allowed, options.grand_total_row) if layout.rows else [()]
        )
        col_tuples = (
            self._axis_tuples(layout.columns, layout, facts, allowed, options.grand_total_column)
            if layout.columns
            else [()]
        )
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
                    _add(totals.setdefault((r, c), [0, 0.0, 0]), fact)

        if layout.options.non_empty:
            keep_rows = sorted({r for r, _ in totals})
            keep_cols = sorted({c for _, c in totals})
        else:
            keep_rows = list(range(len(row_tuples)))
            keep_cols = list(range(len(col_tuples)))

        cellset = CellSet()
        columns_axis = []
        for c in keep_cols:
            for value in values:
                if value.show_as == "normal":
                    measure = AxisMember(value.measure, _MEASURES[value.measure][0], "[Measures]", 0, 0)
                else:
                    caption = mdx.show_as_caption(value)
                    measure = AxisMember("[Measures]." + quote_identifier(caption), caption, "[Measures]", 0, 0)
                columns_axis.append(tuple(self._axis_member(k) for k in col_tuples[c]) + (measure,))
        cellset.axes.append(columns_axis)
        if layout.rows:
            cellset.axes.append([tuple(self._axis_member(k) for k in row_tuples[r]) for r in keep_rows])

        width = len(columns_axis)
        cache: dict[tuple, object] = {}
        for out_r, r in enumerate(keep_rows):
            for out_c, c in enumerate(keep_cols):
                slot = totals.get((r, c))
                if slot is None:
                    continue
                context = {h: path for h, path in (*row_tuples[r], *col_tuples[c])}
                for m_index, value in enumerate(values):
                    if value.show_as == "normal":
                        number, fmt = _measure_value(value.measure, slot), _MEASURES[value.measure][2]
                    else:
                        number = self._shown_as(value, layout, context, facts, cache)
                        fmt = _SHOWN_AS_FORMATS.get(value.show_as, "")
                        if value.show_as in ("differenceFrom", "runningTotalIn"):
                            fmt = _MEASURES[value.measure][2]
                    if number is not None:
                        ordinal = out_r * width + out_c * len(values) + m_index
                        cellset.cells[ordinal] = Cell(number, fmt, "")
        return cellset, text

    def _shown_as(self, value: ValueField, layout: Layout, context: dict, facts: list[dict], cache: dict):
        """One "Show Values As" cell, by Excel's definitions, as pivot/mdx.py asks a cube for it."""
        m = value.measure

        def at(changes: dict) -> object:
            where = dict(context)
            where.update(changes)
            key = (m, tuple(sorted((h, p) for h, p in where.items() if p)))
            if key not in cache:
                slot, found = [0, 0.0, 0], False
                for fact in facts:
                    if all(self._path_of(h, fact)[: len(p)] == p for h, p in key[1]):
                        _add(slot, fact)
                        found = True
                cache[key] = _measure_value(m, slot) if found else None
            return cache[key]

        def ratio(a: object, b: object) -> object:
            return None if a is None or not b else a / b

        here = at({})
        rows = [f.hierarchy for f in layout.rows]
        columns = [f.hierarchy for f in layout.columns]
        kind = value.show_as
        if kind == "percentOfGrandTotal":
            return ratio(here, at({h: () for h in rows + columns}))
        if kind == "percentOfColumnTotal":
            return ratio(here, at({h: () for h in rows}))
        if kind == "percentOfRowTotal":
            return ratio(here, at({h: () for h in columns}))
        if kind in ("percentOfParentRowTotal", "percentOfParentColumnTotal"):
            for h in reversed(rows if kind == "percentOfParentRowTotal" else columns):
                if context[h]:
                    return ratio(here, at({h: context[h][:-1]}))
            return 1.0
        if kind == "index":
            row_total, column_total = at({h: () for h in columns}), at({h: () for h in rows})
            everything = at({h: () for h in rows + columns})
            return None if here is None or not row_total or not column_total else here * everything / (row_total * column_total)

        axis = rows if value.base_field in rows else columns
        base = value.base_field
        current = context[base]
        if kind == "percentOfParentTotal":
            return ratio(here, at({h: () for h in axis[axis.index(base) + 1 :]}))
        if not current:  # the base field's total line
            return {"runningTotalIn": here, "percentRunningTotalIn": 1.0}.get(kind)
        level = self._level_paths(base, len(current))
        index = level.index(current)
        if kind in ("runningTotalIn", "percentRunningTotalIn"):
            numbers = [at({base: p}) for p in level[: index + 1]]
            running = None if all(n is None for n in numbers) else sum(n or 0 for n in numbers)
            return running if kind == "runningTotalIn" else ratio(running, at({base: ()}))
        if kind in ("rankAscending", "rankDescending"):
            if here is None:
                return None
            others = [n for n in (at({base: p}) for p in level) if n is not None]
            better = [n for n in others if (n < here if kind == "rankAscending" else n > here)]
            return len(better) + 1
        if value.base_item == "previous":
            item = level[index - 1] if index > 0 else None
        elif value.base_item == "next":
            item = level[index + 1] if index + 1 < len(level) else None
        else:
            item = self._by_name[value.base_item][1] if value.base_item in self._by_name else None
            if item == current:
                return None if kind != "percentOf" else ratio(here, here)
        other = at({base: item}) if item else None
        if kind == "percentOf":
            return ratio(here, other)
        if other is None:
            return None
        if kind == "differenceFrom":
            return (here or 0) - other
        return None if not other else ((here or 0) - other) / other  # percentDifferenceFrom

    def _path_of(self, hierarchy: str, fact: dict) -> tuple:
        return tuple(fact[attr] for _, attr in self._hierarchies[hierarchy].levels)

    def _level_paths(self, hierarchy: str, depth: int) -> list[tuple]:
        """Every member of one level, in the hierarchy's order (what PrevMember walks)."""
        out: list[tuple] = []

        def visit(path: tuple) -> None:
            for child in self._children.get((hierarchy, path), []):
                if len(child) == depth:
                    out.append(child)
                else:
                    visit(child)

        visit(())
        return out

    def details(
        self, catalog: str, cube: str, layout: Layout, members: list[str], measure: str, max_rows: int
    ) -> dict:
        self._check(catalog, cube)
        slicer = mdx.details_slicer(layout, members, measure)
        if measure not in _MEASURES:
            raise CubeError("server", "The cube rejected the request.", f"The measure {measure} was not found in the cube.")
        facts = self._facts
        for name in slicer[1:]:
            key = self._by_name.get(name)
            if key is None:
                raise CubeError("not-found", "That item is not in the sample cube.", name)
            h = self._hierarchies[key[0]]
            path = key[1]
            if path:
                facts = [f for f in facts if tuple(f[attr] for _, attr in h.levels)[: len(path)] == path]
        rows = [
            [f["year"], _MONTHS[f["month"] - 1], f["country"], f["city"], f["service"], f["segment"],
             f["rides"], f["revenue"], f["cancelled"]]
            for f in facts
        ]
        return {
            "columns": ["Year", "Month", "Country", "City", "Service Type", "Segment", "Rides", "Revenue", "Cancelled Rides"],
            "rows": rows[:max_rows],
            "truncated": len(rows) > max_rows,
        }

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

    def _detail(self, placed: PlacedField, facts: list[dict], allowed: set[tuple] | None) -> list[_Key]:
        h = self._hierarchy(placed.hierarchy)
        expanded = {self._by_name[name][1] for name in placed.expanded if name in self._by_name}
        selected = self._selection(placed)
        sort_key = self._sort_key(placed, facts)
        out: list[_Key] = []

        def visit(parent: tuple) -> None:
            children = [
                child
                for child in self._children.get((h.unique_name, parent), [])
                if self._in_space(child, selected) and (allowed is None or child[:1] in allowed)
            ]
            if sort_key is not None:
                children.sort(key=sort_key, reverse=placed.sort.order == "desc")
            for child in children:
                out.append((h.unique_name, child))
                if child in expanded:
                    visit(child)

        visit(())
        return out

    def _sort_key(self, placed: PlacedField, facts: list[dict]):
        """How siblings are ordered: by caption, or by a measure over the current facts."""
        if placed.sort is None:
            return None
        h = self._hierarchies[placed.hierarchy]
        if placed.sort.by == "caption":
            return lambda path: self._caption((h.unique_name, path)).lower()
        attrs = [attr for _, attr in h.levels]
        slots: dict[tuple, list[float]] = {}
        for fact in facts:
            values = tuple(fact[a] for a in attrs)
            for depth in range(1, len(values) + 1):
                _add(slots.setdefault(values[:depth], [0, 0.0, 0]), fact)
        measure = placed.sort.measure

        def key(path: tuple) -> float:
            number = _measure_value(measure, slots[path]) if path in slots else None
            return float("-inf") if number is None else number  # empty sorts lowest, as in MDX

        return key

    def _top_path(self, placed: PlacedField, fact: dict) -> tuple:
        h = self._hierarchies[placed.hierarchy]
        return (fact[h.levels[0][1]],)

    def _items(self, placed: PlacedField, allowed: dict[str, set[tuple]]) -> list[tuple]:
        """The field's top-level items left by its selection and any earlier filter."""
        h = self._hierarchy(placed.hierarchy)
        selected = self._selection(placed)
        earlier = allowed.get(placed.hierarchy)
        return [
            p
            for p in self._children.get((h.unique_name, ()), [])
            if self._in_space(p, selected) and (earlier is None or p in earlier)
        ]

    def _value_items(self, placed: PlacedField, facts: list[dict], allowed: dict[str, set[tuple]]) -> list[tuple]:
        """The items that pass the field's value filter, measured over `facts`."""
        items = self._items(placed, allowed)
        slots: dict[tuple, list[float]] = {}
        for fact in facts:
            _add(slots.setdefault(self._top_path(placed, fact), [0, 0.0, 0]), fact)
        vf = placed.value_filter
        numbers = {p: _measure_value(vf.measure, slots[p]) if p in slots else None for p in items}
        return _apply_value_filter(items, numbers, vf)

    def _axis_tuples(
        self,
        fields: tuple[PlacedField, ...],
        layout: Layout,
        facts: list[dict],
        allowed: dict[str, set[tuple]],
        grand_total: bool,
    ) -> list[tuple[_Key, ...]]:
        alls = [(f.hierarchy, ()) for f in fields]
        by_value = [bool(f.sort and f.sort.by == "measure") for f in fields]

        def build(depth: int, scope: list[dict]) -> list[tuple[_Key, ...]]:
            detail = self._detail(fields[depth], scope, allowed.get(fields[depth].hierarchy))
            if depth == len(fields) - 1:
                body = [(k,) for k in detail]
            elif any(by_value[depth + 1 :]):
                # Inner order by value is worked out within each outer item (Generate).
                body = []
                for k in detail:
                    inside = [f for f in scope if self._path_of(k[0], f)[: len(k[1])] == k[1]]
                    body.extend((k,) + rest for rest in build(depth + 1, inside))
            else:
                inner = build(depth + 1, scope)
                body = [(k,) + rest for k in detail for rest in inner]
            wanted = grand_total if depth == 0 else layout.options.subtotals
            if wanted:
                # The grand total always ends the axis; subtotals may head their group.
                if depth > 0 and layout.options.subtotals_at_top:
                    body.insert(0, tuple(alls[depth:]))
                else:
                    body.append(tuple(alls[depth:]))
            return body

        return build(0, facts)

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


_SHOWN_AS_FORMATS = {
    **dict.fromkeys(
        (
            "percentOfGrandTotal", "percentOfColumnTotal", "percentOfRowTotal", "percentOf",
            "percentOfParentRowTotal", "percentOfParentColumnTotal", "percentOfParentTotal",
            "percentDifferenceFrom", "percentRunningTotalIn",
        ),
        "Percent",
    ),
    "index": "0.00",
    "rankAscending": "#,##0",
    "rankDescending": "#,##0",
}


def _add(slot: list[float], fact: dict) -> None:
    slot[0] += fact["rides"]
    slot[1] += fact["revenue"]
    slot[2] += fact["cancelled"]


def _label_match(caption: str, label: LabelFilter) -> bool:
    """The same tests the MDX in pivot/mdx.py makes, without regard to case."""
    text, value, upper = caption.lower(), label.value.lower(), (label.value2 or "").lower()
    return {
        "equals": lambda: text == value,
        "notEquals": lambda: text != value,
        "beginsWith": lambda: text.startswith(value),
        "notBeginsWith": lambda: not text.startswith(value),
        "endsWith": lambda: text.endswith(value),
        "notEndsWith": lambda: not text.endswith(value),
        "contains": lambda: value in text,
        "notContains": lambda: value not in text,
        "greaterThan": lambda: text > value,
        "greaterOrEqual": lambda: text >= value,
        "lessThan": lambda: text < value,
        "lessOrEqual": lambda: text <= value,
        "between": lambda: value <= text <= upper,
        "notBetween": lambda: text < value or text > upper,
    }[label.operator]()


def _apply_value_filter(items: list[tuple], numbers: dict[tuple, object], vf: ValueFilter) -> list[tuple]:
    """MDX semantics: an empty value compares as 0; Bottom ignores empty items."""
    if vf.operator in ("top", "bottom"):
        if vf.operator == "top":
            ranked = sorted(items, key=lambda p: float("-inf") if numbers[p] is None else numbers[p], reverse=True)
        else:
            ranked = sorted((p for p in items if numbers[p] is not None), key=lambda p: numbers[p])
        if vf.top_by == "items":
            return ranked[: int(vf.value)]
        target = vf.value
        if vf.top_by == "percent":
            target = sum(numbers[p] or 0 for p in ranked) * vf.value / 100
        kept, running = [], 0.0
        for p in ranked:
            if kept and running >= target:
                break
            kept.append(p)
            running += numbers[p] or 0
        return kept
    lo, hi = vf.value, vf.value2 if vf.value2 is not None else vf.value
    test = {
        "equals": lambda v: v == lo,
        "notEquals": lambda v: v != lo,
        "greaterThan": lambda v: v > lo,
        "greaterOrEqual": lambda v: v >= lo,
        "lessThan": lambda v: v < lo,
        "lessOrEqual": lambda v: v <= lo,
        "between": lambda v: lo <= v <= hi,
        "notBetween": lambda v: v < lo or v > hi,
    }[vf.operator]
    return [p for p in items if test(numbers[p] if numbers[p] is not None else 0)]


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
