"""Turn a pivot layout into one MDX SELECT, and a value cell into a DRILLTHROUGH.

MDX works against both multidimensional and tabular models, which is why Windows
Excel uses it for PivotTables too. The shape follows what Excel generates:

- Each hierarchy on an axis contributes its top-level members ("detail"), drilled
  down where the user expanded a member, optionally sorted.
- Totals come from the hierarchies' All members, so the cube calculates them. A
  ratio or distinct count is therefore right on a total row, which it would not be
  if the totals were summed in the sheet.
- Item selections, label filters and value filters become sub-selects, so totals
  reflect only what is left ("visual totals", Excel's default for cube PivotTables).
  A single-member page filter goes in WHERE instead, which also gives calculated
  measures the right current member, except when a value filter needs it inside the
  sub-selects to be evaluated against the right numbers.
- "Show Values As" percentages are calculated members defined in the query.

Every name comes through `Layout.from_json`, which has already validated it.
"""

from __future__ import annotations

import re
from decimal import Decimal

from ..errors import CubeError
from .model import Layout, LabelFilter, PlacedField, ValueField, ValueFilter
from .names import hierarchy_of, quote_identifier, quote_string, validate_unique_name

CELL_PROPERTIES = "CELL PROPERTIES VALUE, FORMAT_STRING, FORMATTED_VALUE"
DEFAULT_MEASURE = "[Measures].DefaultMember"
MAX_DETAIL_MEMBERS = 64

SHOW_AS_LABELS = {
    "percentOfGrandTotal": "% of Grand Total",
    "percentOfColumnTotal": "% of Column Total",
    "percentOfRowTotal": "% of Row Total",
    "percentOf": "% Of",
    "percentOfParentRowTotal": "% of Parent Row Total",
    "percentOfParentColumnTotal": "% of Parent Column Total",
    "percentOfParentTotal": "% of Parent Total",
    "differenceFrom": "Difference From",
    "percentDifferenceFrom": "% Difference From",
    "runningTotalIn": "Running Total In",
    "percentRunningTotalIn": "% Running Total In",
    "rankAscending": "Rank Smallest to Largest",
    "rankDescending": "Rank Largest to Smallest",
    "index": "Index",
}
_PERCENT_KINDS = frozenset(
    {
        "percentOfGrandTotal", "percentOfColumnTotal", "percentOfRowTotal", "percentOf",
        "percentOfParentRowTotal", "percentOfParentColumnTotal", "percentOfParentTotal",
        "percentDifferenceFrom", "percentRunningTotalIn",
    }
)
_FORMATS = {"index": "0.00", "rankAscending": "#,##0", "rankDescending": "#,##0"}

_PART = re.compile(r"\[(?:[^\]]|\]\])*\]")


def build_query(cube: str, layout: Layout) -> str:
    non_empty = "NON EMPTY " if layout.options.non_empty else ""
    definitions, measure_names = _value_members(layout)
    measures = "{" + ", ".join(measure_names) + "}" if measure_names else "{" + DEFAULT_MEASURE + "}"

    options = layout.options
    if layout.columns:
        columns = f"CrossJoin({_axis_set(layout.columns, options.grand_total_column, options.subtotals)}, {measures})"
    else:
        columns = measures
    axes = [f"{non_empty}{columns} ON COLUMNS"]
    if layout.rows:
        axes.append(f"{non_empty}{_axis_set(layout.rows, options.grand_total_row, options.subtotals)} ON ROWS")

    on_axes = (*layout.rows, *layout.columns)
    value_filtered = any(placed.value_filter for placed in on_axes)
    source = quote_identifier(cube)
    slicer: list[str] = []
    for placed in on_axes:
        if placed.members:
            source = _subselect("(" + _set(placed.members) + ")", source)
    for placed in layout.filters:
        if len(placed.members) == 1 and not value_filtered:
            slicer.append(placed.members[0])
        elif placed.members:
            source = _subselect("(" + _set(placed.members) + ")", source)
        elif placed.all_member:
            # The sheet shows "All" for this filter. Say so explicitly, or the cube
            # would apply the hierarchy's default member, which need not be All.
            slicer.append(placed.all_member)
    for fields in (layout.rows, layout.columns):
        for depth, placed in enumerate(fields):
            for narrowed in _narrowing(fields, depth):
                source = _subselect("{" + narrowed + "}", source)

    with_clause = "WITH " + " ".join(definitions) + " " if definitions else ""
    where = f" WHERE ({', '.join(slicer)})" if slicer else ""
    return f"{with_clause}SELECT {', '.join(axes)} FROM {source}{where} {CELL_PROPERTIES}"


def members_query(
    cube: str,
    hierarchy: str,
    all_member: str | None,
    parent: str | None,
    search: str | None,
    limit: int,
) -> str:
    """A query whose ROWS axis lists members; it returns no cells."""
    hierarchy = validate_unique_name(hierarchy)
    if search:
        members = (
            f"Filter({{{hierarchy}.AllMembers}}, "
            f"InStr(1, LCase({hierarchy}.CurrentMember.Member_Caption), {quote_string(search.lower())}) > 0)"
        )
    elif parent:
        members = f"{{{validate_unique_name(parent)}.Children}}"
    else:
        members = _top_level(hierarchy, validate_unique_name(all_member) if all_member else None)
    return f"SELECT {{}} ON COLUMNS, Head({members}, {int(limit) + 1}) ON ROWS FROM {quote_identifier(cube)}"


def show_as_caption(value: ValueField) -> str:
    """The header caption of a value field shown another way, e.g. "Rides (% Of Calendar, 2025)"."""
    label = SHOW_AS_LABELS[value.show_as]
    if value.base_field:
        label += f" {_last_part(value.base_field)}"
    if value.base_item:
        item = value.base_item if value.base_item in ("previous", "next") else _last_part(value.base_item)
        label += f", {item}"
    return f"{value.caption or _last_part(value.measure)} ({label})"


def details_slicer(layout: Layout, members: list[str], measure: str) -> list[str]:
    """The WHERE tuple of "Show Details" for one cell, or a refusal.

    A DRILLTHROUGH can only be sliced, not sub-selected, so a field narrowed to several
    items cannot be honoured unless the cell itself pins it to one member. Excel on
    Windows refuses in the same situation.
    """
    measure = validate_unique_name(measure)
    if not measure.lower().startswith("[measures]."):
        raise CubeError("bad-request", "Show Details needs a measure.")
    if len(members) > MAX_DETAIL_MEMBERS:
        raise CubeError("bad-request", "Show Details was given too many members.")
    pinned: set[str] = set()
    slicer = [measure]
    for name in members:
        name = validate_unique_name(name)
        hierarchy = hierarchy_of(name)
        if hierarchy in pinned or hierarchy == "[measures]":
            raise CubeError("bad-request", "Show Details was given two members of one field.")
        pinned.add(hierarchy)
        slicer.append(name)
    for placed in layout.filters:
        if len(placed.members) == 1:
            slicer.append(placed.members[0])
            pinned.add(placed.hierarchy.lower())
    for placed in (*layout.rows, *layout.columns, *layout.filters):
        if placed.narrowed and placed.hierarchy.lower() not in pinned:
            raise CubeError(
                "bad-request",
                f"Show Details is not available here because '{placed.caption}' is filtered to several "
                "items. Pick a single item in that field, or clear its filter, and try again.",
            )
    return slicer


def details_query(cube: str, layout: Layout, members: list[str], measure: str, max_rows: int) -> str:
    slicer = details_slicer(layout, members, measure)
    # One row more than asked for tells the caller whether there were more.
    return (
        f"DRILLTHROUGH MAXROWS {int(max_rows) + 1} SELECT FROM {quote_identifier(cube)} "
        f"WHERE ({', '.join(slicer)})"
    )


# -- query pieces ------------------------------------------------------------------
def _set(names: tuple[str, ...] | list[str]) -> str:
    return "{" + ", ".join(names) + "}"


def _subselect(set_expression: str, source: str) -> str:
    return f"(SELECT {set_expression} ON COLUMNS FROM {source})"


def _top_level(hierarchy: str, all_member: str | None) -> str:
    if all_member:
        return f"{{{all_member}.Children}}"
    return f"{{{hierarchy}.Levels(0).Members}}"


def _detail(placed: PlacedField) -> str:
    detail = _top_level(placed.hierarchy, placed.all_member)
    if placed.expanded:
        detail = f"DrilldownMember({detail}, {_set(placed.expanded)}, RECURSIVE)"
    if placed.sort:
        key = (
            f"{placed.hierarchy}.CurrentMember.Member_Caption"
            if placed.sort.by == "caption"
            else placed.sort.measure
        )
        # ASC/DESC (not BASC/BDESC) keep each parent above its own children.
        detail = f"Order({detail}, {key}, {placed.sort.order.upper()})"
    return detail


def _axis_set(fields: tuple[PlacedField, ...], grand_total: bool, subtotals: bool) -> str:
    """Detail tuples followed by their totals, outermost total last.

    T(n) = {D(n), {All(n)}}
    T(k) = {CrossJoin(D(k), T(k+1)), {(All(k), …, All(n))}}

    The All tuple at depth 0 is the grand total; deeper ones are subtotals of the
    hierarchy just outside them. When a field further in sorts by a measure, the
    CrossJoin becomes a Generate, so that the inner order is worked out separately
    for each outer item, as in Excel.
    """
    count = len(fields)
    sorts_by_value = [bool(f.sort and f.sort.by == "measure") for f in fields]

    def all_tuple(start: int) -> str | None:
        alls = [f.all_member for f in fields[start:]]
        if any(a is None for a in alls):
            return None
        return alls[0] if len(alls) == 1 else "(" + ", ".join(alls) + ")"

    def build(depth: int) -> str:
        detail = _detail(fields[depth])
        if depth == count - 1:
            body = detail
        elif any(sorts_by_value[depth + 1 :]):
            body = f"Generate({detail}, CrossJoin({{{fields[depth].hierarchy}.CurrentMember}}, {build(depth + 1)}))"
        else:
            body = f"CrossJoin({detail}, {build(depth + 1)})"
        wanted = grand_total if depth == 0 else subtotals
        total = all_tuple(depth) if wanted else None
        return f"{{{body}, {{{total}}}}}" if total else body

    return build(0)


def _narrowing(fields: tuple[PlacedField, ...], depth: int) -> list[str]:
    """Sub-select sets for the field at `depth` on its axis: label filter, then value filter.

    A value filter on a field inside other fields is evaluated for each combination of
    the outer fields' items (the top 3 cities in each country), as Excel does.
    """
    placed = fields[depth]
    top = _top_level(placed.hierarchy, placed.all_member)
    sets = []
    if placed.label_filter is not None:
        sets.append(f"Filter({top}, {_label_condition(placed.hierarchy, placed.label_filter)})")
    if placed.value_filter is not None:
        if depth == 0:
            sets.append(_value_expression(top, placed.value_filter))
        else:
            outer = fields[:depth]
            outer_items = (
                _top_level(outer[0].hierarchy, outer[0].all_member)
                if len(outer) == 1
                else "CrossJoin(" + ", ".join(_top_level(f.hierarchy, f.all_member) for f in outer) + ")"
            )
            current = ", ".join(f"{f.hierarchy}.CurrentMember" for f in outer)
            current = "{" + (f"({current})" if len(outer) > 1 else current) + "}"
            per_outer = _value_expression(f"CrossJoin({current}, {top})", placed.value_filter)
            sets.append(f"Generate({outer_items}, {per_outer})")
    return sets


def _label_condition(hierarchy: str, label: LabelFilter) -> str:
    caption = f"LCase({hierarchy}.CurrentMember.Member_Caption)"
    value = quote_string(label.value.lower())
    op = label.operator
    if op in ("equals", "notEquals"):
        return f"{caption} {'=' if op == 'equals' else '<>'} {value}"
    if op in ("beginsWith", "notBeginsWith"):
        return f"InStr(1, {caption}, {value}) {'=' if op == 'beginsWith' else '<>'} 1"
    if op in ("endsWith", "notEndsWith"):
        return f"Right({caption}, {len(label.value)}) {'=' if op == 'endsWith' else '<>'} {value}"
    if op in ("contains", "notContains"):
        return f"InStr(1, {caption}, {value}) {'>' if op == 'contains' else '='} 0"
    comparisons = {"greaterThan": ">", "greaterOrEqual": ">=", "lessThan": "<", "lessOrEqual": "<="}
    if op in comparisons:
        return f"{caption} {comparisons[op]} {value}"
    upper = quote_string((label.value2 or "").lower())
    if op == "between":
        return f"({caption} >= {value} AND {caption} <= {upper})"
    return f"({caption} < {value} OR {caption} > {upper})"  # notBetween


def _value_expression(items: str, vf: ValueFilter) -> str:
    measure = vf.measure
    first = _number(vf.value)
    if vf.operator in ("top", "bottom"):
        name = {"items": "Count", "percent": "Percent", "sum": "Sum"}[vf.top_by]
        if vf.operator == "top":
            return f"Top{name}({items}, {first}, {measure})"
        # Empty cells sort first; Excel's "Bottom" means the smallest actual values.
        return f"Bottom{name}(NonEmpty({items}, {{{measure}}}), {first}, {measure})"
    comparisons = {
        "equals": "=", "notEquals": "<>", "greaterThan": ">",
        "greaterOrEqual": ">=", "lessThan": "<", "lessOrEqual": "<=",
    }
    if vf.operator in comparisons:
        return f"Filter({items}, {measure} {comparisons[vf.operator]} {first})"
    second = _number(vf.value2 if vf.value2 is not None else vf.value)
    if vf.operator == "between":
        return f"Filter({items}, {measure} >= {first} AND {measure} <= {second})"
    return f"Filter({items}, {measure} < {first} OR {measure} > {second})"  # notBetween


def _value_members(layout: Layout) -> tuple[list[str], list[str]]:
    """WITH MEMBER definitions for "Show Values As", and the measures to put on columns."""
    definitions: list[str] = []
    names: list[str] = []
    for value in layout.values:
        if value.show_as == "normal":
            names.append(value.measure)
            continue
        name = "[Measures]." + quote_identifier(show_as_caption(value))
        definition = f"MEMBER {name} AS {_shown_as(value, layout)}"
        if value.show_as in _PERCENT_KINDS:
            definition += f", FORMAT_STRING = {quote_string('Percent')}"
        elif value.show_as in _FORMATS:
            definition += f", FORMAT_STRING = {quote_string(_FORMATS[value.show_as])}"
        definitions.append(definition)
        names.append(name)
    return definitions, names


def _tuple(*parts: str) -> str:
    return "(" + ", ".join(parts) + ")"


def _ratio(numerator: str, denominator: str) -> str:
    return f"IIF({denominator} = 0, NULL, {numerator} / {denominator})"


def _is_all(placed: PlacedField) -> str | None:
    return f"{placed.hierarchy}.CurrentMember IS {placed.all_member}" if placed.all_member else None


def _shown_as(value: ValueField, layout: Layout) -> str:
    """The MDX for one "Show Values As" kind, Excel's definitions."""
    m = value.measure
    kind = value.show_as
    row_alls = [f.all_member for f in layout.rows if f.all_member]
    column_alls = [f.all_member for f in layout.columns if f.all_member]

    if kind == "percentOfGrandTotal":
        return _ratio(m, _tuple(m, *row_alls, *column_alls))
    if kind == "percentOfColumnTotal":
        return _ratio(m, _tuple(m, *row_alls))
    if kind == "percentOfRowTotal":
        return _ratio(m, _tuple(m, *column_alls))
    if kind in ("percentOfParentRowTotal", "percentOfParentColumnTotal"):
        # The innermost field not at its total decides whose parent divides.
        expression = "1"
        for placed in layout.rows if kind == "percentOfParentRowTotal" else layout.columns:
            ratio = _ratio(m, _tuple(m, f"{placed.hierarchy}.CurrentMember.Parent"))
            at_all = _is_all(placed)
            expression = f"IIF(NOT ({at_all}), {ratio}, {expression})" if at_all else ratio
        return expression
    if kind == "index":
        everything = _tuple(m, *row_alls, *column_alls)
        row_total, column_total = _tuple(m, *column_alls), _tuple(m, *row_alls)
        return (
            f"IIF({row_total} = 0 OR {column_total} = 0, NULL, "
            f"({m} * {everything}) / ({row_total} * {column_total}))"
        )

    axis = layout.rows if any(f.hierarchy.lower() == value.base_field.lower() for f in layout.rows) else layout.columns
    position = next(i for i, f in enumerate(axis) if f.hierarchy.lower() == value.base_field.lower())
    base = axis[position]
    at_all = _is_all(base)
    current = f"{base.hierarchy}.CurrentMember"

    if kind == "percentOfParentTotal":
        inside = [f.all_member for f in axis[position + 1 :] if f.all_member]
        return _ratio(m, _tuple(m, *inside))
    if kind in ("runningTotalIn", "percentRunningTotalIn"):
        running = f"Sum({{{current}.Level.Members.Item(0) : {current}}}, {m})"
        if kind == "runningTotalIn":
            return f"IIF({at_all}, {m}, {running})" if at_all else running
        if not base.all_member:
            return "NULL"
        return f"IIF({at_all}, 1, {_ratio(running, _tuple(m, base.all_member))})"
    if kind in ("rankAscending", "rankDescending"):
        order = "BASC" if kind == "rankAscending" else "BDESC"
        blank = " OR ".join(c for c in (at_all, f"IsEmpty({m})") if c)
        ranked = f"Order(NonEmpty({current}.Level.Members, {{{m}}}), {m}, {order})"
        return f"IIF({blank}, NULL, Rank({current}, {ranked}, {m}))"

    # percentOf, differenceFrom, percentDifferenceFrom: against a base item
    if value.base_item == "previous":
        item = f"{current}.PrevMember"
    elif value.base_item == "next":
        item = f"{current}.NextMember"
    else:
        item = value.base_item
    other = _tuple(m, item)
    blanks = [c for c in (at_all,) if c]
    if value.base_item not in ("previous", "next"):
        blanks.append(f"{current} IS {item}")
    if kind == "percentOf":
        ratio = _ratio(m, other)
        return f"IIF({at_all}, NULL, {ratio})" if at_all else ratio
    blanks.append(f"IsEmpty({other})")
    if kind == "differenceFrom":
        return f"IIF({' OR '.join(blanks)}, NULL, {m} - {other})"
    blanks.append(f"{other} = 0")
    return f"IIF({' OR '.join(blanks)}, NULL, ({m} - {other}) / {other})"


def _number(value: float) -> str:
    """An MDX numeric literal, never in exponent form."""
    if float(value).is_integer():
        return str(int(value))
    return format(Decimal(repr(float(value))), "f")


def _last_part(unique_name: str) -> str:
    parts = _PART.findall(unique_name)
    return parts[-1][1:-1].replace("]]", "]") if parts else unique_name
