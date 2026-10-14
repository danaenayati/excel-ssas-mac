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
from .names import quote_identifier, quote_string, validate_unique_name

CELL_PROPERTIES = "CELL PROPERTIES VALUE, FORMAT_STRING, FORMATTED_VALUE"
DEFAULT_MEASURE = "[Measures].DefaultMember"
MAX_DETAIL_MEMBERS = 64

SHOW_AS_LABELS = {
    "percentOfGrandTotal": "% of Grand Total",
    "percentOfColumnTotal": "% of Column Total",
    "percentOfRowTotal": "% of Row Total",
}

_PART = re.compile(r"\[(?:[^\]]|\]\])*\]")


def build_query(cube: str, layout: Layout) -> str:
    non_empty = "NON EMPTY " if layout.options.non_empty else ""
    definitions, measure_names = _value_members(layout)
    measures = "{" + ", ".join(measure_names) + "}" if measure_names else "{" + DEFAULT_MEASURE + "}"

    if layout.columns:
        columns = f"CrossJoin({_axis_set(layout.columns, layout)}, {measures})"
    else:
        columns = measures
    axes = [f"{non_empty}{columns} ON COLUMNS"]
    if layout.rows:
        axes.append(f"{non_empty}{_axis_set(layout.rows, layout)} ON ROWS")

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
    for placed in on_axes:
        narrowed = _narrowing(placed)
        if narrowed:
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
    """The header caption of a value field shown as a percentage."""
    return f"{value.caption or _last_part(value.measure)} ({SHOW_AS_LABELS[value.show_as]})"


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


def hierarchy_of(unique_name: str) -> str:
    """The hierarchy part of a member's unique name, lower-cased for comparison."""
    parts = _PART.findall(unique_name)
    return ".".join(parts[:2]).lower() if len(parts) >= 2 else unique_name.lower()


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


def _axis_set(fields: tuple[PlacedField, ...], layout: Layout) -> str:
    """Detail tuples followed by their totals, outermost total last.

    T(n) = {D(n), {All(n)}}
    T(k) = {CrossJoin(D(k), T(k+1)), {(All(k), …, All(n))}}

    The All tuple at depth 0 is the grand total; deeper ones are subtotals of the
    hierarchy just outside them.
    """
    count = len(fields)

    def all_tuple(start: int) -> str | None:
        alls = [f.all_member for f in fields[start:]]
        if any(a is None for a in alls):
            return None
        return alls[0] if len(alls) == 1 else "(" + ", ".join(alls) + ")"

    def build(depth: int) -> str:
        detail = _detail(fields[depth])
        body = detail if depth == count - 1 else f"CrossJoin({detail}, {build(depth + 1)})"
        wanted = layout.options.grand_totals if depth == 0 else layout.options.subtotals
        total = all_tuple(depth) if wanted else None
        return f"{{{body}, {{{total}}}}}" if total else body

    return build(0)


def _narrowing(placed: PlacedField) -> str | None:
    """The field's top-level items after its label and value filters, or None."""
    if placed.label_filter is None and placed.value_filter is None:
        return None
    expression = _top_level(placed.hierarchy, placed.all_member)
    if placed.label_filter is not None:
        expression = f"Filter({expression}, {_label_condition(placed.hierarchy, placed.label_filter)})"
    if placed.value_filter is not None:
        expression = _value_expression(expression, placed.value_filter)
    return expression


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
    row_alls = [f.all_member for f in layout.rows if f.all_member]
    column_alls = [f.all_member for f in layout.columns if f.all_member]
    for value in layout.values:
        if value.show_as == "normal":
            names.append(value.measure)
            continue
        alls = {
            "percentOfGrandTotal": row_alls + column_alls,
            "percentOfColumnTotal": row_alls,
            "percentOfRowTotal": column_alls,
        }[value.show_as]
        total = "(" + ", ".join([value.measure, *alls]) + ")"
        name = "[Measures]." + quote_identifier(show_as_caption(value))
        definitions.append(
            f"MEMBER {name} AS IIF({total} = 0, NULL, {value.measure} / {total}), "
            f"FORMAT_STRING = {quote_string('Percent')}"
        )
        names.append(name)
    return definitions, names


def _number(value: float) -> str:
    """An MDX numeric literal, never in exponent form."""
    if float(value).is_integer():
        return str(int(value))
    return format(Decimal(repr(float(value))), "f")


def _last_part(unique_name: str) -> str:
    parts = _PART.findall(unique_name)
    return parts[-1][1:-1].replace("]]", "]") if parts else unique_name
