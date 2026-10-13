"""Turn a pivot layout into one MDX SELECT.

MDX works against both multidimensional and tabular models, which is why Windows
Excel uses it for PivotTables too. The shape follows what Excel generates:

- Each hierarchy on an axis contributes its top-level members ("detail"), drilled
  down where the user expanded a member.
- Totals come from the hierarchies' All members, so the cube calculates them. A
  ratio or distinct count is therefore right on a total row, which it would not be
  if the totals were summed in the sheet.
- Member selections become sub-selects, so totals reflect only what is selected
  ("visual totals"). A single-member page filter goes in WHERE instead, which also
  gives calculated measures the right current member.

Every name comes through `Layout.from_json`, which has already validated it.
"""

from __future__ import annotations

from .model import Layout, PlacedField
from .names import quote_identifier, quote_string, validate_unique_name

CELL_PROPERTIES = "CELL PROPERTIES VALUE, FORMAT_STRING, FORMATTED_VALUE"
DEFAULT_MEASURE = "[Measures].DefaultMember"


def build_query(cube: str, layout: Layout) -> str:
    non_empty = "NON EMPTY " if layout.options.non_empty else ""
    measures = "{" + ", ".join(v.measure for v in layout.values) + "}" if layout.values else "{" + DEFAULT_MEASURE + "}"

    if layout.columns:
        columns = f"CrossJoin({_axis_set(layout.columns, layout)}, {measures})"
    else:
        columns = measures
    axes = [f"{non_empty}{columns} ON COLUMNS"]
    if layout.rows:
        axes.append(f"{non_empty}{_axis_set(layout.rows, layout)} ON ROWS")

    source = quote_identifier(cube)
    slicer: list[str] = []
    for placed in (*layout.rows, *layout.columns):
        if placed.members:
            source = _subselect(placed, source)
    for placed in layout.filters:
        if len(placed.members) == 1:
            slicer.append(placed.members[0])
        elif placed.members:
            source = _subselect(placed, source)

    where = f" WHERE ({', '.join(slicer)})" if slicer else ""
    return f"SELECT {', '.join(axes)} FROM {source}{where} {CELL_PROPERTIES}"


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


def _top_level(hierarchy: str, all_member: str | None) -> str:
    if all_member:
        return f"{{{all_member}.Children}}"
    return f"{{{hierarchy}.Levels(0).Members}}"


def _detail(placed: PlacedField) -> str:
    base = _top_level(placed.hierarchy, placed.all_member)
    if placed.expanded:
        return f"DrilldownMember({base}, {{{', '.join(placed.expanded)}}}, RECURSIVE)"
    return base


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


def _subselect(placed: PlacedField, source: str) -> str:
    return f"(SELECT ({{{', '.join(placed.members)}}}) ON COLUMNS FROM {source})"
