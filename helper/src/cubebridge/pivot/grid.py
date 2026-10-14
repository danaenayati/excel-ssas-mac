"""Turn a query result into the `Grid` JSON the pane writes into the sheet.

The output shape is `Grid` in addin/src/types.ts. Everything the pane needs to lay
the block out is decided here, so the pane stays a dumb writer: label text, indent,
which labels are suppressed as repeats, which lines are totals.
"""

from __future__ import annotations

import math
import re

from ..errors import CubeError
from .formats import GENERAL, to_excel
from .mdx import custom_item_name, value_members
from .model import AxisMember, CellSet, Layout, PlacedField

DEFAULT_CELL_LIMIT = 300_000
MAX_SAFE_INT = 2**53

_LAST_PART = re.compile(r"\[((?:[^\]]|\]\])*)\]$")


def build_grid(
    cellset: CellSet,
    layout: Layout,
    *,
    elapsed_ms: int = 0,
    mdx: str = "",
    cell_limit: int = DEFAULT_CELL_LIMIT,
) -> dict:
    col_tuples = cellset.axes[0] if cellset.axes else []
    has_row_axis = len(cellset.axes) > 1
    row_tuples = cellset.axes[1] if has_row_axis else []
    all_cols = len(col_tuples)
    n_cols = all_cols
    n_rows = len(row_tuples) if has_row_axis else (1 if n_cols else 0)

    if n_rows * n_cols > cell_limit:
        raise CubeError(
            "too-large",
            f"This view would return {n_rows * n_cols:,} cells ({n_rows:,} rows × {n_cols:,} columns). "
            f"Add a filter or remove a field to bring it under {cell_limit:,}.",
        )

    # The tuples to show, by their place in the cellset. A subtotal with no item left
    # under it is dropped, as Excel shows an outer item only with items under it: with
    # "Include filtered items in totals" its total is not empty, so NON EMPTY keeps it.
    col_index = _shown(col_tuples, layout.columns)
    col_tuples = [col_tuples[i] for i in col_index]
    n_cols = len(col_tuples)
    row_index = _shown(row_tuples, layout.rows) if has_row_axis and layout.rows else list(range(n_rows))
    row_tuples = [row_tuples[i] for i in row_index] if has_row_axis and layout.rows else row_tuples
    n_rows = len(row_index)

    # "Include filtered items in totals" marks the totals with an asterisk, as Excel does.
    star = "*" if layout.options.include_filtered and _anything_narrowed(layout) else ""
    col_width = sum(f.width for f in layout.columns)
    col_lines, col_kinds = _axis_lines(col_tuples, layout.columns, _calculated_items(layout.columns), star)
    col_header_rows = [[line[i] for line in col_lines] for i in range(col_width)]
    col_header_rows.append(
        [_header(_measure_caption(tup, layout, col_width), 0, "measure") for tup in col_tuples]
    )

    # The value field each column shows. Read from the column's own measure member, not
    # from its position: NON EMPTY can drop single tuple × measure columns.
    by_name: dict[str, int] = {}
    for index, member in enumerate(value_members(layout)):
        by_name.setdefault(member.name.lower(), index)
    col_value_index = [
        by_name.get(tup[-1].unique_name.lower(), -1) if len(tup) > col_width else -1 for tup in col_tuples
    ]

    if has_row_axis and layout.rows:
        row_headers, row_kinds = _axis_lines(row_tuples, layout.rows, _calculated_items(layout.rows), star)
        row_captions = _field_captions(layout.rows)
        row_field_of = _field_of(layout.rows)
    else:
        row_headers = [[_header("Total" + star, 0, "grandTotal")] for _ in range(n_rows)]
        row_kinds = ["data"] * n_rows
        row_captions = [""]
        row_field_of = [0]

    palette = [GENERAL]
    palette_index = {GENERAL: 0}
    values, cell_formats = [], []
    for r in row_index:
        value_line, format_line = [], []
        base = r * all_cols
        for c in col_index:
            cell = cellset.cells.get(base + c)
            if cell is None:
                value_line.append(None)
                format_line.append(0)
                continue
            value_line.append(clean_value(cell.value))
            fmt = to_excel(cell.format_string)
            index = palette_index.get(fmt)
            if index is None:
                index = palette_index[fmt] = len(palette)
                palette.append(fmt)
            format_line.append(index)
        values.append(value_line)
        cell_formats.append(format_line)

    return {
        "filterRows": [{"caption": f.caption, "value": _filter_text(f)} for f in layout.filters],
        "rowFieldCaptions": row_captions,
        "colHeaderRows": col_header_rows,
        "rowHeaders": row_headers,
        "values": values,
        "formatPalette": palette,
        "cellFormats": cell_formats,
        "rowKinds": row_kinds,
        "colKinds": col_kinds,
        "colValueIndex": col_value_index,
        "rowFieldOf": row_field_of,
        "colFieldOf": _field_of(layout.columns),
        "stats": {"rows": n_rows, "columns": n_cols, "elapsedMs": int(elapsed_ms), "mdx": mdx},
    }


def _header(text: str, indent: int, kind: str, member: dict | None = None) -> dict:
    return {"text": text, "indent": indent, "kind": kind, "member": member}


def _member_ref(member: AxisMember, hierarchy: str, expanded: frozenset[str], calculated: bool) -> dict:
    ref = {
        "uniqueName": member.unique_name,
        "hierarchy": hierarchy,
        "caption": member.caption,
        "level": member.level,
        "hasChildren": member.child_count > 0 and not calculated,
        "expanded": member.unique_name in expanded,
    }
    if calculated:
        ref["calculated"] = True
    return ref


def _calculated_items(fields: tuple[PlacedField, ...]) -> frozenset[str]:
    """Unique names (lower case) of the groups and calculated items on one axis."""
    return frozenset(custom_item_name(f, item).lower() for f in fields for item in f.custom_items)


def _anything_narrowed(layout: Layout) -> bool:
    return any(f.narrowed for f in (*layout.rows, *layout.columns, *layout.filters, *layout.slicers))


def _field_of(fields: tuple[PlacedField, ...]) -> list[int]:
    """For each header column (or row) of an axis, the index of the field it belongs to."""
    return [i for i, f in enumerate(fields) for _ in range(f.width)]


def _field_captions(fields: tuple[PlacedField, ...]) -> list[str]:
    """One caption per header column: a set spanning several fields adds the others' names."""
    out = []
    for f in fields:
        out.append(f.caption)
        for hierarchy in f.hierarchies[1:]:
            match = _LAST_PART.search(hierarchy)
            out.append(match.group(1).replace("]]", "]") if match else hierarchy)
    return out


def _starts(fields: tuple[PlacedField, ...]) -> list[int]:
    """Where each field's first hierarchy is in a tuple of the axis."""
    starts, at = [], 0
    for f in fields:
        starts.append(at)
        at += f.width
    return starts


def _first_all(members: tuple[AxisMember, ...], fields: tuple[PlacedField, ...], starts: list[int]) -> int | None:
    """The outermost field at its All member: 0 for the grand total, more for a subtotal,
    None for an item's own line."""
    return next(
        (
            i for i, f in enumerate(fields)
            if f.all_members[0] and members[starts[i]].unique_name == f.all_members[0]
        ),
        None,
    )


def _shown(tuples: list[tuple[AxisMember, ...]], fields: tuple[PlacedField, ...]) -> list[int]:
    """The indexes of the tuples to show: all but a subtotal with no item's line that
    shares its outer items."""
    if len(fields) < 2:
        return list(range(len(tuples)))
    starts = _starts(fields)
    width = sum(f.width for f in fields)
    firsts = [_first_all(tup[:width], fields, starts) if len(tup) >= width else None for tup in tuples]
    # The outer items of every item's line, for each depth a subtotal can be at.
    under: dict[int, set[tuple[str, ...]]] = {}
    for depth in range(1, len(fields)):
        under[depth] = {
            tuple(m.unique_name for m in tup[: starts[depth]])
            for tup, first in zip(tuples, firsts)
            if first is None
        }
    return [
        i for i, (tup, first) in enumerate(zip(tuples, firsts))
        if not first or tuple(m.unique_name for m in tup[: starts[first]]) in under[first]
    ]


def _axis_lines(
    tuples: list[tuple[AxisMember, ...]],
    fields: tuple[PlacedField, ...],
    calculated: frozenset[str] = frozenset(),
    star: str = "",
) -> tuple[list[list[dict]], list[str]]:
    """One line of header cells per tuple (one cell per hierarchy), plus each line's kind.

    A field is usually one hierarchy, but a named set may span several: it then takes one
    column per hierarchy, and its total is the tuple of their All members.
    """
    columns = [(i, f, h, a) for i, f in enumerate(fields) for h, a in zip(f.hierarchies, f.all_members)]
    starts = _starts(fields)
    width = len(columns)
    expanded = [frozenset(f.expanded) for f in fields]
    lines: list[list[dict]] = []
    kinds: list[str] = []
    previous: tuple[AxisMember, ...] | None = None
    for tup in tuples:
        members = tup[:width]
        if len(members) < width:
            raise CubeError("server", "The cube returned a result that does not match the report layout.")
        first_all = _first_all(members, fields, starts)
        kind ="data" if first_all is None else ("grandTotal" if first_all == 0 else "subtotal")
        # With measures innermost on columns, the same hierarchy tuple repeats once per
        # measure; only its first occurrence is labelled.
        repeat_of_previous = previous is not None and _same(previous, members, width)

        line = []
        for v, (fi, placed, hierarchy, all_member) in enumerate(columns):
            mem = members[v]
            last_of_field = v == starts[fi] + placed.width - 1
            if kind == "grandTotal":
                line.append(_header("Grand Total" + star if v == 0 and not repeat_of_previous else "", 0, "grandTotal"))
            elif kind == "subtotal" and (fi > first_all - 1 or (fi == first_all - 1 and last_of_field)):
                label = f"{mem.caption} Total{star}" if fi == first_all - 1 and not repeat_of_previous else ""
                line.append(_header(label, 0, "subtotal"))
            else:
                suppressed = previous is not None and _same(previous, members, v + 1)
                if placed.level is not None and placed.width == 1:
                    base = placed.level.number
                else:
                    base = 1 if all_member else 0
                indent = max(0, mem.level - base)
                line.append(
                    _header(
                        "" if suppressed else mem.caption,
                        indent,
                        "blank" if suppressed else "member",
                        _member_ref(mem, hierarchy, expanded[fi], mem.unique_name.lower() in calculated),
                    )
                )
        lines.append(line)
        kinds.append(kind)
        previous = members
    return lines, kinds


def _same(a: tuple[AxisMember, ...], b: tuple[AxisMember, ...], count: int) -> bool:
    return all(a[j].unique_name == b[j].unique_name for j in range(count))


def _measure_caption(tup: tuple[AxisMember, ...], layout: Layout, col_width: int) -> str:
    if len(tup) > col_width:
        return tup[-1].caption
    return layout.values[0].caption if layout.values else ""


def _filter_text(placed: PlacedField) -> str:
    if not placed.members:
        # Without an All member the cube applies the hierarchy's own default.
        return "All" if placed.all_member else "(Default)"
    if len(placed.members) > 1:
        return "(Multiple Items)"
    if placed.member_captions:
        return placed.member_captions[0]
    match = _LAST_PART.search(placed.members[0])
    return match.group(1).replace("]]", "]") if match else placed.members[0]


def clean_value(value: object) -> object:
    """Make a cell value safe for JSON and for an Excel cell."""
    if isinstance(value, bool) or value is None or isinstance(value, str):
        return value
    if isinstance(value, float):
        return value if math.isfinite(value) else "#NUM!"
    if isinstance(value, int):
        return value if abs(value) <= MAX_SAFE_INT else float(value)
    return str(value)
