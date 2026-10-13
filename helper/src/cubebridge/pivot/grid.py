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
    n_cols = len(col_tuples)
    n_rows = len(row_tuples) if has_row_axis else (1 if n_cols else 0)

    if n_rows * n_cols > cell_limit:
        raise CubeError(
            "too-large",
            f"This view would return {n_rows * n_cols:,} cells ({n_rows:,} rows × {n_cols:,} columns). "
            f"Add a filter or remove a field to bring it under {cell_limit:,}.",
        )

    col_lines, col_kinds = _axis_lines(col_tuples, layout.columns)
    col_header_rows = [[line[i] for line in col_lines] for i in range(len(layout.columns))]
    col_header_rows.append(
        [_header(_measure_caption(tup, layout), 0, "measure") for tup in col_tuples]
    )

    if has_row_axis and layout.rows:
        row_headers, row_kinds = _axis_lines(row_tuples, layout.rows)
        row_captions = [f.caption for f in layout.rows]
    else:
        row_headers = [[_header("Total", 0, "grandTotal")] for _ in range(n_rows)]
        row_kinds = ["data"] * n_rows
        row_captions = [""]

    palette = [GENERAL]
    palette_index = {GENERAL: 0}
    values, cell_formats = [], []
    for r in range(n_rows):
        value_line, format_line = [], []
        base = r * n_cols
        for c in range(n_cols):
            cell = cellset.cells.get(base + c)
            if cell is None:
                value_line.append(None)
                format_line.append(0)
                continue
            value_line.append(_clean(cell.value))
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
        "stats": {"rows": n_rows, "columns": n_cols, "elapsedMs": int(elapsed_ms), "mdx": mdx},
    }


def _header(text: str, indent: int, kind: str, member: dict | None = None) -> dict:
    return {"text": text, "indent": indent, "kind": kind, "member": member}


def _member_ref(member: AxisMember, placed: PlacedField, expanded: frozenset[str]) -> dict:
    return {
        "uniqueName": member.unique_name,
        "hierarchy": placed.hierarchy,
        "caption": member.caption,
        "level": member.level,
        "hasChildren": member.child_count > 0,
        "expanded": member.unique_name in expanded,
    }


def _axis_lines(
    tuples: list[tuple[AxisMember, ...]], fields: tuple[PlacedField, ...]
) -> tuple[list[list[dict]], list[str]]:
    """One line of header cells per tuple (one cell per field), plus each line's kind."""
    expanded = [frozenset(f.expanded) for f in fields]
    lines: list[list[dict]] = []
    kinds: list[str] = []
    previous: tuple[AxisMember, ...] | None = None
    for tup in tuples:
        members = tup[: len(fields)]
        if len(members) < len(fields):
            raise CubeError("server", "The cube returned a result that does not match the report layout.")
        first_all = next(
            (i for i, (mem, f) in enumerate(zip(members, fields)) if f.all_member and mem.unique_name == f.all_member),
            None,
        )
        kind = "data" if first_all is None else ("grandTotal" if first_all == 0 else "subtotal")
        # With measures innermost on columns, the same hierarchy tuple repeats once per
        # measure; only its first occurrence is labelled.
        repeat_of_previous = previous is not None and _same(previous, members, len(members))

        line = []
        for i, (mem, placed) in enumerate(zip(members, fields)):
            if kind == "grandTotal":
                line.append(_header("Grand Total" if i == 0 and not repeat_of_previous else "", 0, "grandTotal"))
            elif kind == "subtotal" and i >= first_all - 1:
                label = f"{mem.caption} Total" if i == first_all - 1 and not repeat_of_previous else ""
                line.append(_header(label, 0, "subtotal"))
            else:
                suppressed = previous is not None and _same(previous, members, i + 1)
                indent = max(0, mem.level - (1 if placed.all_member else 0))
                line.append(
                    _header(
                        "" if suppressed else mem.caption,
                        indent,
                        "blank" if suppressed else "member",
                        _member_ref(mem, placed, expanded[i]),
                    )
                )
        lines.append(line)
        kinds.append(kind)
        previous = members
    return lines, kinds


def _same(a: tuple[AxisMember, ...], b: tuple[AxisMember, ...], count: int) -> bool:
    return all(a[j].unique_name == b[j].unique_name for j in range(count))


def _measure_caption(tup: tuple[AxisMember, ...], layout: Layout) -> str:
    if len(tup) > len(layout.columns):
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


def _clean(value: object) -> object:
    """Make a cell value safe for JSON and for an Excel cell."""
    if isinstance(value, bool) or value is None or isinstance(value, str):
        return value
    if isinstance(value, float):
        return value if math.isfinite(value) else "#NUM!"
    if isinstance(value, int):
        return value if abs(value) <= MAX_SAFE_INT else float(value)
    return str(value)
