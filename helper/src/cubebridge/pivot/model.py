"""Data shapes shared by the SSAS layer, the pivot code and the backends.

`Layout` mirrors `Layout` in addin/src/types.ts. `CellSet` is what any backend returns
for a query: axis 0 is columns, axis 1 (if present) is rows, and cells are sparse,
keyed by ordinal `column + row * column_count`.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from ..errors import CubeError
from .names import validate_unique_name

MAX_FIELDS_PER_AREA = 16
MAX_MEMBERS_PER_FIELD = 5000
MAX_CAPTION = 400


@dataclass(frozen=True)
class AxisMember:
    unique_name: str
    caption: str
    hierarchy: str
    level: int
    child_count: int = 0


@dataclass(frozen=True)
class Cell:
    value: object = None
    format_string: str = ""
    formatted: str = ""


@dataclass
class CellSet:
    axes: list[list[tuple[AxisMember, ...]]] = field(default_factory=list)
    cells: dict[int, Cell] = field(default_factory=dict)


@dataclass(frozen=True)
class PlacedField:
    hierarchy: str
    caption: str
    all_member: str | None = None
    expanded: tuple[str, ...] = ()
    members: tuple[str, ...] = ()
    member_captions: tuple[str, ...] = ()


@dataclass(frozen=True)
class ValueField:
    measure: str
    caption: str


@dataclass(frozen=True)
class Options:
    grand_totals: bool = True
    subtotals: bool = False
    non_empty: bool = True


@dataclass(frozen=True)
class Layout:
    rows: tuple[PlacedField, ...] = ()
    columns: tuple[PlacedField, ...] = ()
    values: tuple[ValueField, ...] = ()
    filters: tuple[PlacedField, ...] = ()
    options: Options = Options()

    @classmethod
    def from_json(cls, data: object) -> Layout:
        """Validate untrusted JSON (it may come from a workbook someone else made)."""
        if not isinstance(data, dict):
            raise _bad("The report layout is not an object.")
        rows = _placed_list(data.get("rows", []), "rows")
        columns = _placed_list(data.get("columns", []), "columns")
        filters = _placed_list(data.get("filters", []), "filters")
        values = _value_list(data.get("values", []))

        seen: set[str] = set()
        for placed in (*rows, *columns, *filters):
            key = placed.hierarchy.lower()
            if key in seen:
                raise _bad(f"'{placed.caption}' is used in more than one area.")
            seen.add(key)

        raw_options = data.get("options", {})
        if not isinstance(raw_options, dict):
            raise _bad("The report options are not an object.")
        options = Options(
            grand_totals=bool(raw_options.get("grandTotals", True)),
            subtotals=bool(raw_options.get("subtotals", False)),
            non_empty=bool(raw_options.get("nonEmpty", True)),
        )
        return cls(rows=rows, columns=columns, values=values, filters=filters, options=options)


def _bad(message: str) -> CubeError:
    return CubeError("bad-request", message)


def _text(value: object, what: str) -> str:
    if not isinstance(value, str):
        raise _bad(f"{what} must be text.")
    if len(value) > MAX_CAPTION:
        raise _bad(f"{what} is too long.")
    return value


def _name_list(value: object, what: str) -> tuple[str, ...]:
    if not isinstance(value, list):
        raise _bad(f"{what} must be a list.")
    if len(value) > MAX_MEMBERS_PER_FIELD:
        raise _bad(f"{what} has more than {MAX_MEMBERS_PER_FIELD} entries.")
    return tuple(validate_unique_name(item) for item in value)


def _placed_list(value: object, area: str) -> tuple[PlacedField, ...]:
    if not isinstance(value, list):
        raise _bad(f"The {area} area must be a list.")
    if len(value) > MAX_FIELDS_PER_AREA:
        raise _bad(f"The {area} area has more than {MAX_FIELDS_PER_AREA} fields.")
    out = []
    for item in value:
        if not isinstance(item, dict):
            raise _bad(f"A field in {area} is not an object.")
        hierarchy = validate_unique_name(item.get("hierarchy"))
        if hierarchy.lower() == "[measures]":
            raise _bad("Measures belong in the Values area.")
        all_member = item.get("allMember")
        members = _name_list(item.get("members", []), "members")
        captions = item.get("memberCaptions", [])
        if not isinstance(captions, list):
            raise _bad("memberCaptions must be a list.")
        out.append(
            PlacedField(
                hierarchy=hierarchy,
                caption=_text(item.get("caption", ""), "caption"),
                all_member=validate_unique_name(all_member) if all_member else None,
                expanded=_name_list(item.get("expanded", []), "expanded"),
                members=members,
                member_captions=tuple(_text(c, "member caption") for c in captions[: len(members)]),
            )
        )
    return tuple(out)


def _value_list(value: object) -> tuple[ValueField, ...]:
    if not isinstance(value, list):
        raise _bad("The values area must be a list.")
    if len(value) > 64:
        raise _bad("The values area has more than 64 measures.")
    out = []
    seen: set[str] = set()
    for item in value:
        if not isinstance(item, dict):
            raise _bad("A value field is not an object.")
        measure = validate_unique_name(item.get("measure"))
        if not measure.lower().startswith("[measures]."):
            raise _bad("Only measures can be placed in the Values area.")
        if measure.lower() in seen:
            raise _bad("A measure is used twice in the Values area.")
        seen.add(measure.lower())
        out.append(ValueField(measure=measure, caption=_text(item.get("caption", ""), "caption")))
    return tuple(out)
