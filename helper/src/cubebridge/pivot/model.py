"""Data shapes shared by the SSAS layer, the pivot code and the backends.

`Layout` mirrors `Layout` in addin/src/types.ts. `CellSet` is what any backend returns
for a query: axis 0 is columns, axis 1 (if present) is rows, and cells are sparse,
keyed by ordinal `column + row * column_count`.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field

from ..errors import CubeError
from .names import hierarchy_of, validate_unique_name

MAX_FIELDS_PER_AREA = 16
MAX_MEMBERS_PER_FIELD = 5000
MAX_CAPTION = 400
MAX_FILTER_TEXT = 200
MAX_TOP_ITEMS = 100_000

LABEL_OPERATORS = frozenset(
    {
        "equals", "notEquals", "beginsWith", "notBeginsWith", "endsWith", "notEndsWith",
        "contains", "notContains", "greaterThan", "greaterOrEqual", "lessThan", "lessOrEqual",
        "between", "notBetween",
    }
)
VALUE_OPERATORS = frozenset(
    {
        "equals", "notEquals", "greaterThan", "greaterOrEqual", "lessThan", "lessOrEqual",
        "between", "notBetween", "top", "bottom",
    }
)
RANGE_OPERATORS = frozenset({"between", "notBetween"})
TOP_BY = frozenset({"items", "percent", "sum"})
SHOW_AS = frozenset(
    {
        "normal", "percentOfGrandTotal", "percentOfColumnTotal", "percentOfRowTotal", "percentOf",
        "percentOfParentRowTotal", "percentOfParentColumnTotal", "percentOfParentTotal",
        "differenceFrom", "percentDifferenceFrom", "runningTotalIn", "percentRunningTotalIn",
        "rankAscending", "rankDescending", "index",
    }
)
NEEDS_BASE_FIELD = frozenset(
    {
        "percentOf", "percentOfParentTotal", "differenceFrom", "percentDifferenceFrom",
        "runningTotalIn", "percentRunningTotalIn", "rankAscending", "rankDescending",
    }
)
NEEDS_BASE_ITEM = frozenset({"percentOf", "differenceFrom", "percentDifferenceFrom"})
RELATIVE_ITEMS = frozenset({"previous", "next"})


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
class FieldSort:
    by: str  # "caption" | "measure"
    order: str  # "asc" | "desc"
    measure: str | None = None


@dataclass(frozen=True)
class LabelFilter:
    operator: str
    value: str
    value2: str | None = None


@dataclass(frozen=True)
class ValueFilter:
    operator: str
    measure: str
    value: float
    value2: float | None = None
    top_by: str = "items"


@dataclass(frozen=True)
class PlacedField:
    hierarchy: str
    caption: str
    all_member: str | None = None
    expanded: tuple[str, ...] = ()
    members: tuple[str, ...] = ()
    member_captions: tuple[str, ...] = ()
    sort: FieldSort | None = None
    label_filter: LabelFilter | None = None
    value_filter: ValueFilter | None = None

    @property
    def narrowed(self) -> bool:
        """True when anything restricts this field to fewer than all of its items."""
        return bool(self.members) or self.label_filter is not None or self.value_filter is not None


@dataclass(frozen=True)
class ValueField:
    measure: str
    caption: str
    show_as: str = "normal"
    base_field: str | None = None  # hierarchy unique name, on rows or columns
    base_item: str | None = None  # "previous", "next" or a member unique name


@dataclass(frozen=True)
class Options:
    grand_totals: bool = True
    subtotals: bool = False
    non_empty: bool = True
    grand_total_row: bool = True  # the bottom row: the rows axis' All tuple
    grand_total_column: bool = True  # the right-hand column: the columns axis' All tuple


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
        grand_totals = bool(raw_options.get("grandTotals", True))
        options = Options(
            grand_totals=grand_totals,
            subtotals=bool(raw_options.get("subtotals", False)),
            non_empty=bool(raw_options.get("nonEmpty", True)),
            grand_total_row=bool(raw_options.get("grandTotalRow", grand_totals)),
            grand_total_column=bool(raw_options.get("grandTotalColumn", grand_totals)),
        )
        on_axes = {placed.hierarchy.lower() for placed in (*rows, *columns)}
        for value in values:
            if value.base_field and value.base_field.lower() not in on_axes:
                raise _bad(f"The base field of '{value.caption}' must be in Rows or Columns.")
            if value.base_item and value.base_item not in RELATIVE_ITEMS:
                if hierarchy_of(value.base_item) != value.base_field.lower():
                    raise _bad(f"The base item of '{value.caption}' is not an item of its base field.")
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


def _measure_name(value: object, what: str) -> str:
    name = validate_unique_name(value)
    if not name.lower().startswith("[measures]."):
        raise _bad(f"{what} must be a measure.")
    return name


def _number(value: object, what: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        raise _bad(f"{what} must be a number.")
    return float(value)


def _filter_text(value: object, what: str) -> str:
    text = _text(value, what)
    if len(text) > MAX_FILTER_TEXT or any(ord(ch) < 0x20 for ch in text):
        raise _bad(f"{what} is too long or contains characters that are not allowed.")
    return text


def _sort(value: object) -> FieldSort | None:
    if value is None:
        return None
    if not isinstance(value, dict):
        raise _bad("A sort setting is not an object.")
    by, order = value.get("by"), value.get("order")
    if by not in ("caption", "measure") or order not in ("asc", "desc"):
        raise _bad("A sort setting is not valid.")
    measure = _measure_name(value.get("measure"), "The measure to sort by") if by == "measure" else None
    return FieldSort(by=by, order=order, measure=measure)


def _label_filter(value: object) -> LabelFilter | None:
    if value is None:
        return None
    if not isinstance(value, dict) or value.get("operator") not in LABEL_OPERATORS:
        raise _bad("A label filter is not valid.")
    operator = value["operator"]
    first = _filter_text(value.get("value", ""), "The label filter's value")
    second = None
    if operator in RANGE_OPERATORS:
        second = _filter_text(value.get("value2"), "The label filter's second value")
        if first.lower() > second.lower():
            raise _bad("In a label filter's range, the first value must come before the second.")
    return LabelFilter(operator=operator, value=first, value2=second)


def _value_filter(value: object) -> ValueFilter | None:
    if value is None:
        return None
    if not isinstance(value, dict) or value.get("operator") not in VALUE_OPERATORS:
        raise _bad("A value filter is not valid.")
    operator = value["operator"]
    measure = _measure_name(value.get("measure"), "The value filter's measure")
    first = _number(value.get("value"), "The value filter's number")
    second = None
    top_by = "items"
    if operator in RANGE_OPERATORS:
        second = _number(value.get("value2"), "The value filter's second number")
        if first > second:
            raise _bad("In a value filter's range, the first number must not be larger than the second.")
    if operator in ("top", "bottom"):
        top_by = value.get("topBy", "items") or "items"
        if top_by not in TOP_BY:
            raise _bad("A Top 10 filter must count items, percent or sum.")
        if top_by == "items" and not (first.is_integer() and 1 <= first <= MAX_TOP_ITEMS):
            raise _bad(f"A Top 10 filter needs a whole number of items from 1 to {MAX_TOP_ITEMS:,}.")
        if top_by == "percent" and not 0 < first <= 100:
            raise _bad("A Top 10 filter's percent must be more than 0 and at most 100.")
    return ValueFilter(operator=operator, measure=measure, value=first, value2=second, top_by=top_by)


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
                # Page filters only select items, as in Excel.
                sort=_sort(item.get("sort")) if area != "filters" else None,
                label_filter=_label_filter(item.get("labelFilter")) if area != "filters" else None,
                value_filter=_value_filter(item.get("valueFilter")) if area != "filters" else None,
            )
        )
    return tuple(out)


def _value_list(value: object) -> tuple[ValueField, ...]:
    if not isinstance(value, list):
        raise _bad("The values area must be a list.")
    if len(value) > 64:
        raise _bad("The values area has more than 64 measures.")
    out = []
    seen: set[tuple[str, str, str, str]] = set()
    for item in value:
        if not isinstance(item, dict):
            raise _bad("A value field is not an object.")
        measure = validate_unique_name(item.get("measure"))
        if not measure.lower().startswith("[measures]."):
            raise _bad("Only measures can be placed in the Values area.")
        show_as = item.get("showAs") or "normal"
        if show_as not in SHOW_AS:
            raise _bad("A value field's 'Show Values As' setting is not valid.")
        caption = _text(item.get("caption", ""), "caption")
        base_field = base_item = None
        if show_as in NEEDS_BASE_FIELD:
            if not item.get("baseField"):
                raise _bad(f"'{caption}' needs a base field for this 'Show Values As' setting.")
            base_field = validate_unique_name(item.get("baseField"))
        if show_as in NEEDS_BASE_ITEM:
            raw_item = item.get("baseItem")
            if not raw_item:
                raise _bad(f"'{caption}' needs a base item for this 'Show Values As' setting.")
            base_item = raw_item if raw_item in RELATIVE_ITEMS else validate_unique_name(raw_item)
        # The same measure may appear twice if it is shown two ways (count and %).
        key = (measure.lower(), show_as, (base_field or "").lower(), (base_item or "").lower())
        if key in seen:
            raise _bad("A measure is used twice in the Values area.")
        seen.add(key)
        out.append(
            ValueField(
                measure=measure, caption=caption, show_as=show_as, base_field=base_field, base_item=base_item
            )
        )
    return tuple(out)
