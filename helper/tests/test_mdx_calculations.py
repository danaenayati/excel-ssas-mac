"""Calculated measures, groups and calculated items, Date Filters, named sets, KPI
parts and repeated measures in the generated MDX, and how the layout validates them."""

import datetime

import pytest

from cubebridge.errors import CubeError
from cubebridge.pivot import dates, mdx
from cubebridge.pivot.model import DateFilter, Layout

RIDES = {"measure": "[Measures].[Rides]", "caption": "Rides"}
CITY = "[Geo].[City]"
DAY = "[Date].[Day]"
TEHRAN, MASHHAD, SHIRAZ = (f"[Geo].[City].&[{c}]" for c in ("Tehran", "Mashhad", "Shiraz"))
TODAY = datetime.date(2026, 10, 14)  # a Wednesday


def field(hierarchy, **extra):
    out = {
        "hierarchy": hierarchy,
        "caption": hierarchy.split(".")[-1].strip("[]"),
        "allMember": f"{hierarchy}.[All]",
        "expanded": [],
        "members": [],
        "memberCaptions": [],
    }
    out.update(extra)
    return out


def layout(rows=(), columns=(), values=(RIDES,), filters=(), calculated=None, **options):
    data = {"rows": list(rows), "columns": list(columns), "values": list(values), "filters": list(filters), "options": options}
    if calculated is not None:
        data["calculatedMeasures"] = calculated
    return Layout.from_json(data)


def query(*args, **kwargs):
    return mdx.build_query("Sales", layout(*args, **kwargs), today=TODAY)


# -- calculated measures ---------------------------------------------------------------
def test_calculated_measure_is_defined_with_a_high_solve_order_and_usable_in_values():
    text = query(
        values=[{"measure": "[Measures].[Fare]", "caption": "Fare"}],
        calculated=[{"name": "Fare", "formula": "[Measures].[Revenue] / [Measures].[Rides]", "formatString": "#,##0.00"}],
    )
    assert text.startswith(
        'WITH MEMBER [Measures].[Fare] AS ([Measures].[Revenue] / [Measures].[Rides]), FORMAT_STRING = "#,##0.00", '
        "SOLVE_ORDER = 100 SELECT NON EMPTY {[Measures].[Fare]} ON COLUMNS"
    )


def test_calculated_measure_is_defined_even_when_only_a_sort_uses_it():
    sort = {"by": "measure", "measure": "[Measures].[Fare]", "order": "desc"}
    text = query(rows=[field(CITY, sort=sort)], calculated=[{"name": "Fare", "formula": "1"}])
    assert "MEMBER [Measures].[Fare] AS (1), SOLVE_ORDER = 100" in text
    assert "Order({[Geo].[City].[All].Children}, [Measures].[Fare], DESC)" in text


@pytest.mark.parametrize(
    "calculated, message",
    [
        ([{"name": "", "formula": "1"}], "needs a name"),
        ([{"name": "A", "formula": "1"}, {"name": "a", "formula": "2"}], "both called"),
        ([{"name": "A", "formula": "1; DROP"}], "';'"),
        ([{"name": "A", "formula": "1", "formatString": "x\x01"}], "number format"),
        ([{"name": "A", "formula": "1"}] * 51, "at most 50"),
    ],
)
def test_calculated_measures_that_cannot_be_used_are_refused(calculated, message):
    with pytest.raises(CubeError) as info:
        layout(calculated=calculated)
    assert message in info.value.message


# -- groups and calculated items -----------------------------------------------------
def test_group_replaces_its_items_and_adds_them_up():
    groups = [{"name": "Big", "members": [TEHRAN, MASHHAD]}]
    text = query(rows=[field(CITY, customItems=groups)])
    assert text.startswith("WITH MEMBER [Geo].[City].[All].[Big] AS Aggregate({[Geo].[City].&[Tehran], [Geo].[City].&[Mashhad]}) SELECT")
    assert (
        "{{Except({[Geo].[City].[All].Children}, {[Geo].[City].&[Tehran], [Geo].[City].&[Mashhad]}), [Geo].[City].[All].[Big]}, "
        "{[Geo].[City].[All]}} ON ROWS"
    ) in text


def test_calculated_item_follows_the_items_and_keeps_them():
    items = [{"name": "Gap", "formula": f"{TEHRAN} - {MASHHAD}"}]
    text = query(columns=[field(CITY, customItems=items)])
    assert "MEMBER [Geo].[City].[All].[Gap] AS ([Geo].[City].&[Tehran] - [Geo].[City].&[Mashhad])" in text
    assert "{{{[Geo].[City].[All].Children}, [Geo].[City].[All].[Gap]}, {[Geo].[City].[All]}}" in text


def test_measures_calculated_in_the_query_are_solved_after_items():
    text = query(
        rows=[field(CITY, customItems=[{"name": "Big", "members": [TEHRAN, MASHHAD]}])],
        values=[RIDES, dict(RIDES, showAs="percentOfGrandTotal")],
    )
    item = text.index("MEMBER [Geo].[City].[All].[Big]")
    shown = text.index("MEMBER [Measures].[Rides (% of Grand Total)]")
    assert "SOLVE_ORDER" not in text[item:shown] and "SOLVE_ORDER = 100" in text[shown:]


def test_group_items_sort_with_the_others():
    groups = [{"name": "Big", "members": [TEHRAN, MASHHAD]}]
    text = query(rows=[field(CITY, customItems=groups, sort={"by": "caption", "order": "asc"})])
    assert "Order({Except(" in text and "[Geo].[City].[All].[Big]}, [Geo].[City].CurrentMember.Member_Caption, ASC)" in text


@pytest.mark.parametrize(
    "items, message",
    [
        ([{"name": "G", "members": [TEHRAN]}], "at least two"),
        ([{"name": "G", "members": [TEHRAN, "[Date].[Day].&[1]"]}], "another field"),
        ([{"name": "G", "members": [TEHRAN, MASHHAD]}, {"name": "H", "members": [TEHRAN, SHIRAZ]}], "two groups"),
        ([{"name": "G", "members": [TEHRAN, MASHHAD], "formula": "1"}], "either"),
        ([{"name": "G"}], "either"),
        ([{"name": "G", "formula": "1"}, {"name": "g", "formula": "2"}], "both called"),
        ([{"name": "G", "formula": "SELECT 1"}], "statement"),
    ],
)
def test_custom_items_that_cannot_be_used_are_refused(items, message):
    with pytest.raises(CubeError) as info:
        layout(rows=[field(CITY, customItems=items)])
    assert message in info.value.message


def test_custom_items_and_date_filters_are_ignored_in_page_filters():
    filters = [field(CITY, customItems=[{"name": "G", "formula": "1"}], dateFilter={"operator": "today"})]
    placed = layout(filters=filters).filters[0]
    assert placed.custom_items == () and placed.date_filter is None


def test_show_details_refuses_groups_and_calculated_measures():
    groups = [{"name": "Big", "members": [TEHRAN, MASHHAD]}]
    grouped = layout(rows=[field(CITY, customItems=groups)])
    with pytest.raises(CubeError) as info:
        mdx.details_slicer(grouped, ["[Geo].[City].[All].[Big]"], "[Measures].[Rides]")
    assert "group or calculated item" in info.value.message
    assert mdx.details_slicer(grouped, [SHIRAZ], "[Measures].[Rides]") == ["[Measures].[Rides]", SHIRAZ]
    calculated = layout(calculated=[{"name": "Fare", "formula": "1"}])
    with pytest.raises(CubeError) as info:
        mdx.details_slicer(calculated, [], "[Measures].[Fare]")
    assert "calculated measure" in info.value.message


# -- Date Filters ------------------------------------------------------------------------
@pytest.mark.parametrize(
    "operator, value, value2, start, end",
    [
        ("equals", "2026-03-05", None, "2026-03-05", "2026-03-06"),
        ("before", "2026-03-05", None, None, "2026-03-05"),
        ("after", "2026-03-05", None, "2026-03-06", None),
        ("between", "2026-03-05", "2026-03-31", "2026-03-05", "2026-04-01"),
        ("today", None, None, "2026-10-14", "2026-10-15"),
        ("yesterday", None, None, "2026-10-13", "2026-10-14"),
        ("tomorrow", None, None, "2026-10-15", "2026-10-16"),
        ("thisWeek", None, None, "2026-10-11", "2026-10-18"),
        ("lastWeek", None, None, "2026-10-04", "2026-10-11"),
        ("nextWeek", None, None, "2026-10-18", "2026-10-25"),
        ("thisMonth", None, None, "2026-10-01", "2026-11-01"),
        ("lastMonth", None, None, "2026-09-01", "2026-10-01"),
        ("nextMonth", None, None, "2026-11-01", "2026-12-01"),
        ("thisQuarter", None, None, "2026-10-01", "2027-01-01"),
        ("lastQuarter", None, None, "2026-07-01", "2026-10-01"),
        ("nextQuarter", None, None, "2027-01-01", "2027-04-01"),
        ("thisYear", None, None, "2026-01-01", "2027-01-01"),
        ("lastYear", None, None, "2025-01-01", "2026-01-01"),
        ("nextYear", None, None, "2027-01-01", "2028-01-01"),
        ("yearToDate", None, None, "2026-01-01", "2026-10-15"),
    ],
)
def test_date_windows(operator, value, value2, start, end):
    as_date = lambda text: datetime.date.fromisoformat(text) if text else None  # noqa: E731
    found = dates.window(DateFilter(operator, as_date(value), as_date(value2)), TODAY)
    assert (found.start, found.end) == (as_date(start), as_date(end))


def test_week_windows_start_on_sunday_even_on_a_sunday():
    sunday = datetime.date(2026, 10, 11)
    found = dates.window(DateFilter("thisWeek"), sunday)
    assert (found.start, found.end) == (sunday, datetime.date(2026, 10, 18))


def test_periods_match_any_year():
    q2 = dates.window(DateFilter("quarter2"), TODAY)
    assert q2.contains(datetime.date(2019, 5, 1)) and not q2.contains(datetime.date(2026, 7, 1))
    assert dates.window(DateFilter("month12"), TODAY).contains(datetime.date(2020, 12, 31))


def test_date_filter_becomes_a_subselect_on_member_values():
    text = query(rows=[field(DAY, dateFilter={"operator": "between", "value": "2026-01-01", "value2": "2026-01-31"})])
    value = "[Date].[Day].CurrentMember.MemberValue"
    assert (
        f"(SELECT {{Filter({{[Date].[Day].[All].Children}}, NOT IsEmpty({value}) AND "
        f"{value} >= DateSerial(2026, 1, 1) AND {value} < DateSerial(2026, 2, 1))}} ON COLUMNS FROM [Sales])"
    ) in text
    assert f'DatePart("q", {value}) = 3' in query(rows=[field(DAY, dateFilter={"operator": "quarter3"})])
    assert f"Month({value}) = 11" in query(rows=[field(DAY, dateFilter={"operator": "month11"})])
    assert f"{value} < DateSerial(2026, 3, 5))" in query(rows=[field(DAY, dateFilter={"operator": "before", "value": "2026-03-05"})])


@pytest.mark.parametrize(
    "date_filter, message",
    [
        ({"operator": "someday"}, "not valid"),
        ({"operator": "equals"}, "yyyy-mm-dd"),
        ({"operator": "equals", "value": "14/10/2026"}, "yyyy-mm-dd"),
        ({"operator": "between", "value": "2026-02-01", "value2": "2026-01-01"}, "must not be after"),
    ],
)
def test_date_filters_that_cannot_be_used_are_refused(date_filter, message):
    with pytest.raises(CubeError) as info:
        layout(rows=[field(DAY, dateFilter=date_filter)])
    assert message in info.value.message


def test_a_date_filter_narrows_the_field():
    placed = layout(rows=[field(DAY, dateFilter={"operator": "today"})]).rows[0]
    assert placed.narrowed


# -- named sets ----------------------------------------------------------------------------
def test_named_set_is_the_fields_items_and_limits_the_totals():
    top = field(CITY, set="[Top Cities]", members=[TEHRAN], sort={"by": "caption", "order": "asc"})
    placed = layout(rows=[top]).rows[0]
    assert placed.named_set == "[Top Cities]" and placed.members == () and placed.sort is None
    text = query(rows=[top])
    assert "NON EMPTY {{[Top Cities]}, {[Geo].[City].[All]}} ON ROWS" in text
    assert "FROM (SELECT {[Top Cities]} ON COLUMNS FROM [Sales])" in text


def test_named_set_cannot_be_a_page_filter():
    with pytest.raises(CubeError) as info:
        layout(filters=[field(CITY, set="[Top Cities]")])
    assert "Rows or Columns" in info.value.message


# -- values --------------------------------------------------------------------------------
def test_kpi_parts_are_calculated_members_named_after_their_captions():
    values = [
        {"measure": "[Measures].[Sales]", "caption": "Sales Goal", "kpi": {"name": "Sales", "part": "goal"}},
        {"measure": "[Measures].[Sales]", "caption": "Sales Status", "kpi": {"name": "Sales", "part": "status"}},
    ]
    text = query(values=values)
    assert 'MEMBER [Measures].[Sales Goal] AS KPIGoal("Sales"), SOLVE_ORDER = 100' in text
    assert 'MEMBER [Measures].[Sales Status] AS KPIStatus("Sales"), SOLVE_ORDER = 100' in text
    assert "{[Measures].[Sales Goal], [Measures].[Sales Status]} ON COLUMNS" in text


def test_kpi_parts_cannot_be_shown_another_way():
    with pytest.raises(CubeError) as info:
        layout(values=[{"measure": "[Measures].[S]", "caption": "S", "showAs": "index", "kpi": {"name": "S", "part": "goal"}}])
    assert "only be shown as it is" in info.value.message
    with pytest.raises(CubeError):
        layout(values=[{"measure": "[Measures].[S]", "caption": "S", "kpi": {"name": "S", "part": "forecast"}}])


def test_a_second_copy_of_a_measure_gets_its_own_name_and_the_measures_format():
    the_layout = layout(values=[RIDES, RIDES, dict(RIDES, showAs="index"), dict(RIDES, showAs="index")])
    members = mdx.value_members(the_layout, {"[Measures].[Rides]": "#,##0"})
    assert [m.name for m in members] == [
        "[Measures].[Rides]", "[Measures].[Rides 2]", "[Measures].[Rides (Index)]", "[Measures].[Rides (Index) 2]",
    ]
    assert members[1].expression == "[Measures].[Rides]" and members[1].format_string == "#,##0"
    text = mdx.build_query("Sales", the_layout, {"[Measures].[Rides]": "#,##0"})
    assert 'MEMBER [Measures].[Rides 2] AS [Measures].[Rides], FORMAT_STRING = "#,##0", SOLVE_ORDER = 100' in text
