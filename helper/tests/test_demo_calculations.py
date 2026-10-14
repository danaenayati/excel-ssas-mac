"""The sample cube's Date Filters, KPI, named sets, groups, calculated items and
measures, and repeated measures, through the same grid builder the pane reads."""

import datetime

import pytest
from test_demo_backend import CITY, SERVICE, demo, placed  # noqa: F401  (fixture)

from cubebridge.backends.demo_backend import CATALOG, CUBE
from cubebridge.errors import CubeError
from cubebridge.pivot.grid import build_grid
from cubebridge.pivot.model import Layout

MONTH = "[Date].[Month Start]"
RIDES = {"measure": "[Measures].[Rides]", "caption": "Rides"}


def member(demo, hierarchy, caption):
    found = demo.members(CATALOG, CUBE, hierarchy, None, None, 500)["members"]
    return next(m["uniqueName"] for m in found if m["caption"] == caption)


def run(demo, rows=(), columns=(), values=(RIDES,), calculated=None, **options):
    data = {"rows": list(rows), "columns": list(columns), "values": list(values), "filters": [], "options": options}
    if calculated:
        data["calculatedMeasures"] = calculated
    layout = Layout.from_json(data)
    cellset, text = demo.query(CATALOG, CUBE, layout)
    return build_grid(cellset, layout, mdx=text)


def labels(grid):
    return [line[0]["text"] for line in grid["rowHeaders"]]


def by_label(grid, col=0):
    return {line[0]["text"]: values[col] for line, values in zip(grid["rowHeaders"], grid["values"])}


def test_month_start_is_a_date_field(demo):
    fields = demo.fields(CATALOG, CUBE)
    month = next(h for d in fields["dimensions"] for h in d["hierarchies"] if h["uniqueName"] == MONTH)
    assert month["isDate"] is True


@pytest.mark.parametrize(
    "date_filter, expected",
    [
        ({"operator": "between", "value": "2025-03-01", "value2": "2025-05-31"}, ["2025-03-01", "2025-04-01", "2025-05-01"]),
        ({"operator": "quarter4"}, ["2025-10-01", "2025-11-01", "2025-12-01"]),
        ({"operator": "month2"}, ["2025-02-01", "2026-02-01"]),
        ({"operator": "before", "value": "2025-02-15"}, ["2025-01-01", "2025-02-01"]),
        ({"operator": "after", "value": "2026-08-01"}, ["2026-09-01"]),
    ],
)
def test_date_filters_keep_the_matching_dates_and_their_totals(demo, date_filter, expected):
    grid = run(demo, rows=[dict(placed(demo, MONTH), dateFilter=date_filter)])
    assert labels(grid)[:-1] == expected and labels(grid)[-1] == "Grand Total"
    assert grid["values"][-1][0] == sum(v[0] for v in grid["values"][:-1])


def test_kpi_parts(demo):
    parts = [
        {"measure": "[Measures].[Avg Fare]", "caption": f"Average Fare {p.title()}", "kpi": {"name": "Average Fare", "part": p}}
        for p in ("value", "goal", "status", "trend")
    ]
    grid = run(demo, rows=[placed(demo, CITY)], values=parts)
    assert [c["text"] for c in grid["colHeaderRows"][-1]] == [
        "Average Fare Value", "Average Fare Goal", "Average Fare Status", "Average Fare Trend",
    ]
    assert grid["colValueIndex"] == [0, 1, 2, 3]
    for value, goal, status, trend in grid["values"]:
        assert goal == 4.0 and status in (-1, 0, 1) and trend in (-1, 0, 1)
        assert status == (1 if value >= 4.0 else 0 if value >= 3.6 else -1)
    with pytest.raises(CubeError):
        run(demo, values=[dict(parts[0], kpi={"name": "Nope", "part": "value"})])


def test_named_set_shows_its_items_in_order_and_totals_only_them(demo):
    city = demo.fields(CATALOG, CUBE)["sets"][1]
    assert city["uniqueName"] == "[Top 3 Cities]" and city["hierarchy"] == CITY
    top = dict(placed(demo, CITY), set="[Top 3 Cities]", caption="Top 3 Cities")
    grid = run(demo, rows=[top])
    everything = by_label(run(demo, rows=[placed(demo, CITY)]))
    assert labels(grid)[:3] == sorted((c for c in everything if c != "Grand Total"), key=lambda c: -everything[c])[:3]
    assert grid["values"][-1][0] == sum(v[0] for v in grid["values"][:3])


def test_group_adds_up_its_items_and_hides_them(demo):
    group = {"name": "Big Two", "members": [member(demo, CITY, "Tehran"), member(demo, CITY, "Istanbul")]}
    grid = run(demo, rows=[dict(placed(demo, CITY), customItems=[group])])
    everything = by_label(run(demo, rows=[placed(demo, CITY)]))
    found = by_label(grid)
    assert "Tehran" not in found and "Istanbul" not in found
    assert found["Big Two"] == everything["Tehran"] + everything["Istanbul"]
    assert found["Grand Total"] == everything["Grand Total"]
    ref = grid["rowHeaders"][labels(grid).index("Big Two")][0]["member"]
    assert ref["uniqueName"] == "[Geography].[City].[All].[Big Two]" and ref["calculated"] is True
    assert ref["hasChildren"] is False


def test_group_sorts_with_the_items(demo):
    group = {"name": "Big Two", "members": [member(demo, CITY, "Tehran"), member(demo, CITY, "Istanbul")]}
    field = dict(placed(demo, CITY), customItems=[group], sort={"by": "measure", "measure": "[Measures].[Rides]", "order": "desc"})
    assert labels(run(demo, rows=[field]))[0] == "Big Two"


def test_calculated_item_works_out_its_formula_on_every_line(demo):
    tehran, mashhad = member(demo, CITY, "Tehran"), member(demo, CITY, "Mashhad")
    gap = {"name": "Gap", "formula": f"{tehran} - {mashhad} * 2"}
    grid = run(demo, rows=[dict(placed(demo, CITY), customItems=[gap])], columns=[placed(demo, SERVICE)])
    everything = run(demo, rows=[placed(demo, CITY)], columns=[placed(demo, SERVICE)])
    rows = {line[0]["text"]: values for line, values in zip(everything["rowHeaders"], everything["values"])}
    found = {line[0]["text"]: values for line, values in zip(grid["rowHeaders"], grid["values"])}
    assert found["Gap"] == [t - m * 2 for t, m in zip(rows["Tehran"], rows["Mashhad"])]
    assert found["Grand Total"] == rows["Grand Total"]  # items are not added twice


def test_calculated_measure_is_worked_out_from_totals(demo):
    calculated = [{"name": "Fare", "formula": "[Measures].[Revenue] / [Measures].[Rides]", "formatString": "#,##0.00"}]
    grid = run(demo, rows=[placed(demo, CITY)], values=[{"measure": "[Measures].[Fare]", "caption": "Fare"}], calculated=calculated)
    plain = run(demo, rows=[placed(demo, CITY)], values=["Rides", "Revenue"] and [RIDES, {"measure": "[Measures].[Revenue]", "caption": "Revenue"}])
    for (rides, revenue), (fare,) in zip(plain["values"], grid["values"]):
        assert fare == pytest.approx(revenue / rides)
    assert grid["colHeaderRows"][-1][0]["text"] == "Fare" and grid["formatPalette"][grid["cellFormats"][0][0]] == "#,##0.00"


def test_calculated_measure_can_sort_and_filter(demo):
    calculated = [{"name": "Fare", "formula": "[Measures].[Revenue] / [Measures].[Rides]"}]
    field = dict(placed(demo, CITY), sort={"by": "measure", "measure": "[Measures].[Fare]", "order": "desc"},
                 valueFilter={"operator": "top", "measure": "[Measures].[Fare]", "value": 2})
    grid = run(demo, rows=[field], values=[{"measure": "[Measures].[Fare]", "caption": "Fare"}], calculated=calculated)
    assert len(labels(grid)) == 3 and grid["values"][0][0] >= grid["values"][1][0]


def test_formula_the_sample_cube_cannot_work_out_is_explained(demo):
    calculated = [{"name": "Odd", "formula": "IIF([Measures].[Rides] > 0, 1, 0)"}]
    with pytest.raises(CubeError) as info:
        run(demo, values=[{"measure": "[Measures].[Odd]", "caption": "Odd"}], calculated=calculated)
    assert "A real cube understands any MDX formula" in info.value.message


def test_the_same_measure_twice(demo):
    grid = run(demo, rows=[placed(demo, CITY)], values=[RIDES, RIDES])
    assert [c["text"] for c in grid["colHeaderRows"][-1]] == ["Rides", "Rides 2"]
    assert grid["colValueIndex"] == [0, 1]
    assert all(a == b for a, b in grid["values"])


def test_show_details_refuses_a_group(demo):
    group = {"name": "Big Two", "members": [member(demo, CITY, "Tehran"), member(demo, CITY, "Istanbul")]}
    layout = Layout.from_json({"rows": [dict(placed(demo, CITY), customItems=[group])], "values": [RIDES]})
    with pytest.raises(CubeError) as info:
        demo.details(CATALOG, CUBE, layout, ["[Geography].[City].[All].[Big Two]"], "[Measures].[Rides]", 10)
    assert "group" in info.value.message


def test_relative_date_filter_uses_today(demo):
    today = datetime.date.today()
    grid = run(demo, rows=[dict(placed(demo, MONTH), dateFilter={"operator": "thisYear"})])
    assert all(label.startswith(str(today.year)) for label in labels(grid)[:-1])
