import pytest

from cubebridge.backends.demo_backend import CATALOG, CUBE, DemoBackend
from cubebridge.errors import CubeError
from cubebridge.pivot.grid import build_grid
from cubebridge.pivot.model import Layout

CITY = "[Geography].[City]"
GEO = "[Geography].[Geography]"
DATE = "[Date].[Calendar]"
SERVICE = "[Service].[Service Type]"


@pytest.fixture(scope="module")
def demo():
    return DemoBackend()


def hierarchy(demo, unique_name):
    for dimension in demo.fields(CATALOG, CUBE)["dimensions"]:
        for h in dimension["hierarchies"]:
            if h["uniqueName"] == unique_name:
                return h
    raise AssertionError(unique_name)


def placed(demo, unique_name, expanded=(), members=()):
    h = hierarchy(demo, unique_name)
    return {
        "hierarchy": unique_name,
        "caption": h["caption"],
        "allMember": h["allMember"],
        "expanded": list(expanded),
        "members": list(members),
        "memberCaptions": [],
    }


def run(demo, rows=(), columns=(), values=("Rides",), filters=(), **options):
    layout = Layout.from_json(
        {
            "rows": list(rows),
            "columns": list(columns),
            "values": [{"measure": f"[Measures].[{v}]", "caption": v} for v in values],
            "filters": list(filters),
            "options": options,
        }
    )
    cellset, text = demo.query(CATALOG, CUBE, layout)
    return build_grid(cellset, layout, mdx=text)


def column(grid, index=0):
    return [line[index] for line in grid["values"]]


def test_catalog_cube_and_fields(demo):
    assert demo.catalogs() == [{"name": CATALOG, "description": "Built-in sample data. No server needed."}]
    assert [c["name"] for c in demo.cubes(CATALOG)] == [CUBE]
    fields = demo.fields(CATALOG, CUBE)
    assert {m["caption"] for m in fields["measures"]} >= {"Rides", "Revenue", "Avg Fare"}
    assert all(m["uniqueName"].startswith("[Measures].") for m in fields["measures"])
    names = {h["uniqueName"] for d in fields["dimensions"] for h in d["hierarchies"]}
    assert names >= {CITY, GEO, DATE, SERVICE}
    assert hierarchy(demo, DATE)["kind"] == "user" and hierarchy(demo, CITY)["kind"] == "attribute"
    assert hierarchy(demo, CITY)["allMember"] == "[Geography].[City].[All]"


def test_unknown_catalog_or_cube_is_not_found(demo):
    with pytest.raises(CubeError) as info:
        demo.cubes("Nope")
    assert info.value.category == "not-found"
    with pytest.raises(CubeError):
        demo.fields(CATALOG, "Nope")


def test_additive_total_is_the_sum_of_its_rows(demo):
    grid = run(demo, rows=[placed(demo, CITY)])
    assert grid["rowKinds"][-1] == "grandTotal" and grid["rowKinds"].count("data") == len(grid["rowKinds"]) - 1
    rides = column(grid)
    assert rides[-1] == sum(rides[:-1])
    assert all(isinstance(v, int) and v > 0 for v in rides)


def test_ratio_total_is_computed_not_summed(demo):
    grid = run(demo, rows=[placed(demo, CITY)], values=("Rides", "Revenue", "Avg Fare"))
    rides, revenue, fare = column(grid, 0), column(grid, 1), column(grid, 2)
    assert fare[-1] == pytest.approx(revenue[-1] / rides[-1])
    assert fare[-1] != pytest.approx(sum(fare[:-1]))
    assert grid["formatPalette"][grid["cellFormats"][0][1]] == "$#,##0.00"


def test_expanding_a_year_puts_its_quarters_directly_under_it(demo):
    collapsed = run(demo, rows=[placed(demo, DATE)], grandTotals=False)
    years = [line[0] for line in collapsed["rowHeaders"]]
    assert all(cell["indent"] == 0 and cell["member"]["hasChildren"] for cell in years)
    first = years[0]["member"]
    expanded = run(demo, rows=[placed(demo, DATE, expanded=[first["uniqueName"]])], grandTotals=False)
    lines = expanded["rowHeaders"]
    assert lines[0][0]["member"]["expanded"] is True
    quarters = [line[0] for line in lines[1:5]]
    assert [q["text"] for q in quarters] == ["Q1", "Q2", "Q3", "Q4"]
    assert all(q["indent"] == 1 for q in quarters)
    # the year's own value equals the sum of its quarters
    values = column(expanded)
    assert values[0] == sum(values[1:5])


def test_columns_and_two_measures(demo):
    grid = run(demo, rows=[placed(demo, CITY)], columns=[placed(demo, SERVICE)], values=("Rides", "Revenue"))
    services = [c["text"] for c in grid["colHeaderRows"][0] if c["text"]]
    assert services[-1] == "Grand Total" and len(services) >= 3
    assert [c["text"] for c in grid["colHeaderRows"][1]][:2] == ["Rides", "Revenue"]
    # last two columns are the row totals for each measure
    for line in grid["values"]:
        assert line[-2] == sum(v or 0 for v in line[0:-2:2])


def test_subtotals_for_two_row_fields(demo):
    grid = run(demo, rows=[placed(demo, CITY), placed(demo, SERVICE)], subtotals=True)
    assert "subtotal" in grid["rowKinds"]
    i = grid["rowKinds"].index("subtotal")
    assert grid["rowHeaders"][i][0]["text"].endswith(" Total")
    assert grid["values"][i][0] == sum(v[0] for v in grid["values"][:i])


def test_filter_reduces_totals_and_single_member_axis_selection_works(demo):
    everything = run(demo, rows=[placed(demo, CITY)])
    economy = next(m for m in demo.members(CATALOG, CUBE, SERVICE, None, None, 50)["members"] if m["caption"] == "Economy")
    filtered = run(demo, rows=[placed(demo, CITY)], filters=[placed(demo, SERVICE, members=[economy["uniqueName"]])])
    assert column(filtered)[-1] < column(everything)[-1]

    tehran = next(m for m in demo.members(CATALOG, CUBE, CITY, None, None, 50)["members"] if m["caption"] == "Tehran")
    only = run(demo, rows=[placed(demo, CITY, members=[tehran["uniqueName"]])])
    assert [line[0]["text"] for line in only["rowHeaders"]] == ["Tehran", "Grand Total"]
    assert column(only)[0] == column(only)[1]  # visual total


def test_non_empty_hides_combinations_without_data(demo):
    hidden = run(demo, rows=[placed(demo, CITY), placed(demo, SERVICE)], grandTotals=False)
    shown = run(demo, rows=[placed(demo, CITY), placed(demo, SERVICE)], grandTotals=False, nonEmpty=False)
    assert len(shown["rowHeaders"]) > len(hidden["rowHeaders"])
    assert any(line[0] is None for line in shown["values"])


def test_members_top_level_children_search_and_limit(demo):
    top = demo.members(CATALOG, CUBE, GEO, None, None, 50)
    assert [m["caption"] for m in top["members"]] == ["Iran", "Turkey", "UAE"]
    assert all(m["hasChildren"] and m["level"] == 1 for m in top["members"])
    children = demo.members(CATALOG, CUBE, GEO, top["members"][0]["uniqueName"], None, 50)
    assert "Tehran" in [m["caption"] for m in children["members"]]
    assert all(m["level"] == 2 and not m["hasChildren"] for m in children["members"])
    found = demo.members(CATALOG, CUBE, GEO, None, "an", 50)
    assert {"Iran", "Tehran", "Istanbul", "Ankara"} <= {m["caption"] for m in found["members"]}
    limited = demo.members(CATALOG, CUBE, CITY, None, None, 3)
    assert len(limited["members"]) == 3 and limited["truncated"] is True


def test_query_text_is_the_mdx_a_real_cube_would_get(demo):
    layout = Layout.from_json({"rows": [placed(demo, CITY)], "values": [{"measure": "[Measures].[Rides]", "caption": "Rides"}]})
    _, text = demo.query(CATALOG, CUBE, layout)
    assert text.startswith("SELECT NON EMPTY {[Measures].[Rides]} ON COLUMNS") and "FROM [Ride Sales]" in text


# -- sorting, filters, Show Values As, Show Details ----------------------------------
def labels(grid):
    return [line[0]["text"] for line in grid["rowHeaders"]]


def test_sort_by_caption_and_by_measure(demo):
    by_name = run(demo, rows=[dict(placed(demo, CITY), sort={"by": "caption", "order": "desc"})], grandTotals=False)
    assert labels(by_name) == sorted(labels(by_name), key=str.lower, reverse=True)
    by_rides = run(demo, rows=[dict(placed(demo, CITY), sort={"by": "measure", "measure": "[Measures].[Rides]", "order": "desc"})], grandTotals=False)
    rides = column(by_rides)
    assert rides == sorted(rides, reverse=True) and labels(by_rides)[0] == "Tehran"


def test_sorting_a_drilled_hierarchy_keeps_children_under_their_parent(demo):
    first_year = demo.members(CATALOG, CUBE, DATE, None, None, 5)["members"][0]["uniqueName"]
    grid = run(demo, rows=[dict(placed(demo, DATE, expanded=[first_year]), sort={"by": "caption", "order": "desc"})], grandTotals=False)
    texts = labels(grid)
    year = demo.members(CATALOG, CUBE, DATE, None, None, 5)["members"][0]["caption"]
    i = texts.index(year)
    assert texts[i + 1 : i + 5] == ["Q4", "Q3", "Q2", "Q1"]


def test_label_filter_keeps_matching_items_and_totals_count_only_them(demo):
    grid = run(demo, rows=[dict(placed(demo, CITY), labelFilter={"operator": "contains", "value": "AN"})])
    assert labels(grid)[:-1] and all("an" in t.lower() for t in labels(grid)[:-1])
    assert column(grid)[-1] == sum(column(grid)[:-1])  # visual total
    none = run(demo, rows=[dict(placed(demo, CITY), labelFilter={"operator": "beginsWith", "value": "zzz"})])
    assert none["rowHeaders"] == []


def test_value_filters_top_bottom_and_comparisons(demo):
    top = run(demo, rows=[dict(placed(demo, CITY), valueFilter={"operator": "top", "measure": "[Measures].[Rides]", "value": 2})])
    assert labels(top) == ["Istanbul", "Tehran", "Grand Total"] or set(labels(top)[:-1]) == {"Tehran", "Istanbul"}
    assert column(top)[-1] == sum(column(top)[:-1])
    everything = run(demo, rows=[placed(demo, CITY)], grandTotals=False)
    rides = dict(zip(labels(everything), column(everything)))
    big = run(demo, rows=[dict(placed(demo, CITY), valueFilter={"operator": "greaterThan", "measure": "[Measures].[Rides]", "value": 300000})], grandTotals=False)
    assert set(labels(big)) == {c for c, v in rides.items() if v > 300000}
    bottom = run(demo, rows=[dict(placed(demo, CITY), valueFilter={"operator": "bottom", "measure": "[Measures].[Rides]", "value": 1})], grandTotals=False)
    assert labels(bottom) == [min(rides, key=rides.get)]
    pct = run(demo, rows=[dict(placed(demo, CITY), valueFilter={"operator": "top", "measure": "[Measures].[Rides]", "value": 50, "topBy": "percent"})], grandTotals=False)
    kept = sum(column(pct))
    assert kept >= 0.5 * sum(rides.values()) and len(labels(pct)) < len(rides)


def test_show_values_as_percentages(demo):
    grid = run(
        demo,
        rows=[placed(demo, CITY)],
        columns=[placed(demo, SERVICE)],
        values=("Rides",),
    )
    layout = Layout.from_json(
        {
            "rows": [placed(demo, CITY)],
            "columns": [placed(demo, SERVICE)],
            "values": [{"measure": "[Measures].[Rides]", "caption": "Rides", "showAs": "percentOfGrandTotal"}],
        }
    )
    cellset, _ = demo.query(CATALOG, CUBE, layout)
    pct = build_grid(cellset, layout)
    assert pct["colHeaderRows"][-1][0]["text"] == "Rides (% of Grand Total)"
    assert pct["formatPalette"][pct["cellFormats"][0][0]] == "0.00%"
    assert pct["values"][-1][-1] == pytest.approx(1.0)  # grand total of grand total
    assert pct["values"][0][0] == pytest.approx(grid["values"][0][0] / grid["values"][-1][-1])

    by_row = Layout.from_json(
        {
            "rows": [placed(demo, CITY)],
            "columns": [placed(demo, SERVICE)],
            "values": [{"measure": "[Measures].[Rides]", "caption": "Rides", "showAs": "percentOfRowTotal"}],
        }
    )
    row_pct = build_grid(demo.query(CATALOG, CUBE, by_row)[0], by_row)
    for line in row_pct["values"]:
        assert sum(v or 0 for v in line[:-1]) == pytest.approx(1.0) and line[-1] == pytest.approx(1.0)


def test_show_details_returns_the_rows_behind_a_cell(demo):
    layout = Layout.from_json({"rows": [placed(demo, CITY)], "values": [{"measure": "[Measures].[Rides]", "caption": "Rides"}]})
    tehran = next(m for m in demo.members(CATALOG, CUBE, CITY, None, None, 50)["members"] if m["caption"] == "Tehran")
    result = demo.details(CATALOG, CUBE, layout, [tehran["uniqueName"]], "[Measures].[Rides]", 50)
    assert "City" in result["columns"] and "Rides" in result["columns"]
    city = result["columns"].index("City")
    assert len(result["rows"]) == 50 and result["truncated"] is True
    assert all(row[city] == "Tehran" for row in result["rows"])
    everything = demo.details(CATALOG, CUBE, layout, [tehran["uniqueName"]], "[Measures].[Rides]", 100000)
    rides = result["columns"].index("Rides")
    grid = run(demo, rows=[placed(demo, CITY)], grandTotals=False)
    assert sum(r[rides] for r in everything["rows"]) == dict(zip(labels(grid), column(grid)))["Tehran"]
    assert everything["truncated"] is False


def test_show_details_is_refused_like_excel_when_a_filter_covers_several_items(demo):
    two = [m["uniqueName"] for m in demo.members(CATALOG, CUBE, SERVICE, None, None, 50)["members"][:2]]
    layout = Layout.from_json({"filters": [placed(demo, SERVICE, members=two)], "values": [{"measure": "[Measures].[Rides]", "caption": "Rides"}]})
    with pytest.raises(CubeError) as info:
        demo.details(CATALOG, CUBE, layout, [], "[Measures].[Rides]", 10)
    assert "Show Details" in info.value.message


# -- Windows parity: per outer item, per-axis totals, every Show Values As kind ------
def shown_grid(demo, rows, columns=(), shows=(), **options):
    values = [{"measure": "[Measures].[Rides]", "caption": "Rides"}] + [
        dict({"measure": "[Measures].[Rides]", "caption": "Rides"}, **s) for s in shows
    ]
    layout = Layout.from_json({"rows": list(rows), "columns": list(columns), "values": values, "options": options})
    cellset, _ = demo.query(CATALOG, CUBE, layout)
    return build_grid(cellset, layout)


def test_grand_totals_per_axis_in_the_sample_cube(demo):
    grid = shown_grid(demo, [placed(demo, CITY)], [placed(demo, SERVICE)], grandTotalRow=False, grandTotalColumn=True)
    assert "grandTotal" not in grid["rowKinds"] and grid["colKinds"][-1] == "grandTotal"
    grid = shown_grid(demo, [placed(demo, CITY)], [placed(demo, SERVICE)], grandTotalRow=True, grandTotalColumn=False)
    assert grid["rowKinds"][-1] == "grandTotal" and "grandTotal" not in grid["colKinds"]


def test_top_1_inner_item_is_chosen_within_each_outer_item(demo):
    inner = dict(placed(demo, SERVICE), valueFilter={"operator": "top", "measure": "[Measures].[Rides]", "value": 1})
    grid = run(demo, rows=[placed(demo, CITY), inner], grandTotals=False)
    cities = [line[0]["member"]["caption"] for line in grid["rowHeaders"]]
    assert len(cities) == len(set(cities)) == 9  # exactly one service per city
    with_totals = run(demo, rows=[placed(demo, CITY), inner], subtotals=True)
    values = column(with_totals)
    kinds = with_totals["rowKinds"]
    assert values[-1] == sum(v for v, k in zip(values, kinds) if k == "data")  # visual grand total


def test_sort_by_value_of_an_inner_field_is_within_each_outer_item(demo):
    inner = dict(placed(demo, SERVICE), sort={"by": "measure", "measure": "[Measures].[Rides]", "order": "desc"})
    grid = run(demo, rows=[placed(demo, CITY), inner], grandTotals=False)
    by_city: dict[str, list[int]] = {}
    for line, values in zip(grid["rowHeaders"], grid["values"]):
        by_city.setdefault(line[0]["member"]["caption"], []).append(values[0])
    assert all(v == sorted(v, reverse=True) for v in by_city.values())


def test_running_total_rank_and_difference_from_previous(demo):
    city = placed(demo, CITY)
    grid = shown_grid(
        demo,
        [city],
        shows=[
            {"showAs": "runningTotalIn", "baseField": CITY},
            {"showAs": "rankDescending", "baseField": CITY},
            {"showAs": "differenceFrom", "baseField": CITY, "baseItem": "previous"},
            {"showAs": "percentRunningTotalIn", "baseField": CITY},
        ],
    )
    assert [c["text"] for c in grid["colHeaderRows"][-1]] == [
        "Rides",
        "Rides (Running Total In City)",
        "Rides (Rank Largest to Smallest City)",
        "Rides (Difference From City, previous)",
        "Rides (% Running Total In City)",
    ]
    lines = grid["values"][:-1]
    rides = [line[0] for line in lines]
    assert [line[1] for line in lines] == [sum(rides[: i + 1]) for i in range(len(rides))]
    assert lines[-1][1] == grid["values"][-1][0]  # the last running total is the grand total
    assert lines[-1][4] == pytest.approx(1.0)
    ranks = [line[2] for line in lines]
    assert sorted(ranks) == list(range(1, len(ranks) + 1)) and ranks[rides.index(max(rides))] == 1
    assert lines[0][3] is None and lines[1][3] == rides[1] - rides[0]
    assert grid["values"][-1][3] is None  # blank on the total line, as in Excel


def test_percent_of_parent_row_total_and_parent_total(demo):
    rows = [placed(demo, CITY), placed(demo, SERVICE)]
    grid = shown_grid(
        demo, rows, shows=[{"showAs": "percentOfParentRowTotal"}, {"showAs": "percentOfParentTotal", "baseField": CITY}], subtotals=True
    )
    shares: dict[str, float] = {}
    for line, kind, values in zip(grid["rowHeaders"], grid["rowKinds"], grid["values"]):
        if kind == "data":
            shares[line[0]["member"]["caption"]] = shares.get(line[0]["member"]["caption"], 0) + values[1]
            assert values[2] == pytest.approx(values[1])  # parent total of City = the city's subtotal
    assert all(total == pytest.approx(1.0) for total in shares.values())
    assert grid["values"][-1][1] == pytest.approx(1.0)


def test_percent_of_an_item_difference_and_index(demo):
    members = demo.members(CATALOG, CUBE, CITY, None, None, 50)["members"]
    tehran = next(m for m in members if m["caption"] == "Tehran")["uniqueName"]
    grid = shown_grid(
        demo,
        [placed(demo, CITY)],
        [placed(demo, SERVICE)],
        shows=[
            {"showAs": "percentOf", "baseField": CITY, "baseItem": tehran},
            {"showAs": "percentDifferenceFrom", "baseField": CITY, "baseItem": tehran},
            {"showAs": "index"},
        ],
    )
    labels_ = [line[0]["text"] for line in grid["rowHeaders"]]
    t = labels_.index("Tehran")
    assert grid["values"][t][1] == pytest.approx(1.0) and grid["values"][t][2] is None
    other = labels_.index("Shiraz")
    assert grid["values"][other][2] == pytest.approx(grid["values"][other][1] - 1)
    # index of the grand total cell is 1 by definition
    assert grid["values"][-1][-1] == pytest.approx(1.0)
