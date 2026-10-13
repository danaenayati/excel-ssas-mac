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
