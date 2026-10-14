"""Slicers, one level of a user hierarchy, "Include filtered items in totals" and named
sets spanning several fields: in the MDX, in the grid, and in the Sample cube."""

import pytest
from test_demo_backend import CITY, SERVICE, demo, placed  # noqa: F401  (fixture)

from cubebridge.backends.demo_backend import CATALOG, CUBE
from cubebridge.errors import CubeError
from cubebridge.pivot import mdx
from cubebridge.pivot.grid import build_grid
from cubebridge.pivot.model import Layout

RIDES = {"measure": "[Measures].[Rides]", "caption": "Rides"}
DATE = "[Date].[Calendar]"
MONTH_LEVEL = {"uniqueName": "[Date].[Calendar].[Month]", "number": 3, "caption": "Month"}


def field(hierarchy, **extra):
    out = {"hierarchy": hierarchy, "caption": hierarchy.split(".")[-1].strip("[]"),
           "allMember": f"{hierarchy}.[All]", "expanded": [], "members": [], "memberCaptions": []}
    out.update(extra)
    return out


def layout(rows=(), columns=(), values=(RIDES,), filters=(), slicers=None, **options):
    data = {"rows": list(rows), "columns": list(columns), "values": list(values), "filters": list(filters),
            "options": options}
    if slicers is not None:
        data["slicers"] = slicers
    return Layout.from_json(data)


def run(demo, *args, **kwargs):
    the_layout = layout(*args, **kwargs)
    cellset, text = demo.query(CATALOG, CUBE, the_layout)
    return build_grid(cellset, the_layout, mdx=text)


def member(demo, hierarchy, caption, level=None):
    found = demo.members(CATALOG, CUBE, hierarchy, None, None, 500, level=level)["members"]
    return next(m["uniqueName"] for m in found if m["caption"] == caption)


def labels(grid, col=0):
    return [line[col]["text"] for line in grid["rowHeaders"]]


# -- slicers ----------------------------------------------------------------------------------
def test_slicer_filters_like_a_page_filter_and_leaves_no_filter_row():
    text = mdx.build_query("Sales", layout(rows=[field("[Geo].[City]")], slicers=[field("[S].[T]", members=["[S].[T].&[1]"])]))
    assert text.endswith("WHERE ([S].[T].&[1]) CELL PROPERTIES VALUE, FORMAT_STRING, FORMATTED_VALUE")
    two = mdx.build_query("Sales", layout(slicers=[field("[S].[T]", members=["[S].[T].&[1]", "[S].[T].&[2]"])]))
    assert "(SELECT ({[S].[T].&[1], [S].[T].&[2]}) ON COLUMNS FROM [Sales])" in two


def test_slicer_field_is_in_one_area_only_and_only_selects():
    with pytest.raises(CubeError) as info:
        layout(rows=[field("[S].[T]")], slicers=[field("[S].[T]")])
    assert "more than one area" in info.value.message
    sliced = layout(slicers=[field("[S].[T]", sort={"by": "caption", "order": "asc"}, labelFilter={"operator": "contains", "value": "x"})])
    assert sliced.slicers[0].sort is None and sliced.slicers[0].label_filter is None


def test_sample_cube_slicer(demo):
    economy = member(demo, SERVICE, "Economy")
    sliced = run(demo, rows=[placed(demo, CITY)], slicers=[dict(placed(demo, SERVICE), members=[economy])])
    plain = run(demo, rows=[placed(demo, CITY)], filters=[dict(placed(demo, SERVICE), members=[economy])])
    assert sliced["values"] == plain["values"]
    assert sliced["filterRows"] == [] and plain["filterRows"] != []


# -- one level of a user hierarchy -------------------------------------------------------------
def test_placed_level_lists_that_levels_members():
    text = mdx.build_query("Sales", layout(rows=[field(DATE, level=MONTH_LEVEL)]))
    assert "NON EMPTY {{[Date].[Calendar].[Month].Members}, {[Date].[Calendar].[All]}} ON ROWS" in text
    filtered = mdx.build_query("Sales", layout(rows=[field(DATE, level=MONTH_LEVEL, labelFilter={"operator": "contains", "value": "jan"})]))
    assert "Filter({[Date].[Calendar].[Month].Members}, InStr(1, LCase(" in filtered


def test_level_must_belong_to_its_field():
    with pytest.raises(CubeError) as info:
        layout(rows=[field(DATE, level={"uniqueName": "[Geo].[City].[City]", "number": 1, "caption": "City"})])
    assert "another field" in info.value.message


def test_members_of_a_level():
    text = mdx.members_query("Sales", DATE, f"{DATE}.[All]", None, None, 10, f"{DATE}.[Month]")
    assert "Head({[Date].[Calendar].[Month].Members}, 11)" in text
    with pytest.raises(CubeError):
        mdx.members_query("Sales", DATE, None, None, None, 10, "[Geo].[City].[City]")


def test_sample_cube_months_alone(demo):
    fields = demo.fields(CATALOG, CUBE)
    calendar = next(h for d in fields["dimensions"] for h in d["hierarchies"] if h["uniqueName"] == DATE)
    assert [lv["caption"] for lv in calendar["levels"]] == ["Year", "Quarter", "Month"]
    months = run(demo, rows=[dict(placed(demo, DATE), level=MONTH_LEVEL)])
    years = run(demo, rows=[placed(demo, DATE)])
    assert len(months["rowHeaders"]) == 21 + 1  # 21 months of facts, and the grand total
    assert months["rowHeaders"][0][0]["indent"] == 0 and months["rowHeaders"][0][0]["text"] == "Jan"
    assert months["values"][-1] == years["values"][-1]
    jan = member(demo, DATE, "Jan", level="[Date].[Calendar].[Month]")
    assert jan.startswith("[Date].[Calendar].[Month].")


# -- Include filtered items in totals ---------------------------------------------------------
def test_non_visual_totals_move_the_filters_onto_the_axis():
    selected = field("[Geo].[City]", members=["[Geo].[City].&[1]"], labelFilter={"operator": "contains", "value": "a"},
                     valueFilter={"operator": "top", "measure": "[Measures].[Rides]", "value": 3})
    visual = mdx.build_query("Sales", layout(rows=[selected]))
    non_visual = mdx.build_query("Sales", layout(rows=[selected], includeFilteredItems=True))
    assert "(SELECT" in visual and "(SELECT" not in non_visual
    assert "FROM [Sales] " in non_visual
    assert "TopCount(Filter(Filter({[Geo].[City].[All].Children}, (Count(Intersect(Descendants(" in non_visual
    assert "{[Geo].[City].[All]}" in non_visual  # the grand total still comes from All


def test_non_visual_inner_value_filter_is_worked_out_for_each_outer_item():
    inner = field("[S].[T]", valueFilter={"operator": "top", "measure": "[Measures].[Rides]", "value": 1})
    text = mdx.build_query("Sales", layout(rows=[field("[Geo].[City]"), inner], includeFilteredItems=True))
    assert "Generate({[Geo].[City].[All].Children}, CrossJoin({[Geo].[City].CurrentMember}, TopCount(" in text


def test_sample_cube_non_visual_totals_count_every_item_and_are_starred(demo):
    tehran = member(demo, CITY, "Tehran")
    shown = dict(placed(demo, CITY), members=[tehran])
    visual = run(demo, rows=[shown])
    non_visual = run(demo, rows=[shown], includeFilteredItems=True)
    everything = run(demo, rows=[placed(demo, CITY)])
    assert labels(visual) == ["Tehran", "Grand Total"] and labels(non_visual) == ["Tehran", "Grand Total*"]
    assert visual["values"][-1] == visual["values"][0]
    assert non_visual["values"][-1] == everything["values"][-1]
    # Nothing filtered: no asterisk.
    assert labels(run(demo, rows=[placed(demo, CITY)], includeFilteredItems=True))[-1] == "Grand Total"


def test_sample_cube_non_visual_top_per_outer_item(demo):
    inner = dict(placed(demo, SERVICE), valueFilter={"operator": "top", "measure": "[Measures].[Rides]", "value": 1})
    grid = run(demo, rows=[placed(demo, CITY), inner], includeFilteredItems=True, subtotals=True)
    kinds = grid["rowKinds"]
    data_per_city = []
    count = 0
    for kind in kinds:
        if kind == "data":
            count += 1
        elif kind == "subtotal":
            data_per_city.append(count)
            count = 0
    assert data_per_city and all(n == 1 for n in data_per_city)


# -- named sets spanning several fields --------------------------------------------------------
def multi_set():
    return field("[Geo].[City]", set="[Pairs]", caption="Pairs", setHierarchies=["[Geo].[City]", "[S].[T]"],
                 setAllMembers=["[Geo].[City].[All]", "[S].[T].[All]"])


def test_multi_field_set_totals_with_every_all_member():
    text = mdx.build_query("Sales", layout(rows=[multi_set()]))
    assert "NON EMPTY {{[Pairs]}, {([Geo].[City].[All], [S].[T].[All])}} ON ROWS" in text
    with pytest.raises(CubeError) as info:
        layout(rows=[multi_set()], columns=[field("[S].[T]")])
    assert "more than one area" in info.value.message


def test_multi_field_set_inside_another_field_uses_its_tuple_as_current_item():
    sort = {"by": "measure", "measure": "[Measures].[Rides]", "order": "desc"}
    text = mdx.build_query("Sales", layout(rows=[multi_set(), field("[D].[Y]", sort=sort)]))
    assert "CrossJoin({([Geo].[City].CurrentMember, [S].[T].CurrentMember)}," in text


def test_sample_cube_multi_field_set(demo):
    pairs = next(s for s in demo.fields(CATALOG, CUBE)["sets"] if s["uniqueName"] == "[Top Service Cities]")
    assert pairs["hierarchies"] == [CITY, SERVICE]
    the_set = dict(placed(demo, CITY), set=pairs["uniqueName"], caption=pairs["caption"],
                   setHierarchies=pairs["hierarchies"], setAllMembers=pairs["allMembers"])
    grid = run(demo, rows=[the_set], columns=[placed(demo, "[Customer].[Segment]")])
    assert grid["rowFieldCaptions"] == ["Top Service Cities", "Service Type"]
    assert grid["rowFieldOf"] == [0, 0] and grid["colFieldOf"] == [0]
    assert len(grid["rowHeaders"]) == 5 + 1
    first = grid["rowHeaders"][0]
    assert first[0]["member"]["hierarchy"] == CITY and first[1]["member"]["hierarchy"] == SERVICE
    assert grid["rowHeaders"][-1][0]["text"] == "Grand Total"
    # Visual totals: the grand total is the five pairs' sum.
    total_col = len(grid["values"][0]) - 1
    assert grid["values"][-1][total_col] == sum(line[total_col] for line in grid["values"][:-1])
