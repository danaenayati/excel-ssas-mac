import math

import pytest

from cubebridge.errors import CubeError
from cubebridge.pivot.formats import to_excel
from cubebridge.pivot.grid import build_grid
from cubebridge.pivot.model import AxisMember, Cell, CellSet, Layout

CITY = "[Geo].[City]"
DATE = "[Date].[Calendar]"
SERVICE = "[Service].[Type]"


def m(hierarchy, key, caption, level=1, children=0):
    return AxisMember(f"{hierarchy}.&[{key}]", caption, hierarchy, level, children)


def all_(hierarchy):
    return AxisMember(f"{hierarchy}.[All]", "All", hierarchy, 0, 5)


def measure(name):
    return AxisMember(f"[Measures].[{name}]", name, "[Measures]", 0, 0)


def field(hierarchy, caption, expanded=(), members=(), member_captions=(), all_member=True):
    return {
        "hierarchy": hierarchy,
        "caption": caption,
        "allMember": f"{hierarchy}.[All]" if all_member else None,
        "expanded": list(expanded),
        "members": list(members),
        "memberCaptions": list(member_captions),
    }


def layout(rows=(), columns=(), values=("Rides",), filters=()):
    return Layout.from_json(
        {
            "rows": list(rows),
            "columns": list(columns),
            "values": [{"measure": f"[Measures].[{v}]", "caption": v} for v in values],
            "filters": list(filters),
        }
    )


def cells(*values, fmt="#,##0"):
    return {i: Cell(v, fmt, str(v)) for i, v in enumerate(values) if v is not None}


def texts(line):
    return [c["text"] for c in line]


# -- formats ---------------------------------------------------------------------
@pytest.mark.parametrize(
    "ssas, excel",
    [
        ("", "General"),
        (None, "General"),
        ("General Number", "General"),
        ("Standard", "#,##0.00"),
        ("Fixed", "0.00"),
        ("Currency", "$#,##0.00"),
        ("Percent", "0.00%"),
        ("percent", "0.00%"),
        ("Scientific", "0.00E+00"),
        ("Short Date", "yyyy-mm-dd"),
        ("#,##0", "#,##0"),
        ("#,##0.00;-#,##0.00", "#,##0.00;-#,##0.00"),
        ("0.0%", "0.0%"),
        ('#,##0 "rides"', '#,##0 "rides"'),
        ("#,##0;(#,##0);0;\\N\\A", "#,##0;(#,##0);0"),
        ("$#,##0.00", "$#,##0.00"),
    ],
)
def test_format_conversion(ssas, excel):
    assert to_excel(ssas) == excel


@pytest.mark.parametrize("ssas", ["line\nbreak", "#,##0 units", "x" * 300, 'unbalanced "quote', "yyyy-MM-dd hh:nn"])
def test_formats_excel_might_reject_fall_back_to_general(ssas):
    assert to_excel(ssas) == "General"


# -- grid ------------------------------------------------------------------------
def test_one_row_field_with_grand_total():
    cs = CellSet(
        axes=[
            [(measure("Rides"),)],
            [(m(CITY, 1, "Tehran"),), (m(CITY, 2, "Shiraz"),), (all_(CITY),)],
        ],
        cells=cells(120, 80, 200),
    )
    grid = build_grid(cs, layout(rows=[field(CITY, "City")]), elapsed_ms=12, mdx="SELECT …")
    assert grid["rowFieldCaptions"] == ["City"]
    assert [texts(line) for line in grid["colHeaderRows"]] == [["Rides"]]
    assert grid["colHeaderRows"][0][0]["kind"] == "measure"
    assert [texts(line) for line in grid["rowHeaders"]] == [["Tehran"], ["Shiraz"], ["Grand Total"]]
    assert grid["rowKinds"] == ["data", "data", "grandTotal"]
    assert grid["colKinds"] == ["data"]
    assert grid["values"] == [[120], [80], [200]]
    assert grid["formatPalette"] == ["General", "#,##0"]
    assert grid["cellFormats"] == [[1], [1], [1]]
    assert grid["stats"] == {"rows": 3, "columns": 1, "elapsedMs": 12, "mdx": "SELECT …"}
    tehran = grid["rowHeaders"][0][0]
    assert tehran["kind"] == "member"
    assert tehran["member"] == {
        "uniqueName": "[Geo].[City].&[1]",
        "hierarchy": CITY,
        "caption": "Tehran",
        "level": 1,
        "hasChildren": False,
        "expanded": False,
    }
    assert grid["rowHeaders"][2][0]["member"] is None


def test_multi_level_hierarchy_is_indented_and_marks_expansion():
    year = AxisMember("[Date].[Calendar].[Year].&[2025]", "2025", DATE, 1, 4)
    q1 = AxisMember("[Date].[Calendar].[Quarter].&[2025]&[1]", "Q1", DATE, 2, 3)
    jan = AxisMember("[Date].[Calendar].[Month].&[2025]&[1]", "Jan", DATE, 3, 0)
    cs = CellSet(axes=[[(measure("Rides"),)], [(year,), (q1,), (jan,)]], cells=cells(10, 4, 1))
    grid = build_grid(cs, layout(rows=[field(DATE, "Calendar", expanded=[year.unique_name, q1.unique_name])]))
    assert [(c[0]["text"], c[0]["indent"]) for c in grid["rowHeaders"]] == [("2025", 0), ("Q1", 1), ("Jan", 2)]
    assert [c[0]["member"]["expanded"] for c in grid["rowHeaders"]] == [True, True, False]
    assert [c[0]["member"]["hasChildren"] for c in grid["rowHeaders"]] == [True, True, False]


def test_level_numbering_without_an_all_member_starts_at_zero():
    actual = AxisMember("[Scenario].[Scenario].&[1]", "Actual", "[Scenario].[Scenario]", 0, 0)
    cs = CellSet(axes=[[(measure("Rides"),)], [(actual,)]], cells=cells(5))
    grid = build_grid(cs, layout(rows=[field("[Scenario].[Scenario]", "Scenario", all_member=False)]))
    assert grid["rowHeaders"][0][0]["indent"] == 0


def test_two_row_fields_blank_repeats_and_label_subtotals():
    tehran, shiraz = m(CITY, 1, "Tehran"), m(CITY, 2, "Shiraz")
    eco, bike = m(SERVICE, 1, "Economy"), m(SERVICE, 2, "Bike")
    cs = CellSet(
        axes=[
            [(measure("Rides"),)],
            [
                (tehran, eco),
                (tehran, bike),
                (tehran, all_(SERVICE)),
                (shiraz, eco),
                (shiraz, all_(SERVICE)),
                (all_(CITY), all_(SERVICE)),
            ],
        ],
        cells=cells(70, 50, 120, 80, 80, 200),
    )
    grid = build_grid(cs, layout(rows=[field(CITY, "City"), field(SERVICE, "Service")]))
    assert grid["rowFieldCaptions"] == ["City", "Service"]
    assert [texts(line) for line in grid["rowHeaders"]] == [
        ["Tehran", "Economy"],
        ["", "Bike"],
        ["Tehran Total", ""],
        ["Shiraz", "Economy"],
        ["Shiraz Total", ""],
        ["Grand Total", ""],
    ]
    assert grid["rowKinds"] == ["data", "data", "subtotal", "data", "subtotal", "grandTotal"]
    kinds = [[c["kind"] for c in line] for line in grid["rowHeaders"]]
    assert kinds[1] == ["blank", "member"]
    assert kinds[2] == ["subtotal", "subtotal"]
    assert kinds[5] == ["grandTotal", "grandTotal"]
    # a suppressed label still knows its member, so the cell can be drilled
    assert grid["rowHeaders"][1][0]["member"]["caption"] == "Tehran"


def test_column_headers_have_one_row_per_field_plus_measures():
    y25, y26 = m(DATE, 2025, "2025"), m(DATE, 2026, "2026")
    cols = [
        (y25, measure("Rides")),
        (y25, measure("Revenue")),
        (y26, measure("Rides")),
        (y26, measure("Revenue")),
        (all_(DATE), measure("Rides")),
        (all_(DATE), measure("Revenue")),
    ]
    cs = CellSet(axes=[cols, [(m(CITY, 1, "Tehran"),)]], cells=cells(1, 2, 3, 4, 4, 6))
    grid = build_grid(cs, layout(rows=[field(CITY, "City")], columns=[field(DATE, "Calendar")], values=("Rides", "Revenue")))
    assert [texts(line) for line in grid["colHeaderRows"]] == [
        ["2025", "", "2026", "", "Grand Total", ""],
        ["Rides", "Revenue", "Rides", "Revenue", "Rides", "Revenue"],
    ]
    assert grid["colKinds"] == ["data", "data", "data", "data", "grandTotal", "grandTotal"]
    assert grid["values"] == [[1, 2, 3, 4, 4, 6]]


def test_no_row_fields_gives_a_single_total_row():
    cs = CellSet(axes=[[(measure("Rides"),), (measure("Revenue"),)]], cells=cells(200, 5000.5))
    grid = build_grid(cs, layout(values=("Rides", "Revenue")))
    assert grid["rowFieldCaptions"] == [""]
    assert [texts(line) for line in grid["rowHeaders"]] == [["Total"]]
    assert grid["values"] == [[200, 5000.5]]
    assert grid["rowKinds"] == ["data"]


def test_empty_axis_means_no_data_rows():
    cs = CellSet(axes=[[(measure("Rides"),)], []], cells={})
    grid = build_grid(cs, layout(rows=[field(CITY, "City")]))
    assert grid["rowHeaders"] == [] and grid["values"] == [] and grid["cellFormats"] == []
    assert [texts(line) for line in grid["colHeaderRows"]] == [["Rides"]]
    assert grid["stats"]["rows"] == 0

    nothing = build_grid(CellSet(axes=[[], []], cells={}), layout(rows=[field(CITY, "City")]))
    assert nothing["colHeaderRows"] == [[]] and nothing["colKinds"] == [] and nothing["rowHeaders"] == []


def test_sparse_cells_and_values_excel_cannot_hold():
    cs = CellSet(
        axes=[[(measure("Rides"),), (measure("Revenue"),)], [(m(CITY, 1, "Tehran"),), (m(CITY, 2, "Shiraz"),)]],
        cells={0: Cell(1, "#,##0", "1"), 3: Cell(math.inf, "Standard", "inf"), 2: Cell(math.nan, "", "nan")},
    )
    grid = build_grid(cs, layout(rows=[field(CITY, "City")], values=("Rides", "Revenue")))
    assert grid["values"] == [[1, None], ["#NUM!", "#NUM!"]]
    assert grid["formatPalette"] == ["General", "#,##0", "#,##0.00"]
    assert grid["cellFormats"] == [[1, 0], [0, 2]]


def test_filter_rows_describe_the_selection():
    cs = CellSet(axes=[[(measure("Rides"),)]], cells=cells(1))
    one = "[Service].[Type].&[1]"
    two = "[Service].[Type].&[2]"
    grid = build_grid(
        cs,
        layout(
            filters=[
                field(SERVICE, "Service", members=[one], member_captions=["Economy"]),
                field(DATE, "Calendar"),
                field(CITY, "City", members=[one.replace("Service", "Geo"), two.replace("Service", "Geo")], member_captions=["A", "B"]),
                field("[Seg].[Seg]", "Segment", members=["[Seg].[Seg].&[New Rider]"]),
            ]
        ),
    )
    assert grid["filterRows"] == [
        {"caption": "Service", "value": "Economy"},
        {"caption": "Calendar", "value": "All"},
        {"caption": "City", "value": "(Multiple Items)"},
        {"caption": "Segment", "value": "New Rider"},
    ]


def test_result_over_the_cell_limit_is_refused():
    cols = [(measure(f"M{i}"),) for i in range(10)]
    rows = [(m(CITY, i, f"C{i}"),) for i in range(11)]
    with pytest.raises(CubeError) as info:
        build_grid(CellSet(axes=[cols, rows], cells={}), layout(rows=[field(CITY, "City")]), cell_limit=100)
    assert info.value.category == "too-large"
    assert "110" in info.value.message
