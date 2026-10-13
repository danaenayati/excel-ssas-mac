import pytest

from cubebridge.errors import CubeError
from cubebridge.pivot import mdx
from cubebridge.pivot.model import Layout
from cubebridge.pivot.names import quote_identifier, quote_string, validate_unique_name

TAIL = " CELL PROPERTIES VALUE, FORMAT_STRING, FORMATTED_VALUE"
RIDES = {"measure": "[Measures].[Rides]", "caption": "Rides"}
REVENUE = {"measure": "[Measures].[Revenue]", "caption": "Revenue"}


def field(hierarchy, all_member="auto", expanded=(), members=()):
    return {
        "hierarchy": hierarchy,
        "caption": hierarchy.split(".")[-1].strip("[]"),
        "allMember": f"{hierarchy}.[All]" if all_member == "auto" else all_member,
        "expanded": list(expanded),
        "members": list(members),
        "memberCaptions": [],
    }


def query(rows=(), columns=(), values=(RIDES,), filters=(), **options):
    layout = Layout.from_json(
        {"rows": list(rows), "columns": list(columns), "values": list(values), "filters": list(filters), "options": options}
    )
    return mdx.build_query("Sales", layout)


CITY = "[Geo].[City]"
DATE = "[Date].[Calendar]"
SERVICE = "[Service].[Type]"


# -- names -----------------------------------------------------------------------
@pytest.mark.parametrize(
    "name",
    [
        "[Measures].[Rides]",
        "[Date].[Calendar].[Year].&[2025]",
        "[Geo].[City].&[10]&[20]",
        "[Geo].[City].[All]",
        "[Product].[Name].&[A ]] bracket]",
        "[شهر].[تهران]",
        "[Measures]",
    ],
)
def test_real_unique_names_are_accepted(name):
    assert validate_unique_name(name) == name


@pytest.mark.parametrize(
    "name",
    [
        "",
        None,
        42,
        "Measures.Rides",
        "[A].[B]} ON 0 FROM [X]; DROP",
        "[A].[B] , [C]",
        "[A].[B",
        "[A]]",
        "[A].Children",
        "[A].[B]\n.[C]",
        "[A]" + ".[B]" * 400,
        "{[A]}",
        "([A].[B])",
    ],
)
def test_anything_else_is_rejected(name):
    with pytest.raises(CubeError) as info:
        validate_unique_name(name)
    assert info.value.category == "bad-request"


def test_quoting():
    assert quote_identifier("Sales [EU]") == "[Sales [EU]]]"
    assert quote_string('say "hi"') == '"say ""hi"""'


# -- layout validation -----------------------------------------------------------
def test_a_hierarchy_cannot_be_in_two_areas():
    with pytest.raises(CubeError):
        query(rows=[field(CITY)], filters=[field(CITY)])


def test_measures_cannot_be_placed_on_an_axis_and_only_measures_are_values():
    with pytest.raises(CubeError):
        query(rows=[field("[Measures]", all_member=None)])
    with pytest.raises(CubeError):
        query(values=[{"measure": "[Geo].[City].&[1]", "caption": "x"}])


def test_injection_through_a_member_name_is_rejected_before_any_mdx_is_built():
    with pytest.raises(CubeError):
        query(rows=[field(CITY, members=["[Geo].[City].&[1]} ON 0 FROM [Sales]) --"])])


# -- query generation ------------------------------------------------------------
def test_measures_only():
    assert query() == "SELECT NON EMPTY {[Measures].[Rides]} ON COLUMNS FROM [Sales]" + TAIL


def test_no_values_uses_the_default_measure():
    assert query(values=[]) == "SELECT NON EMPTY {[Measures].DefaultMember} ON COLUMNS FROM [Sales]" + TAIL


def test_one_row_hierarchy_with_grand_total():
    assert query(rows=[field(CITY)]) == (
        "SELECT NON EMPTY {[Measures].[Rides]} ON COLUMNS, "
        "NON EMPTY {{[Geo].[City].[All].Children}, {[Geo].[City].[All]}} ON ROWS "
        "FROM [Sales]" + TAIL
    )


def test_rows_columns_and_two_measures():
    assert query(rows=[field(CITY)], columns=[field(DATE)], values=[RIDES, REVENUE]) == (
        "SELECT NON EMPTY CrossJoin({{[Date].[Calendar].[All].Children}, {[Date].[Calendar].[All]}}, "
        "{[Measures].[Rides], [Measures].[Revenue]}) ON COLUMNS, "
        "NON EMPTY {{[Geo].[City].[All].Children}, {[Geo].[City].[All]}} ON ROWS "
        "FROM [Sales]" + TAIL
    )


def test_expanded_members_are_drilled_recursively():
    expanded = ["[Date].[Calendar].[Year].&[2025]", "[Date].[Calendar].[Quarter].&[2025]&[1]"]
    assert query(rows=[field(DATE, expanded=expanded)], grandTotals=False) == (
        "SELECT NON EMPTY {[Measures].[Rides]} ON COLUMNS, "
        "NON EMPTY DrilldownMember({[Date].[Calendar].[All].Children}, "
        "{[Date].[Calendar].[Year].&[2025], [Date].[Calendar].[Quarter].&[2025]&[1]}, RECURSIVE) ON ROWS "
        "FROM [Sales]" + TAIL
    )


def test_grand_totals_off():
    assert "ON ROWS" in query(rows=[field(CITY)], grandTotals=False)
    assert "NON EMPTY {[Geo].[City].[All].Children} ON ROWS" in query(rows=[field(CITY)], grandTotals=False)


def test_two_row_hierarchies_without_subtotals_have_only_the_grand_total():
    assert (
        "NON EMPTY {CrossJoin({[Geo].[City].[All].Children}, {[Service].[Type].[All].Children}), "
        "{([Geo].[City].[All], [Service].[Type].[All])}} ON ROWS"
    ) in query(rows=[field(CITY), field(SERVICE)])


def test_three_row_hierarchies_with_subtotals_nest_the_all_tuples():
    text = query(rows=[field("[A].[A]"), field("[B].[B]"), field("[C].[C]")], subtotals=True)
    assert (
        "NON EMPTY {CrossJoin({[A].[A].[All].Children}, "
        "{CrossJoin({[B].[B].[All].Children}, {{[C].[C].[All].Children}, {[C].[C].[All]}}), "
        "{([B].[B].[All], [C].[C].[All])}}), "
        "{([A].[A].[All], [B].[B].[All], [C].[C].[All])}} ON ROWS"
    ) in text


def test_hierarchy_without_all_member_has_no_total():
    text = query(rows=[field("[Scenario].[Scenario]", all_member=None)])
    assert "NON EMPTY {[Scenario].[Scenario].Levels(0).Members} ON ROWS" in text


def test_total_tuples_needing_a_missing_all_member_are_dropped():
    text = query(rows=[field(CITY), field("[Scenario].[Scenario]", all_member=None)], subtotals=True)
    assert "NON EMPTY CrossJoin({[Geo].[City].[All].Children}, {[Scenario].[Scenario].Levels(0).Members}) ON ROWS" in text


def test_single_member_filter_goes_in_where():
    text = query(rows=[field(CITY)], filters=[field(SERVICE, members=["[Service].[Type].&[1]"])])
    assert text.endswith("FROM [Sales] WHERE ([Service].[Type].&[1])" + TAIL)


def test_multi_member_filter_becomes_a_subselect():
    members = ["[Service].[Type].&[1]", "[Service].[Type].&[2]"]
    text = query(rows=[field(CITY)], filters=[field(SERVICE, members=members)])
    assert text.endswith(
        "FROM (SELECT ({[Service].[Type].&[1], [Service].[Type].&[2]}) ON COLUMNS FROM [Sales])" + TAIL
    )
    assert "WHERE" not in text


def test_axis_field_selection_is_a_subselect_even_with_one_member():
    text = query(rows=[field(CITY, members=["[Geo].[City].&[1]"])])
    assert text.endswith("FROM (SELECT ({[Geo].[City].&[1]}) ON COLUMNS FROM [Sales])" + TAIL)


def test_several_selections_nest_and_where_is_kept():
    text = query(
        rows=[field(CITY, members=["[Geo].[City].&[1]"])],
        columns=[field(DATE, members=["[Date].[Calendar].[Year].&[2025]", "[Date].[Calendar].[Year].&[2026]"])],
        filters=[field(SERVICE, members=["[Service].[Type].&[1]"])],
    )
    assert text.endswith(
        "FROM (SELECT ({[Date].[Calendar].[Year].&[2025], [Date].[Calendar].[Year].&[2026]}) ON COLUMNS "
        "FROM (SELECT ({[Geo].[City].&[1]}) ON COLUMNS FROM [Sales])) "
        "WHERE ([Service].[Type].&[1])" + TAIL
    )


def test_non_empty_off():
    text = query(rows=[field(CITY)], nonEmpty=False)
    assert "NON EMPTY" not in text


def test_cube_name_is_bracket_quoted():
    layout = Layout.from_json({"values": [RIDES]})
    assert "FROM [Sales [EU]]]" in mdx.build_query("Sales [EU]", layout)


# -- member lists ----------------------------------------------------------------
def test_members_query_top_level_children_and_search():
    top = mdx.members_query("Sales", CITY, "[Geo].[City].[All]", parent=None, search=None, limit=500)
    assert top == "SELECT {} ON COLUMNS, Head({[Geo].[City].[All].Children}, 501) ON ROWS FROM [Sales]"

    no_all = mdx.members_query("Sales", "[S].[S]", None, parent=None, search=None, limit=10)
    assert no_all == "SELECT {} ON COLUMNS, Head({[S].[S].Levels(0).Members}, 11) ON ROWS FROM [Sales]"

    children = mdx.members_query("Sales", DATE, "[Date].[Calendar].[All]", parent="[Date].[Calendar].[Year].&[2025]", search=None, limit=500)
    assert children == "SELECT {} ON COLUMNS, Head({[Date].[Calendar].[Year].&[2025].Children}, 501) ON ROWS FROM [Sales]"

    search = mdx.members_query("Sales", CITY, "[Geo].[City].[All]", parent=None, search='Teh"ran', limit=50)
    assert search == (
        "SELECT {} ON COLUMNS, Head(Filter({[Geo].[City].AllMembers}, "
        'InStr(1, LCase([Geo].[City].CurrentMember.Member_Caption), "teh""ran") > 0), 51) ON ROWS FROM [Sales]'
    )


# -- review fixes --------------------------------------------------------------------
@pytest.mark.parametrize("name", ["[Measures]\n", "[D].[H]\n", "[A].[B]\r", "\n[A]"])
def test_a_trailing_or_leading_line_break_is_rejected(name):
    with pytest.raises(CubeError):
        validate_unique_name(name)


def test_filter_showing_all_slices_on_the_all_member_explicitly():
    # The sheet says "All", so the query must not fall back to the cube's default member.
    text = query(rows=[field(CITY)], filters=[field(SERVICE)])
    assert text.endswith("FROM [Sales] WHERE ([Service].[Type].[All])" + TAIL)


def test_filter_without_all_member_and_without_selection_adds_nothing():
    assert query(filters=[field("[Scenario].[Scenario]", all_member=None)]) == query()
