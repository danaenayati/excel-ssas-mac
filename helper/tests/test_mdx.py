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


# -- sorting, label and value filters, Show Values As, Show Details ------------------
def with_extras(base, **extras):
    out = dict(base)
    out.update(extras)
    return out


def test_layout_validation_of_the_new_settings():
    good = with_extras(field(CITY), sort={"by": "measure", "measure": "[Measures].[Rides]", "order": "desc"},
                       labelFilter={"operator": "between", "value": "A", "value2": "M"},
                       valueFilter={"operator": "top", "measure": "[Measures].[Rides]", "value": 5, "topBy": "items"})
    placed = Layout.from_json({"rows": [good], "values": [RIDES]}).rows[0]
    assert placed.sort.by == "measure" and placed.label_filter.value2 == "M" and placed.value_filter.top_by == "items"
    assert placed.narrowed

    bad = [
        {"sort": {"by": "colour", "order": "asc"}},
        {"sort": {"by": "measure", "order": "asc"}},  # no measure
        {"labelFilter": {"operator": "like", "value": "x"}},
        {"labelFilter": {"operator": "between", "value": "Z"}},  # no upper bound
        {"labelFilter": {"operator": "between", "value": "Z", "value2": "A"}},
        {"labelFilter": {"operator": "equals", "value": "x" * 201}},
        {"valueFilter": {"operator": "greaterThan", "measure": "[Geo].[City].&[1]", "value": 1}},
        {"valueFilter": {"operator": "greaterThan", "measure": "[Measures].[Rides]", "value": "10"}},
        {"valueFilter": {"operator": "greaterThan", "measure": "[Measures].[Rides]", "value": float("nan")}},
        {"valueFilter": {"operator": "top", "measure": "[Measures].[Rides]", "value": 2.5}},
        {"valueFilter": {"operator": "top", "measure": "[Measures].[Rides]", "value": 150, "topBy": "percent"}},
        {"valueFilter": {"operator": "between", "measure": "[Measures].[Rides]", "value": 9, "value2": 1}},
    ]
    for extras in bad:
        with pytest.raises(CubeError):
            Layout.from_json({"rows": [with_extras(field(CITY), **extras)], "values": [RIDES]})


def test_page_filters_ignore_sorting_and_filters_other_than_item_selection():
    placed = Layout.from_json(
        {"filters": [with_extras(field(SERVICE), labelFilter={"operator": "equals", "value": "x"})], "values": [RIDES]}
    ).filters[0]
    assert placed.label_filter is None and not placed.narrowed


def test_same_measure_may_appear_twice_when_shown_two_ways():
    layout = Layout.from_json({"values": [RIDES, dict(RIDES, showAs="percentOfGrandTotal")]})
    assert [v.show_as for v in layout.values] == ["normal", "percentOfGrandTotal"]
    with pytest.raises(CubeError):
        Layout.from_json({"values": [RIDES, dict(RIDES)]})
    with pytest.raises(CubeError):
        Layout.from_json({"values": [dict(RIDES, showAs="percentOfEverything")]})


def test_sort_by_caption_and_by_measure_wraps_the_detail_set_in_order():
    text = query(rows=[with_extras(field(CITY), sort={"by": "caption", "order": "asc"})])
    assert "NON EMPTY {Order({[Geo].[City].[All].Children}, [Geo].[City].CurrentMember.Member_Caption, ASC), {[Geo].[City].[All]}} ON ROWS" in text
    text = query(rows=[with_extras(field(DATE, expanded=["[Date].[Calendar].[Year].&[2025]"]), sort={"by": "measure", "measure": "[Measures].[Rides]", "order": "desc"})])
    assert ("Order(DrilldownMember({[Date].[Calendar].[All].Children}, {[Date].[Calendar].[Year].&[2025]}, RECURSIVE), "
            "[Measures].[Rides], DESC)") in text


@pytest.mark.parametrize(
    "operator, value, value2, condition",
    [
        ("equals", "Tehran", None, 'LCase([Geo].[City].CurrentMember.Member_Caption) = "tehran"'),
        ("notEquals", "Tehran", None, 'LCase([Geo].[City].CurrentMember.Member_Caption) <> "tehran"'),
        ("beginsWith", "Te", None, 'InStr(1, LCase([Geo].[City].CurrentMember.Member_Caption), "te") = 1'),
        ("notBeginsWith", "Te", None, 'InStr(1, LCase([Geo].[City].CurrentMember.Member_Caption), "te") <> 1'),
        ("endsWith", "ran", None, 'Right(LCase([Geo].[City].CurrentMember.Member_Caption), 3) = "ran"'),
        ("notEndsWith", "ran", None, 'Right(LCase([Geo].[City].CurrentMember.Member_Caption), 3) <> "ran"'),
        ("contains", 'a"b', None, 'InStr(1, LCase([Geo].[City].CurrentMember.Member_Caption), "a""b") > 0'),
        ("notContains", "x", None, 'InStr(1, LCase([Geo].[City].CurrentMember.Member_Caption), "x") = 0'),
        ("greaterThan", "M", None, 'LCase([Geo].[City].CurrentMember.Member_Caption) > "m"'),
        ("lessOrEqual", "M", None, 'LCase([Geo].[City].CurrentMember.Member_Caption) <= "m"'),
        ("between", "A", "M", '(LCase([Geo].[City].CurrentMember.Member_Caption) >= "a" AND LCase([Geo].[City].CurrentMember.Member_Caption) <= "m")'),
        ("notBetween", "A", "M", '(LCase([Geo].[City].CurrentMember.Member_Caption) < "a" OR LCase([Geo].[City].CurrentMember.Member_Caption) > "m")'),
    ],
)
def test_label_filters_become_a_subselect_on_the_fields_top_level(operator, value, value2, condition):
    label = {"operator": operator, "value": value}
    if value2 is not None:
        label["value2"] = value2
    text = query(rows=[with_extras(field(CITY), labelFilter=label)])
    assert text.endswith(f"FROM (SELECT {{Filter({{[Geo].[City].[All].Children}}, {condition})}} ON COLUMNS FROM [Sales])" + TAIL)


@pytest.mark.parametrize(
    "vf, expression",
    [
        ({"operator": "greaterThan", "value": 1000}, "Filter({[Geo].[City].[All].Children}, [Measures].[Rides] > 1000)"),
        ({"operator": "lessThan", "value": 0.25}, "Filter({[Geo].[City].[All].Children}, [Measures].[Rides] < 0.25)"),
        ({"operator": "between", "value": 10, "value2": 20}, "Filter({[Geo].[City].[All].Children}, [Measures].[Rides] >= 10 AND [Measures].[Rides] <= 20)"),
        ({"operator": "notBetween", "value": 10, "value2": 20}, "Filter({[Geo].[City].[All].Children}, [Measures].[Rides] < 10 OR [Measures].[Rides] > 20)"),
        ({"operator": "top", "value": 3}, "TopCount({[Geo].[City].[All].Children}, 3, [Measures].[Rides])"),
        ({"operator": "top", "value": 30, "topBy": "percent"}, "TopPercent({[Geo].[City].[All].Children}, 30, [Measures].[Rides])"),
        ({"operator": "top", "value": 5000, "topBy": "sum"}, "TopSum({[Geo].[City].[All].Children}, 5000, [Measures].[Rides])"),
        ({"operator": "bottom", "value": 2}, "BottomCount(NonEmpty({[Geo].[City].[All].Children}, {[Measures].[Rides]}), 2, [Measures].[Rides])"),
        ({"operator": "bottom", "value": 10, "topBy": "percent"}, "BottomPercent(NonEmpty({[Geo].[City].[All].Children}, {[Measures].[Rides]}), 10, [Measures].[Rides])"),
    ],
)
def test_value_filters_become_a_subselect(vf, expression):
    text = query(rows=[with_extras(field(CITY), valueFilter=dict(vf, measure="[Measures].[Rides]"))])
    assert text.endswith(f"FROM (SELECT {{{expression}}} ON COLUMNS FROM [Sales])" + TAIL)


def test_label_then_value_filter_on_one_field_and_item_selection_innermost():
    text = query(
        rows=[
            with_extras(
                field(CITY, members=["[Geo].[City].&[1]", "[Geo].[City].&[2]"]),
                labelFilter={"operator": "contains", "value": "a"},
                valueFilter={"operator": "top", "measure": "[Measures].[Rides]", "value": 1},
            )
        ]
    )
    assert text.endswith(
        "FROM (SELECT {TopCount(Filter({[Geo].[City].[All].Children}, "
        'InStr(1, LCase([Geo].[City].CurrentMember.Member_Caption), "a") > 0), 1, [Measures].[Rides])} ON COLUMNS '
        "FROM (SELECT ({[Geo].[City].&[1], [Geo].[City].&[2]}) ON COLUMNS FROM [Sales]))" + TAIL
    )


def test_with_a_value_filter_single_page_filters_move_into_the_subselects():
    # A value filter is evaluated in the sub-select's space, which a WHERE slicer does
    # not reach, so the page filter must be applied inside it.
    text = query(
        rows=[with_extras(field(CITY), valueFilter={"operator": "top", "measure": "[Measures].[Rides]", "value": 3})],
        filters=[field(SERVICE, members=["[Service].[Type].&[1]"])],
    )
    assert "WHERE" not in text
    assert text.endswith(
        "FROM (SELECT {TopCount({[Geo].[City].[All].Children}, 3, [Measures].[Rides])} ON COLUMNS "
        "FROM (SELECT ({[Service].[Type].&[1]}) ON COLUMNS FROM [Sales]))" + TAIL
    )


def test_numbers_are_written_without_exponents():
    text = query(rows=[with_extras(field(CITY), valueFilter={"operator": "greaterThan", "measure": "[Measures].[Rides]", "value": 1e-7})])
    assert "[Measures].[Rides] > 0.0000001" in text
    text = query(rows=[with_extras(field(CITY), valueFilter={"operator": "greaterThan", "measure": "[Measures].[Rides]", "value": 2e20})])
    assert "[Measures].[Rides] > 200000000000000000000" in text


def test_show_values_as_defines_calculated_measures():
    text = query(
        rows=[field(CITY)],
        columns=[field(SERVICE)],
        values=[RIDES, dict(RIDES, showAs="percentOfGrandTotal"), dict(REVENUE, showAs="percentOfColumnTotal"), dict(RIDES, showAs="percentOfRowTotal")],
    )
    assert text.startswith(
        "WITH "
        "MEMBER [Measures].[Rides (% of Grand Total)] AS "
        "IIF(([Measures].[Rides], [Geo].[City].[All], [Service].[Type].[All]) = 0, NULL, "
        "[Measures].[Rides] / ([Measures].[Rides], [Geo].[City].[All], [Service].[Type].[All])), FORMAT_STRING = \"Percent\" "
        "MEMBER [Measures].[Revenue (% of Column Total)] AS "
        "IIF(([Measures].[Revenue], [Geo].[City].[All]) = 0, NULL, [Measures].[Revenue] / ([Measures].[Revenue], [Geo].[City].[All])), "
        "FORMAT_STRING = \"Percent\" "
        "MEMBER [Measures].[Rides (% of Row Total)] AS "
        "IIF(([Measures].[Rides], [Service].[Type].[All]) = 0, NULL, [Measures].[Rides] / ([Measures].[Rides], [Service].[Type].[All])), "
        "FORMAT_STRING = \"Percent\" "
        "SELECT "
    )
    assert ("{[Measures].[Rides], [Measures].[Rides (% of Grand Total)], [Measures].[Revenue (% of Column Total)], "
            "[Measures].[Rides (% of Row Total)]}") in text


def test_show_values_as_with_nothing_on_the_other_axis_divides_by_itself():
    text = query(rows=[field(CITY)], values=[dict(RIDES, showAs="percentOfRowTotal")])
    assert "IIF(([Measures].[Rides]) = 0, NULL, [Measures].[Rides] / ([Measures].[Rides]))" in text


# -- Show Details ------------------------------------------------------------------
def details(layout_json, members, measure="[Measures].[Rides]", max_rows=1000):
    return mdx.details_query("Sales", Layout.from_json(layout_json), members, measure, max_rows)


def test_details_slices_on_the_cell_and_single_page_filters():
    text = details(
        {"rows": [field(CITY)], "columns": [field(SERVICE)], "filters": [field(DATE, members=["[Date].[Calendar].[Year].&[2025]"])], "values": [RIDES]},
        ["[Geo].[City].&[1]", "[Service].[Type].&[2]"],
    )
    assert text == (
        "DRILLTHROUGH MAXROWS 1001 SELECT FROM [Sales] "
        "WHERE ([Measures].[Rides], [Geo].[City].&[1], [Service].[Type].&[2], [Date].[Calendar].[Year].&[2025])"
    )


def test_details_for_a_total_is_allowed_when_that_field_is_not_filtered():
    text = details({"rows": [field(CITY)], "values": [RIDES]}, [])
    assert text == "DRILLTHROUGH MAXROWS 1001 SELECT FROM [Sales] WHERE ([Measures].[Rides])"


@pytest.mark.parametrize(
    "layout_json, members",
    [
        ({"filters": [field(SERVICE, members=["[Service].[Type].&[1]", "[Service].[Type].&[2]"])], "values": [RIDES]}, []),
        ({"rows": [field(CITY, members=["[Geo].[City].&[1]", "[Geo].[City].&[2]"])], "values": [RIDES]}, []),
        ({"rows": [with_extras(field(CITY), labelFilter={"operator": "contains", "value": "a"})], "values": [RIDES]}, []),
    ],
)
def test_details_are_refused_where_a_filter_covers_several_items(layout_json, members):
    with pytest.raises(CubeError) as info:
        details(layout_json, members)
    assert info.value.category == "bad-request" and "Show Details" in info.value.message


def test_details_validate_names_and_the_measure():
    with pytest.raises(CubeError):
        details({"values": [RIDES]}, ["[Geo].[City].&[1]) FROM [x]"])
    with pytest.raises(CubeError):
        details({"values": [RIDES]}, [], measure="[Geo].[City].&[1]")
    with pytest.raises(CubeError):
        details({"values": [RIDES]}, ["[Geo].[City].&[1]", "[Geo].[City].&[2]"])  # two members of one hierarchy
