"""Cube formulas: batching, grouping, splitting and per-item answers (formulas.py)."""

import re

import pytest

from cubebridge import formulas
from cubebridge.errors import CubeError
from cubebridge.pivot.model import AxisMember, Cell, CellSet

YEAR_2025 = "[Date].[Year].&[2025]"
YEAR_2026 = "[Date].[Year].&[2026]"
TEHRAN = "[Geo].[City].&[1]"
SHIRAZ = "[Geo].[City].&[2]"
RIDES = "[Measures].[Rides]"
REVENUE = "[Measures].[Revenue]"

CAPTIONS = {
    YEAR_2025: "2025", YEAR_2026: "2026", TEHRAN: "Tehran", SHIRAZ: "Shiraz",
    RIDES: "Rides", REVENUE: "Revenue", "[Geo].[City].[All]": "All",
}
FACTS = {  # (year, city) -> (rides, revenue)
    ("2025", "Tehran"): (10, 100.5), ("2025", "Shiraz"): (4, 40.0),
    ("2026", "Tehran"): (7, 70.0),
}
_NAME = re.compile(r"\[(?:[^\]]|\]\])*\](?:\.\[(?:[^\]]|\]\])*\]|\.?&\[(?:[^\]]|\]\])*\])*")


class FakeCube:
    """Answers the query shapes formulas.py sends, from the small table above.

    `missing` is the cube's MDX Missing Member Mode: "error" rejects a query that names
    an unknown member; "ignore" drops the tuple from the set without a word.
    """

    def __init__(self, missing="error"):
        self.missing = missing
        self.queries: list[str] = []
        self.fail_with: CubeError | None = None

    def __call__(self, text: str) -> CellSet:
        self.queries.append(text)
        if self.fail_with:
            raise self.fail_with
        if text.startswith("WITH MEMBER [Measures].[__CubeExplorer Property]"):
            return self._properties(text)
        if text.startswith("WITH MEMBER [Measures].[__CubeExplorer Count"):
            return self._counts(text)
        if "Head(" in text:
            return self._head(text)
        if text.startswith("SELECT {} ON 0"):
            return self._members(text)
        return self._values(text)

    def _known(self, names):
        unknown = [n for n in names if n not in CAPTIONS]
        if unknown and self.missing == "error":
            raise CubeError("server", formulas.REJECTED, f"The member {unknown[0]} was not found in the cube.")
        return not unknown

    def _values(self, text):
        # The cube's default measure is Rides.
        inner = text[len("SELECT {") : text.index("} ON 0")].replace("[Measures].DefaultMember", RIDES)
        tuples = [_NAME.findall(t) for t in re.findall(r"\(([^()]*)\)", inner)]
        kept = [t for t in tuples if self._known(t)]
        cellset = CellSet(axes=[[tuple(AxisMember(n, CAPTIONS[n], "", 0) for n in t) for t in kept]])
        for i, names in enumerate(kept):
            value = self._value(names)
            if value is not None:
                cellset.cells[i] = Cell(value)
        return cellset

    def _value(self, names):
        year = next((CAPTIONS[n] for n in names if n.startswith("[Date]")), None)
        city = next((CAPTIONS[n] for n in names if n.startswith("[Geo]") and CAPTIONS[n] != "All"), None)
        measure = next((n for n in names if n.startswith("[Measures]")), RIDES)
        if measure == REVENUE and city == "Shiraz":
            return formulas.CELL_ERROR  # a cell that failed to calculate
        rows = [v for (y, c), v in FACTS.items() if year in (None, y) and city in (None, c)]
        if not rows:
            return None
        return sum(r[0] if measure == RIDES else r[1] for r in rows)

    def _members(self, text):
        names = _NAME.findall(text[text.index("{}") + 2 : text.index("ON 1")])
        kept = [n for n in names if self._known([n])]
        return CellSet(axes=[[], [(AxisMember(n, CAPTIONS[n], "", 1),) for n in kept]])

    def _properties(self, text):
        found = re.search(r'Properties\("([^"]*)"\)', text)
        prop = found.group(1) if found else ""
        names = _NAME.findall(text[text.index("ON 0, {") + 6 : text.index("ON 1")])
        kept = [n for n in names if self._known([n])]
        cellset = CellSet(axes=[[(AxisMember(formulas._PROPERTY_MEASURE, "p", "", 0),)], [(AxisMember(n, CAPTIONS[n], "", 1),) for n in kept]])
        for i, n in enumerate(kept):
            cellset.cells[i] = Cell(n.rsplit("&[", 1)[-1].rstrip("]") if prop.lower() == "key" else formulas.CELL_ERROR)
        return cellset

    def _counts(self, text):
        sets = re.findall(r"Count\(\{(.*?)\}\)(?: MEMBER| SELECT)", text)
        cellset = CellSet(axes=[[(AxisMember(f"c{i}", "c", "", 0),) for i in range(len(sets))]])
        for i, expression in enumerate(sets):
            if "Nope" in expression:
                raise CubeError("server", formulas.REJECTED, "Parser: The syntax for 'Nope' is incorrect.")
            cellset.cells[i] = Cell(len(_NAME.findall(expression)) if "{" in expression else 2)
        return cellset

    def _head(self, text):
        expression = re.search(r"Head\(\{(.*)\}, (\d+)\) ON 1", text)
        names = [TEHRAN, SHIRAZ] if ".Children" in expression.group(1) else _NAME.findall(expression.group(1))
        names = names[: int(expression.group(2))]
        return CellSet(axes=[[], [(AxisMember(n, CAPTIONS[n], "", 1),) for n in names]])


@pytest.fixture
def cube():
    return FakeCube()


def values(answers):
    return [a.get("value", a.get("error", {}).get("category")) for a in answers]


# -- names -------------------------------------------------------------------------
@pytest.mark.parametrize(
    "name, key",
    [
        ("[Date].[Calendar].[Year].&[2025]", "[date].[calendar]"),
        ("[Date].[Calendar].&[2025]", "[date].[calendar]"),
        ("[Date].&[2025]", "[date]"),
        ("[Measures].[Rides]", "[measures]"),
        ("[MEASURES].[x]", "[measures]"),
        ("[Odd]]Name].[H]", "[odd]]name].[h]"),
    ],
)
def test_hierarchy_key(name, key):
    assert formulas.hierarchy_key(name) == key


def test_prepare_tuple_orders_by_hierarchy_and_drops_repeats():
    assert formulas.prepare_tuple([RIDES, TEHRAN, YEAR_2025, " " + TEHRAN]) == (YEAR_2025, TEHRAN, RIDES)
    assert formulas.prepare_tuple([]) == ()


@pytest.mark.parametrize(
    "raw, needle",
    [
        (["2025"], "not a member's unique name"),
        (["[Date].[Year].&[2025]; DROP"], "not a member's unique name"),
        ([YEAR_2025, YEAR_2026], "same field"),
        ([RIDES, REVENUE], "same field"),
        ([f"[D{i}].[H].&[1]" for i in range(65)], "at most 64"),
    ],
)
def test_prepare_tuple_refusals_are_per_item_bad_requests(raw, needle):
    answer = formulas.prepare_tuple(raw)
    assert answer["error"]["category"] == "bad-request" and needle in answer["error"]["message"]


# -- CUBE.VALUE --------------------------------------------------------------------
def test_values_of_one_group_come_from_one_query(cube):
    tuples = [[YEAR_2025, TEHRAN, RIDES], [RIDES, SHIRAZ, YEAR_2025], [YEAR_2026, SHIRAZ, RIDES]]
    assert values(formulas.evaluate_cells("Sales", tuples, cube)) == [10, 4, None]
    assert cube.queries == [
        "SELECT {([Date].[Year].&[2025], [Geo].[City].&[1], [Measures].[Rides]), "
        "([Date].[Year].&[2025], [Geo].[City].&[2], [Measures].[Rides]), "
        "([Date].[Year].&[2026], [Geo].[City].&[2], [Measures].[Rides])} ON 0 FROM [Sales] CELL PROPERTIES VALUE"
    ]


def test_tuples_naming_other_hierarchies_go_in_separate_queries(cube):
    tuples = [[YEAR_2025], [TEHRAN, REVENUE], [REVENUE, YEAR_2026], [], [YEAR_2026]]
    assert values(formulas.evaluate_cells("Sales", tuples, cube)) == [14, 170.5, 70.0, 21, 7]
    assert len(cube.queries) == 4
    # No measure: the cube's default one, as in Excel.
    assert any("([Measures].DefaultMember)" in q for q in cube.queries)


def test_identical_tuples_are_asked_once(cube):
    answers = formulas.evaluate_cells("Sales", [[YEAR_2025, RIDES], [RIDES, YEAR_2025], [YEAR_2025.lower().replace("date", "Date"), RIDES]], cube)
    assert values(answers) == [14, 14, 14]
    assert len(cube.queries) == 1 and cube.queries[0].count("(") == 1  # one tuple


def test_a_bad_name_fails_only_its_own_cell_and_is_never_sent(cube):
    answers = formulas.evaluate_cells("Sales", [[YEAR_2025], ["Tehran"], [YEAR_2025, YEAR_2026]], cube)
    assert values(answers) == [14, "bad-request", "bad-request"]
    assert "Tehran" not in " ".join(cube.queries)


@pytest.mark.parametrize("mode", ["error", "ignore"])
def test_an_unknown_member_is_found_by_splitting(mode):
    cube = FakeCube(missing=mode)
    tuples = [[YEAR_2025, TEHRAN], [YEAR_2026, TEHRAN], ["[Date].[Year].&[2030]", TEHRAN], [YEAR_2025, SHIRAZ], [YEAR_2026, SHIRAZ]]
    answers = formulas.evaluate_cells("Sales", tuples, cube)
    assert values(answers) == [10, 7, "not-found", 4, None]
    assert len(cube.queries) <= 7  # 1 + halves, not one query per tuple
    if mode == "error":
        assert "2030" in answers[2]["error"]["detail"]


def test_a_cell_the_cube_could_not_calculate_is_an_error_of_that_cell(cube):
    answers = formulas.evaluate_cells("Sales", [[SHIRAZ, REVENUE], [TEHRAN, REVENUE]], cube)
    assert values(answers) == ["not-found", 170.5]
    assert "could not calculate" in answers[0]["error"]["message"]


@pytest.mark.parametrize(
    "error",
    [
        CubeError("unreachable", "The server cannot be reached."),
        CubeError("server", "Your account does not have access to this database or cube."),
        CubeError("cancelled", "The query was cancelled."),
    ],
)
def test_failures_that_are_not_about_one_item_fail_the_request(cube, error):
    cube.fail_with = error
    with pytest.raises(CubeError) as info:
        formulas.evaluate_cells("Sales", [[YEAR_2025], [YEAR_2026]], cube)
    assert info.value is error
    assert len(cube.queries) == 1  # not split


def test_splitting_stops_at_the_query_budget(cube, monkeypatch):
    monkeypatch.setattr(formulas, "MAX_QUERIES", 10)
    tuples = [[f"[Date].[Year].&[{1900 + i}]"] for i in range(50)]
    answers = formulas.evaluate_cells("Sales", tuples, cube)
    assert len(cube.queries) == 10
    assert all(a["error"]["category"] == "not-found" for a in answers)
    assert any("Too many" in a["error"]["message"] for a in answers)


def test_large_groups_are_sent_in_several_queries(cube, monkeypatch):
    monkeypatch.setattr(formulas, "MAX_PER_QUERY", 2)
    answers = formulas.evaluate_cells("Sales", [[YEAR_2025, TEHRAN], [YEAR_2025, SHIRAZ], [YEAR_2026, TEHRAN]], cube)
    assert values(answers) == [10, 4, 7] and len(cube.queries) == 2


def test_the_cube_name_is_quoted(cube):
    formulas.evaluate_cells("My ]Cube", [[YEAR_2025]], cube)
    assert "FROM [My ]]Cube]" in cube.queries[0]


# -- lookups -----------------------------------------------------------------------
def test_member_captions_and_tuples_show_their_last_member(cube):
    items = [
        {"kind": "member", "members": [YEAR_2025]},
        {"kind": "member", "members": [TEHRAN, YEAR_2026]},
        {"kind": "member", "members": [RIDES]},
        {"kind": "member", "members": ["[Geo].[City].&[99]"]},
        {"kind": "member", "members": []},
        {"kind": "member", "members": [YEAR_2025, YEAR_2026]},
    ]
    answers = formulas.evaluate_lookups("Sales", items, cube)
    assert values(answers) == ["2025", "2026", "Rides", "not-found", "bad-request", "bad-request"]
    member_queries = [q for q in cube.queries if q.startswith("SELECT {} ON 0, {")]
    assert member_queries[0] == "SELECT {} ON 0, {[Date].[Year].&[2025], [Date].[Year].&[2026]} ON 1 FROM [Sales]"


def test_member_properties(cube):
    items = [
        {"kind": "property", "member": TEHRAN, "property": "Key"},
        {"kind": "property", "member": SHIRAZ, "property": "key"},
        {"kind": "property", "member": SHIRAZ, "property": "Population"},
        {"kind": "property", "member": RIDES, "property": "Key"},
        {"kind": "property", "member": TEHRAN, "property": ""},
        {"kind": "property", "member": "[Geo].[City].&[99]", "property": "Key"},
    ]
    answers = formulas.evaluate_lookups("Sales", items, cube)
    assert values(answers) == ["1", "2", "not-found", "bad-request", "bad-request", "not-found"]
    assert 'Properties("Key")' in cube.queries[0] and "[Geo].[City].CurrentMember" in cube.queries[0]
    assert "no property called Population" in answers[2]["error"]["message"]


def test_property_names_cannot_break_out_of_their_string(cube):
    formulas.evaluate_lookups("Sales", [{"kind": "property", "member": TEHRAN, "property": 'Key") } SELECT'}], cube)
    assert 'Properties("Key"") } SELECT")' in cube.queries[0]


def test_set_counts_share_one_query_and_a_bad_set_fails_alone(cube):
    items = [
        {"kind": "count", "set": "{[Geo].[City].&[1], [Geo].[City].&[2]}"},
        {"kind": "count", "set": "[Geo].[City].Nope"},
        {"kind": "count", "set": "[Geo].[City].Children"},
        {"kind": "count", "set": "SELECT 1"},
        {"kind": "count", "set": "[Geo].[City].Children"},
    ]
    answers = formulas.evaluate_lookups("Sales", items, cube)
    assert values(answers) == [2, "not-found", 2, "bad-request", 2]
    assert cube.queries[0].startswith("WITH MEMBER [Measures].[__CubeExplorer Count 0] AS Count({{[Geo].[City].&[1]")
    assert "Count({[Geo].[City].Children})" in cube.queries[0]


def test_ranked_members_read_each_set_once(cube):
    items = [
        {"kind": "ranked", "set": "[Geo].[City].Children", "rank": 2},
        {"kind": "ranked", "set": "[Geo].[City].Children", "rank": 1},
        {"kind": "ranked", "set": "[Geo].[City].Children", "rank": 3},
        {"kind": "ranked", "set": "[Geo].[City].Children", "rank": 0},
        {"kind": "ranked", "set": "{[Date].[Year].&[2026]}", "rank": 1},
    ]
    answers = formulas.evaluate_lookups("Sales", items, cube)
    assert values(answers) == ["Shiraz", "Tehran", "not-found", "bad-request", "2026"]
    assert "fewer than 3" in answers[2]["error"]["message"]
    heads = [q for q in cube.queries if "Head(" in q]
    assert heads[0] == "SELECT {} ON 0, Head({[Geo].[City].Children}, 3) ON 1 FROM [Sales]"
    assert len(heads) == 2


# -- request shapes ----------------------------------------------------------------
@pytest.mark.parametrize("raw", ["x", [["a"], "b"], [[1]], [None]])
def test_read_tuples_refuses_wrong_shapes(raw):
    with pytest.raises(CubeError) as info:
        formulas.read_tuples(raw)
    assert info.value.category == "bad-request"


def test_read_tuples_limit():
    assert len(formulas.read_tuples([[]] * formulas.MAX_ITEMS)) == formulas.MAX_ITEMS
    with pytest.raises(CubeError):
        formulas.read_tuples([[]] * (formulas.MAX_ITEMS + 1))


@pytest.mark.parametrize(
    "raw",
    [
        "x",
        [{"kind": "nope"}],
        [{"kind": "member", "members": "x"}],
        [{"kind": "property", "member": TEHRAN}],
        [{"kind": "count"}],
        [{"kind": "ranked", "set": "x", "rank": "1"}],
        [{"kind": "ranked", "set": "x", "rank": True}],
        [{"kind": "count", "set": "x"}] * (formulas.MAX_ITEMS + 1),
    ],
)
def test_read_lookups_refuses_wrong_shapes(raw):
    with pytest.raises(CubeError) as info:
        formulas.read_lookups(raw)
    assert info.value.category == "bad-request"


def test_read_lookups_keeps_only_known_fields():
    assert formulas.read_lookups([{"kind": "ranked", "set": "s", "rank": 2, "extra": 1}]) == [{"kind": "ranked", "set": "s", "rank": 2}]
