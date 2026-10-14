"""Cube formulas against the built-in sample cube."""

import pytest

from cubebridge.backends.demo_backend import CATALOG, CUBE, DemoBackend
from cubebridge.errors import CubeError
from cubebridge.pivot.model import Layout

TEHRAN = "[Geography].[City].[City].&[Tehran]"
SHIRAZ = "[Geography].[City].[City].&[Shiraz]"
IRAN = "[Geography].[Geography].[Country].&[Iran]"
Y2025 = "[Date].[Calendar].[Year].&[2025]"
JAN_2025 = "[Date].[Calendar].[Month].&[2025]&[Q1]&[1]"
RIDES = "[Measures].[Rides]"
REVENUE = "[Measures].[Revenue]"


@pytest.fixture(scope="module")
def demo():
    return DemoBackend()


def facts_sum(demo, measure, **match):
    return sum(f[measure] for f in demo._facts if all(f[k] == v for k, v in match.items()))


def test_values_match_the_fact_rows(demo):
    answers = demo.cells(
        CATALOG,
        CUBE,
        [[TEHRAN, Y2025, RIDES], [Y2025, TEHRAN], [IRAN, REVENUE], [], [TEHRAN.lower(), "[measures].[rides]"]],
    )
    tehran_2025 = facts_sum(demo, "rides", city="Tehran", year=2025)
    assert answers[0] == {"value": tehran_2025}
    assert answers[1] == {"value": tehran_2025}  # Rides is the default measure
    assert answers[2]["value"] == pytest.approx(facts_sum(demo, "revenue", country="Iran"))
    assert answers[3] == {"value": facts_sum(demo, "rides")}
    assert answers[4] == {"value": facts_sum(demo, "rides", city="Tehran")}


def test_values_agree_with_a_report(demo):
    layout = Layout.from_json(
        {
            "rows": [{"hierarchy": "[Geography].[City]", "caption": "City", "allMember": "[Geography].[City].[All]"}],
            "values": [{"measure": RIDES, "caption": "Rides"}],
            "filters": [{"hierarchy": "[Date].[Calendar]", "caption": "Calendar", "allMember": "[Date].[Calendar].[All]", "members": [Y2025]}],
        }
    )
    cellset, _ = demo.query(CATALOG, CUBE, layout)
    rows = cellset.axes[1]
    names = [[t[0].unique_name, Y2025, RIDES] for t in rows]
    answers = demo.cells(CATALOG, CUBE, names)
    assert [a["value"] for a in answers] == [cellset.cells[i].value for i in range(len(rows))]


def test_an_empty_combination_is_an_empty_value(demo):
    # Bike rides do not exist in Dubai.
    assert demo.cells(CATALOG, CUBE, [["[Geography].[City].[City].&[Dubai]", "[Service].[Service Type].[Service Type].&[Bike]"]]) == [{"value": None}]


def test_per_item_failures(demo):
    answers = demo.cells(CATALOG, CUBE, [["[Geography].[City].[City].&[Paris]"], ["Tehran"], [TEHRAN, SHIRAZ], ["[Measures].[Nope]"]])
    assert [a["error"]["category"] for a in answers] == ["not-found", "bad-request", "bad-request", "not-found"]


def test_wrong_database_or_cube_fails_the_request(demo):
    with pytest.raises(CubeError) as info:
        demo.cells("Nope", CUBE, [[TEHRAN]])
    assert info.value.category == "not-found"
    with pytest.raises(CubeError):
        demo.lookup(CATALOG, "Nope", [])


def test_member_captions_and_properties(demo):
    answers = demo.lookup(
        CATALOG,
        CUBE,
        [
            {"kind": "member", "members": [JAN_2025]},
            {"kind": "member", "members": [TEHRAN, Y2025]},
            {"kind": "member", "members": [REVENUE]},
            {"kind": "member", "members": ["[Geography].[City].[City].&[Paris]"]},
            {"kind": "property", "member": TEHRAN, "property": "Key"},
            {"kind": "property", "member": JAN_2025, "property": "LEVEL_NUMBER"},
            {"kind": "property", "member": TEHRAN, "property": "Population"},
        ],
    )
    assert [a.get("value", a.get("error", {}).get("category")) for a in answers] == [
        "Jan", "2025", "Revenue", "not-found", "Tehran", "3", "not-found",
    ]


def test_sets_count_and_rank(demo):
    by_rides = sorted({f["city"] for f in demo._facts}, key=lambda c: -facts_sum(demo, "rides", city=c))
    answers = demo.lookup(
        CATALOG,
        CUBE,
        [
            {"kind": "count", "set": "[Geography].[City].Members"},  # All and 9 cities
            {"kind": "count", "set": "[Geography].[Geography].[Country].Members"},
            {"kind": "count", "set": "{" + TEHRAN + ", " + SHIRAZ + "}"},
            {"kind": "count", "set": IRAN + ".Children"},
            {"kind": "ranked", "set": "TopCount([Geography].[City].[City].Members, 3, [Measures].[Rides])", "rank": 2},
            {"kind": "ranked", "set": "{" + SHIRAZ + ", " + TEHRAN + "}", "rank": 2},
            {"kind": "ranked", "set": "{" + SHIRAZ + "}", "rank": 2},
            {"kind": "count", "set": "Filter([Geography].[City].Members, 1 = 1)"},
            {"kind": "count", "set": "SELECT"},
        ],
    )
    got = [a.get("value", a.get("error", {}).get("category")) for a in answers]
    assert got == [10, 3, 2, 5, by_rides[1], "Tehran", "not-found", "bad-request", "bad-request"]
