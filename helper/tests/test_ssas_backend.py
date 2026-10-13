import pytest
import xml_samples as x
from fake_ssas import FakeSsas

from cubebridge.backends.ssas_backend import SsasBackend
from cubebridge.errors import CubeError
from cubebridge.pivot.model import Layout

CITY = "[Geo].[City]"

RESPONSES = {
    "DBSCHEMA_CATALOGS": x.rowset(
        [{"CATALOG_NAME": "Sales", "DESCRIPTION": "Main"}, {"CATALOG_NAME": "Archive", "DESCRIPTION": ""}]
    ),
    "MDSCHEMA_CUBES": x.rowset(
        [
            {"CUBE_NAME": "Sales", "CUBE_CAPTION": "Sales Cube", "CUBE_SOURCE": "1"},
            {"CUBE_NAME": "$Geo", "CUBE_CAPTION": "Geo", "CUBE_SOURCE": "2"},
            {"CUBE_NAME": "Finance View", "CUBE_CAPTION": "", "CUBE_SOURCE": "1"},
        ]
    ),
    "MDSCHEMA_MEASURES": x.rowset(
        [
            {
                "MEASURE_NAME": "Rides",
                "MEASURE_UNIQUE_NAME": "[Measures].[Rides]",
                "MEASURE_CAPTION": "Rides",
                "MEASUREGROUP_NAME": "Trips",
                "MEASURE_DISPLAY_FOLDER": "",
                "MEASURE_IS_VISIBLE": "true",
            },
            {
                "MEASURE_NAME": "Internal",
                "MEASURE_UNIQUE_NAME": "[Measures].[Internal]",
                "MEASURE_CAPTION": "Internal",
                "MEASUREGROUP_NAME": "Trips",
                "MEASURE_DISPLAY_FOLDER": "",
                "MEASURE_IS_VISIBLE": "false",
            },
            {
                "MEASURE_NAME": "Avg Fare",
                "MEASURE_UNIQUE_NAME": "[Measures].[Avg Fare]",
                "MEASURE_CAPTION": "Average Fare",
                "MEASUREGROUP_NAME": "",
                "MEASURE_DISPLAY_FOLDER": "Ratios\\Money",
                "MEASURE_IS_VISIBLE": "true",
            },
        ]
    ),
    "MDSCHEMA_DIMENSIONS": x.rowset(
        [
            {"DIMENSION_UNIQUE_NAME": "[Measures]", "DIMENSION_CAPTION": "Measures", "DIMENSION_TYPE": "2", "DIMENSION_IS_VISIBLE": "true"},
            {"DIMENSION_UNIQUE_NAME": "[Geo]", "DIMENSION_CAPTION": "Geography", "DIMENSION_TYPE": "3", "DIMENSION_IS_VISIBLE": "true"},
            {"DIMENSION_UNIQUE_NAME": "[Date]", "DIMENSION_CAPTION": "Date", "DIMENSION_TYPE": "1", "DIMENSION_IS_VISIBLE": "true"},
            {"DIMENSION_UNIQUE_NAME": "[Audit]", "DIMENSION_CAPTION": "Audit", "DIMENSION_TYPE": "3", "DIMENSION_IS_VISIBLE": "false"},
        ]
    ),
    "MDSCHEMA_HIERARCHIES": x.rowset(
        [
            {
                "DIMENSION_UNIQUE_NAME": "[Measures]",
                "HIERARCHY_UNIQUE_NAME": "[Measures]",
                "HIERARCHY_CAPTION": "Measures",
                "HIERARCHY_ORIGIN": "0",
                "ALL_MEMBER": "",
                "HIERARCHY_IS_VISIBLE": "true",
            },
            {
                "DIMENSION_UNIQUE_NAME": "[Geo]",
                "HIERARCHY_UNIQUE_NAME": "[Geo].[City]",
                "HIERARCHY_CAPTION": "City",
                "HIERARCHY_DISPLAY_FOLDER": "Places",
                "HIERARCHY_ORIGIN": "2",
                "ALL_MEMBER": "[Geo].[City].[All]",
                "HIERARCHY_IS_VISIBLE": "true",
            },
            {
                "DIMENSION_UNIQUE_NAME": "[Geo]",
                "HIERARCHY_UNIQUE_NAME": "[Geo].[City Key]",
                "HIERARCHY_CAPTION": "City Key",
                "HIERARCHY_ORIGIN": "6",
                "ALL_MEMBER": "[Geo].[City Key].[All]",
                "HIERARCHY_IS_VISIBLE": "false",
            },
            {
                "DIMENSION_UNIQUE_NAME": "[Date]",
                "HIERARCHY_UNIQUE_NAME": "[Date].[Calendar]",
                "HIERARCHY_CAPTION": "Calendar",
                "HIERARCHY_ORIGIN": "1",
                "ALL_MEMBER": "[Date].[Calendar].[All Periods]",
                "HIERARCHY_IS_VISIBLE": "true",
            },
            {
                "DIMENSION_UNIQUE_NAME": "[Date]",
                "HIERARCHY_UNIQUE_NAME": "[Date].[Scenario]",
                "HIERARCHY_CAPTION": "Scenario",
                "HIERARCHY_ORIGIN": "2",
                "ALL_MEMBER": "",
                "HIERARCHY_IS_VISIBLE": "true",
            },
            {
                "DIMENSION_UNIQUE_NAME": "[Audit]",
                "HIERARCHY_UNIQUE_NAME": "[Audit].[Batch]",
                "HIERARCHY_CAPTION": "Batch",
                "HIERARCHY_ORIGIN": "2",
                "ALL_MEMBER": "[Audit].[Batch].[All]",
                "HIERARCHY_IS_VISIBLE": "true",
            },
        ]
    ),
    "SELECT {} ON COLUMNS": x.cellset(
        axes=[
            x.axis("Axis0", []),
            x.axis(
                "Axis1",
                [
                    [x.member(CITY, "[Geo].[City].[All]", "All", 0, 3)],
                    [x.member(CITY, "[Geo].[City].&[1]", "Tehran", 1, 0)],
                    [x.member(CITY, "[Geo].[City].&[2]", "Shiraz", 1, 0)],
                    [x.member(CITY, "[Geo].[City].&[3]", "Tabriz", 1, 0)],
                ],
            ),
        ],
        cells=[],
    ),
    "SELECT NON EMPTY": x.cellset(
        axes=[
            x.axis("Axis0", [[x.member("[Measures]", "[Measures].[Rides]", "Rides", 0)]]),
            x.axis("Axis1", [[x.member(CITY, "[Geo].[City].&[1]", "Tehran", 1)], [x.member(CITY, "[Geo].[City].[All]", "All", 0, 3)]]),
        ],
        cells=[x.cell(0, "120", "xsd:int", "#,##0"), x.cell(1, "200", "xsd:int", "#,##0")],
    ),
}


@pytest.fixture
def backend(ntlm_users):
    user, password = ntlm_users
    with FakeSsas(dict(RESPONSES)) as server:
        made = SsasBackend(f"127.0.0.1:{server.port}", user, password)
        made.server = server
        yield made
        made.close()


def test_catalogs_are_sorted(backend):
    assert backend.catalogs() == [
        {"name": "Archive", "description": ""},
        {"name": "Sales", "description": "Main"},
    ]


def test_cubes_hide_dimension_cubes_and_fall_back_to_the_name(backend):
    assert backend.cubes("Sales") == [
        {"name": "Finance View", "caption": "Finance View"},
        {"name": "Sales", "caption": "Sales Cube"},
    ]
    request = backend.server.requests[-1]
    assert "<CATALOG_NAME>Sales</CATALOG_NAME>" in request and "<Catalog>Sales</Catalog>" in request


def test_fields_are_grouped_filtered_and_typed(backend):
    fields = backend.fields("Sales", "Sales")
    assert fields["measures"] == [
        {"uniqueName": "[Measures].[Avg Fare]", "caption": "Average Fare", "group": "", "folder": "Ratios\\Money"},
        {"uniqueName": "[Measures].[Rides]", "caption": "Rides", "group": "Trips", "folder": ""},
    ]
    assert [d["caption"] for d in fields["dimensions"]] == ["Date", "Geography"]
    date, geo = fields["dimensions"]
    assert date["hierarchies"] == [
        {"uniqueName": "[Date].[Calendar]", "caption": "Calendar", "folder": "", "kind": "user", "allMember": "[Date].[Calendar].[All Periods]"},
        {"uniqueName": "[Date].[Scenario]", "caption": "Scenario", "folder": "", "kind": "attribute", "allMember": None},
    ]
    assert geo["hierarchies"] == [
        {"uniqueName": "[Geo].[City]", "caption": "City", "folder": "Places", "kind": "attribute", "allMember": "[Geo].[City].[All]"},
    ]
    restricted = [r for r in backend.server.requests if "MDSCHEMA_HIERARCHIES" in r][0]
    assert "<CUBE_NAME>Sales</CUBE_NAME>" in restricted


def test_members_drop_the_all_member_and_report_truncation(backend):
    result = backend.members("Sales", "Sales", CITY, None, None, 2)
    assert result == {
        "members": [
            {"uniqueName": "[Geo].[City].&[1]", "caption": "Tehran", "level": 1, "hasChildren": False},
            {"uniqueName": "[Geo].[City].&[2]", "caption": "Shiraz", "level": 1, "hasChildren": False},
        ],
        "truncated": True,
    }
    # The backend learned from the cube's metadata that this hierarchy has an All member.
    assert "Head({[Geo].[City].[All].Children}, 3)" in backend.server.requests[-1]


def test_members_of_a_hierarchy_without_all_member_start_at_level_zero(backend):
    backend.members("Sales", "Sales", "[Date].[Scenario]", None, None, 10)
    assert "Head({[Date].[Scenario].Levels(0).Members}, 11)" in backend.server.requests[-1]


def test_members_of_an_unknown_hierarchy_are_not_found(backend):
    with pytest.raises(CubeError) as info:
        backend.members("Sales", "Sales", "[Nope].[Nope]", None, None, 10)
    assert info.value.category == "not-found"


def test_query_sends_generated_mdx_and_returns_the_cellset(backend):
    layout = Layout.from_json(
        {
            "rows": [{"hierarchy": CITY, "caption": "City", "allMember": "[Geo].[City].[All]", "expanded": [], "members": [], "memberCaptions": []}],
            "values": [{"measure": "[Measures].[Rides]", "caption": "Rides"}],
        }
    )
    cellset, text = backend.query("Sales", "Sales", layout)
    assert text.startswith("SELECT NON EMPTY {[Measures].[Rides]} ON COLUMNS")
    assert cellset.cells[1].value == 200 and len(cellset.axes[1]) == 2
    sent = backend.server.requests[-1]
    assert "SELECT NON EMPTY {[Measures].[Rides]} ON COLUMNS" in sent and "<Catalog>Sales</Catalog>" in sent


def test_a_connection_the_server_dropped_is_reopened_transparently(backend):
    assert len(backend.catalogs()) == 2
    assert backend.server.connections == 1
    backend.server.drop_after = 1  # every connection now dies after answering one request
    assert len(backend.catalogs()) == 2  # first try hits the dead connection, the retry succeeds
    assert backend.server.connections == 2


def test_a_server_that_keeps_dropping_is_reported_not_retried_forever(backend):
    backend.catalogs()
    backend.server.drop_after = 0  # no connection answers anything
    with pytest.raises(CubeError) as info:
        backend.catalogs()
    assert info.value.category == "unreachable"
    assert backend.server.connections == 2


def test_wrong_password_surfaces_as_auth(ntlm_users):
    user, _ = ntlm_users
    with FakeSsas(dict(RESPONSES)) as server:
        with pytest.raises(CubeError) as info:
            SsasBackend(f"127.0.0.1:{server.port}", user, "nope").catalogs()
    assert info.value.category == "auth"


def test_bad_server_text_is_rejected_before_any_network_use():
    with pytest.raises(CubeError) as info:
        SsasBackend("https://bi/olap/msmdpump.dll", "CORP\\dana", "x")
    assert info.value.category == "bad-request"
