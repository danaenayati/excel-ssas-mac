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
                "DEFAULT_FORMAT_STRING": "#,##0",
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
    "MDSCHEMA_LEVELS": x.rowset(
        [
            {"HIERARCHY_UNIQUE_NAME": "[Date].[Calendar]", "LEVEL_NUMBER": "0", "LEVEL_TYPE": "1", "LEVEL_DBTYPE": "3"},
            {"HIERARCHY_UNIQUE_NAME": "[Date].[Calendar]", "LEVEL_NUMBER": "2", "LEVEL_TYPE": "0", "LEVEL_DBTYPE": "3"},
            {"HIERARCHY_UNIQUE_NAME": "[Date].[Calendar]", "LEVEL_NUMBER": "1", "LEVEL_TYPE": "0", "LEVEL_DBTYPE": "7"},
            {"HIERARCHY_UNIQUE_NAME": "[Geo].[City]", "LEVEL_NUMBER": "0", "LEVEL_TYPE": "1", "LEVEL_DBTYPE": "130"},
            {"HIERARCHY_UNIQUE_NAME": "[Geo].[City]", "LEVEL_NUMBER": "1", "LEVEL_TYPE": "0", "LEVEL_DBTYPE": "130"},
            {"HIERARCHY_UNIQUE_NAME": "[Date].[Scenario]", "LEVEL_NUMBER": "0", "LEVEL_TYPE": "516", "LEVEL_DBTYPE": "3"},
        ]
    ),
    "MDSCHEMA_KPIS": x.rowset(
        [
            {
                "KPI_NAME": "Ride Growth", "KPI_CAPTION": "Ride Growth", "KPI_DISPLAY_FOLDER": "Targets",
                "KPI_VALUE": "[Measures].[Rides]", "KPI_GOAL": "[Measures].[Rides Goal]", "KPI_STATUS": "",
                "KPI_TREND": "[Measures].[Rides Trend]",
            },
            {"KPI_NAME": "Broken", "KPI_VALUE": "not a name"},
        ]
    ),
    "MDSCHEMA_SETS": x.rowset(
        [
            {"SET_NAME": "Top Cities", "SET_CAPTION": "Top Cities", "SET_DISPLAY_FOLDER": "", "DIMENSIONS": "[Geo].[City]"},
            {"SET_NAME": "Mixed", "SET_CAPTION": "", "SET_DISPLAY_FOLDER": "F", "DIMENSIONS": "[Geo].[City],[Date].[Calendar]"},
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
        {"uniqueName": "[Date].[Calendar]", "caption": "Calendar", "folder": "", "kind": "user",
         "allMember": "[Date].[Calendar].[All Periods]", "isDate": True},
        {"uniqueName": "[Date].[Scenario]", "caption": "Scenario", "folder": "", "kind": "attribute", "allMember": None,
         "isDate": True},
    ]
    assert geo["hierarchies"] == [
        {"uniqueName": "[Geo].[City]", "caption": "City", "folder": "Places", "kind": "attribute", "allMember": "[Geo].[City].[All]"},
    ]
    restricted = [r for r in backend.server.requests if "MDSCHEMA_HIERARCHIES" in r][0]
    assert "<CUBE_NAME>Sales</CUBE_NAME>" in restricted
    assert fields["kpis"] == [
        {"name": "Ride Growth", "caption": "Ride Growth", "folder": "Targets", "measure": "[Measures].[Rides]",
         "parts": ["value", "goal", "trend"]},
    ]
    assert fields["sets"] == [
        # A set spanning two fields lists both, so it can be placed.
        {"uniqueName": "[Mixed]", "caption": "Mixed", "folder": "F", "dimension": "[Geo]", "hierarchy": "[Geo].[City]",
         "allMember": "[Geo].[City].[All]", "hierarchies": ["[Geo].[City]", "[Date].[Calendar]"],
         "allMembers": ["[Geo].[City].[All]", "[Date].[Calendar].[All Periods]"]},
        {"uniqueName": "[Top Cities]", "caption": "Top Cities", "folder": "", "dimension": "[Geo]",
         "hierarchy": "[Geo].[City]", "allMember": "[Geo].[City].[All]", "hierarchies": ["[Geo].[City]"],
         "allMembers": ["[Geo].[City].[All]"]},
    ]


def test_fields_go_without_what_a_server_refuses_to_list(backend):
    for kind in ("MDSCHEMA_LEVELS", "MDSCHEMA_KPIS", "MDSCHEMA_SETS"):
        backend.server.responses.pop(kind, None)
    fields = backend.fields("Sales", "Sales")
    assert fields["kpis"] == [] and fields["sets"] == []
    assert not any(h.get("isDate") for d in fields["dimensions"] for h in d["hierarchies"])


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


# -- review fixes --------------------------------------------------------------------
import threading  # noqa: E402
import time  # noqa: E402

from cubebridge.ssas.session import Session  # noqa: E402

ROW_LAYOUT = Layout.from_json({"values": [{"measure": "[Measures].[Rides]", "caption": "Rides"}]})


def test_cancel_with_no_query_running_changes_nothing(backend):
    backend.catalogs()
    backend.cancel()
    assert len(backend.catalogs()) == 2
    assert backend.server.connections == 1


def test_cancel_stops_a_running_query_and_the_next_request_works(backend):
    backend.catalogs()
    backend.server.delay = 5.0
    threading.Timer(0.3, backend.cancel).start()
    started = time.monotonic()
    with pytest.raises(CubeError) as info:
        backend.query("Sales", "Sales", ROW_LAYOUT)
    assert info.value.category == "cancelled" and time.monotonic() - started < 3
    backend.server.delay = 0.0
    assert len(backend.catalogs()) == 2


def test_cancel_does_not_hit_a_metadata_call_that_happens_to_be_running(backend):
    backend.catalogs()
    backend.server.delay = 0.8
    threading.Timer(0.2, backend.cancel).start()
    assert len(backend.catalogs()) == 2  # no query is running, so cancel is a no-op
    assert backend.server.connections == 1


def test_a_query_that_times_out_is_not_silently_sent_again(ntlm_users):
    user, password = ntlm_users

    def quick_timeout(host, port, username, secret):
        return Session(host, port, username, secret, io_timeout=0.4, query_timeout=0.4)

    with FakeSsas(dict(RESPONSES)) as server:
        made = SsasBackend(f"127.0.0.1:{server.port}", user, password, session_factory=quick_timeout)
        made.catalogs()  # the session is now an established one, the case that used to be retried
        before = len(server.requests)
        server.delay = 3.0
        with pytest.raises(CubeError) as info:
            made.query("Sales", "Sales", ROW_LAYOUT)
        assert info.value.category == "unreachable"
        time.sleep(0.2)
        assert len(server.requests) == before + 1
        assert server.connections == 1
        made.close()


def test_closing_during_a_query_returns_promptly(backend):
    backend.catalogs()
    backend.server.delay = 5.0
    worker = threading.Thread(target=lambda: pytest.raises(CubeError, backend.query, "Sales", "Sales", ROW_LAYOUT))
    worker.start()
    time.sleep(0.3)
    started = time.monotonic()
    backend.close()
    assert time.monotonic() - started < 2
    worker.join(3)
    assert not worker.is_alive()


def test_a_named_instance_port_is_looked_up_again_after_a_failed_connect(ntlm_users):
    user, password = ntlm_users
    dead = __import__("socket").socket()
    dead.bind(("127.0.0.1", 0))
    dead_port = dead.getsockname()[1]
    dead.close()
    with FakeSsas(dict(RESPONSES)) as server:
        ports = iter([dead_port, server.port])
        calls = []

        def lookup(host, instance):
            calls.append(instance)
            return next(ports)

        made = SsasBackend("127.0.0.1\\TAB", user, password, lookup=lookup)
        with pytest.raises(CubeError):
            made.catalogs()  # the instance had moved; the stale port is not kept
        assert len(made.catalogs()) == 2
        assert calls == ["TAB", "TAB"]
        made.close()


def test_details_sends_a_drillthrough_and_names_columns_plainly(backend):
    backend.server.responses["DRILLTHROUGH"] = x.drillthrough_rowset(
        [("[$Trips].[City]", "xsd:string"), ("[$Trips].[Rides]", "xsd:long"), ("[$Drivers].[City]", "xsd:string")],
        [["Tehran", "5", "Tehran"], ["Tehran", "7", "Karaj"], ["Tehran", "1", "Tehran"]],
    )
    layout = Layout.from_json({"rows": [{"hierarchy": CITY, "caption": "City", "allMember": "[Geo].[City].[All]"}], "values": [{"measure": "[Measures].[Rides]", "caption": "Rides"}]})
    result = backend.details("Sales", "Sales", layout, ["[Geo].[City].&[1]"], "[Measures].[Rides]", 2)
    assert result == {
        "columns": ["Trips City", "Rides", "Drivers City"],
        "rows": [["Tehran", 5, "Tehran"], ["Tehran", 7, "Karaj"]],
        "truncated": True,
    }
    sent = backend.server.requests[-1]
    assert "DRILLTHROUGH MAXROWS 3 SELECT FROM [Sales] WHERE ([Measures].[Rides], [Geo].[City].&amp;[1])" in sent
    assert "<Format>Tabular</Format>" in sent


def test_query_passes_the_measures_format_to_running_totals(backend):
    layout = Layout.from_json(
        {
            "rows": [{"hierarchy": CITY, "caption": "City", "allMember": "[Geo].[City].[All]"}],
            "values": [{"measure": "[Measures].[Rides]", "caption": "Rides", "showAs": "runningTotalIn", "baseField": CITY}],
        }
    )
    backend.server.responses["WITH MEMBER"] = RESPONSES["SELECT NON EMPTY"]
    _, text = backend.query("Sales", "Sales", layout)
    assert 'FORMAT_STRING = "#,##0"' in text  # learnt from MDSCHEMA_MEASURES, fetched on demand
