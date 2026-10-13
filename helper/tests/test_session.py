import socket
import threading
import time

import pytest
import xml_samples as x
from fake_ssas import FakeSsas

from cubebridge.errors import CubeError
from cubebridge.ssas.session import Session

CATALOGS = x.rowset([{"CATALOG_NAME": "Sales", "DESCRIPTION": "demo"}, {"CATALOG_NAME": "Ops", "DESCRIPTION": ""}])
ONE_CELL = x.cellset(
    axes=[x.axis("Axis0", [[x.member("[Measures]", "[Measures].[Rides]", "Rides", 0)]])],
    cells=[x.cell(0, "42", "xsd:int", "#,##0")],
)


def open_session(server, user, password, **kwargs) -> Session:
    session = Session("127.0.0.1", server.port, user, password, connect_timeout=5, io_timeout=5, **kwargs)
    session.open()
    return session


def test_sign_in_then_discover(ntlm_users):
    user, password = ntlm_users
    with FakeSsas({"DBSCHEMA_CATALOGS": CATALOGS}) as server:
        session = open_session(server, user, password)
        rows = session.discover("DBSCHEMA_CATALOGS")
        assert [r["CATALOG_NAME"] for r in rows] == ["Sales", "Ops"]
        session.close()
    # NEGO is clear on the first record of the connection and set on every later one.
    assert server.handshake_options[0] & 0x01 == 0
    assert all(o & 0x01 for o in server.handshake_options[1:])
    assert all(o & 0x01 for o in server.sealed_options)
    # The client never asks for binary or compressed responses.
    assert all(o & 0x1E == 0 for o in server.handshake_options + server.sealed_options)


def test_session_id_is_begun_once_and_reused(ntlm_users):
    user, password = ntlm_users
    with FakeSsas({"DBSCHEMA_CATALOGS": CATALOGS}) as server:
        session = open_session(server, user, password)
        session.discover("DBSCHEMA_CATALOGS")
        session.discover("DBSCHEMA_CATALOGS")
        session.close()
    assert "<BeginSession" in server.requests[0]
    assert 'SessionId="S-1"' in server.requests[1] and "<BeginSession" not in server.requests[1]


def test_execute_returns_a_cellset(ntlm_users):
    user, password = ntlm_users
    with FakeSsas({"SELECT": ONE_CELL}) as server:
        session = open_session(server, user, password)
        cellset = session.execute("SELECT {[Measures].[Rides]} ON 0 FROM [Sales]", catalog="Sales")
        assert cellset.cells[0].value == 42
        assert cellset.axes[0][0][0].caption == "Rides"
        session.close()
    assert "<Catalog>Sales</Catalog>" in server.requests[0]
    assert "<Format>Multidimensional</Format>" in server.requests[0]


def test_wrong_password_is_an_auth_error(ntlm_users):
    user, _ = ntlm_users
    with FakeSsas() as server:
        with pytest.raises(CubeError) as info:
            open_session(server, user, "not the password")
    assert info.value.category == "auth"


def test_nothing_listening_is_unreachable():
    probe = socket.socket()
    probe.bind(("127.0.0.1", 0))
    port = probe.getsockname()[1]
    probe.close()
    with pytest.raises(CubeError) as info:
        Session("127.0.0.1", port, "CORP\\dana", "x", connect_timeout=2).open()
    assert info.value.category == "unreachable"


def test_unknown_host_is_unreachable_and_names_the_host():
    with pytest.raises(CubeError) as info:
        Session("no-such-host.invalid", 2383, "CORP\\dana", "x", connect_timeout=2).open()
    assert info.value.category == "unreachable"
    assert "no-such-host.invalid" in info.value.message


def test_three_megabyte_response_round_trips(ntlm_users):
    user, password = ntlm_users
    big = x.rowset([{"N": str(i), "TEXT": "row " * 30} for i in range(20000)])
    assert len(big) > 3_000_000
    with FakeSsas({"BIG": big}, record_size=50000, frame_size=8000) as server:
        session = open_session(server, user, password)
        rows = session.discover("BIG")
        assert len(rows) == 20000 and rows[-1]["N"] == "19999"
        session.close()


def test_server_fault_is_reported_and_the_session_stays_usable(ntlm_users):
    user, password = ntlm_users
    responses = {"DBSCHEMA_CATALOGS": CATALOGS, "SELECT": x.inband_error("Query (1, 9) Parser: The syntax for 'x' is incorrect.")}
    with FakeSsas(responses) as server:
        session = open_session(server, user, password)
        with pytest.raises(CubeError) as info:
            session.execute("SELECT x", catalog="Sales")
        assert info.value.category == "server" and "syntax" in info.value.detail
        assert session.is_open
        assert len(session.discover("DBSCHEMA_CATALOGS")) == 2
        session.close()


def test_abort_from_another_thread_cancels_a_running_query(ntlm_users):
    user, password = ntlm_users
    with FakeSsas({"SELECT": ONE_CELL}) as server:
        session = open_session(server, user, password)
        server.delay = 5.0
        threading.Timer(0.3, session.abort).start()
        started = time.monotonic()
        with pytest.raises(CubeError) as info:
            session.execute("SELECT slow", catalog="Sales")
        assert info.value.category == "cancelled"
        assert time.monotonic() - started < 3
        assert not session.is_open


def test_dropped_connection_closes_the_session(ntlm_users):
    user, password = ntlm_users
    with FakeSsas({"DBSCHEMA_CATALOGS": CATALOGS}) as server:
        session = open_session(server, user, password)
        server.drop_after = 0
        with pytest.raises(CubeError) as info:
            session.discover("DBSCHEMA_CATALOGS")
        assert info.value.category == "unreachable"
        assert not session.is_open


def test_requests_on_a_closed_session_fail_cleanly(ntlm_users):
    user, password = ntlm_users
    with FakeSsas({"DBSCHEMA_CATALOGS": CATALOGS}) as server:
        session = open_session(server, user, password)
        session.close()
        with pytest.raises(CubeError) as info:
            session.discover("DBSCHEMA_CATALOGS")
        assert info.value.category == "unreachable"
