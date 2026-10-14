"""The backend picks the transport from the server as typed, and signs in as the
connection says: over HTTP against fake_http_xmla, Kerberos against fake_ssas."""

import threading
import time

import pytest
from fake_http_xmla import FakeXmlaHttp
from fake_kerberos import FakeKerberosAcceptor, FakeKerberosClient
from fake_ssas import FakeSsas
from test_ssas_backend import RESPONSES

from cubebridge.backends.ssas_backend import SsasBackend
from cubebridge.errors import CubeError
from cubebridge.pivot.model import Layout
from cubebridge.ssas import kerberos
from cubebridge.ssas.http_session import HttpSession
from cubebridge.ssas.session import Session

ROW_LAYOUT = Layout.from_json({"values": [{"measure": "[Measures].[Rides]", "caption": "Rides"}]})


def test_a_web_address_is_reached_over_http(ntlm_users):
    user, password = ntlm_users
    with FakeXmlaHttp(dict(RESPONSES), offers=("Negotiate", "NTLM")) as server:
        backend = SsasBackend(server.url, user, password)
        assert [c["name"] for c in backend.catalogs()] == ["Archive", "Sales"]
        assert backend.cubes("Sales")[0]["name"] == "Finance View"
        assert isinstance(backend._session, HttpSession)
        assert backend.sign_in_note == "NTLM (NTLM)"
        backend.close()
    assert server.connections == 1


def test_a_wrong_password_over_http_is_an_auth_error(ntlm_users):
    user, _ = ntlm_users
    with FakeXmlaHttp(dict(RESPONSES), offers=("NTLM",)) as server:
        with pytest.raises(CubeError) as info:
            SsasBackend(server.url, user, "nope").catalogs()
    assert info.value.category == "auth"


def test_cancel_stops_a_query_over_http_and_the_next_request_works(ntlm_users):
    user, password = ntlm_users
    with FakeXmlaHttp(dict(RESPONSES), offers=("NTLM",)) as server:
        backend = SsasBackend(server.url, user, password)
        backend.catalogs()
        server.delay = 5.0
        threading.Timer(0.3, backend.cancel).start()
        started = time.monotonic()
        with pytest.raises(CubeError) as info:
            backend.query("Sales", "Sales", ROW_LAYOUT)
        assert info.value.category == "cancelled" and time.monotonic() - started < 3
        server.delay = 0.0
        assert len(backend.catalogs()) == 2
        backend.close()


@pytest.mark.parametrize("server", ["bi-server", "bi-server\\TAB", "http://bi/olap/msmdpump.dll"])
def test_basic_is_refused_unless_the_address_is_https(server):
    with pytest.raises(CubeError) as info:
        SsasBackend(server, "CORP\\dana", "x", sign_in="basic")
    assert info.value.category == "bad-request" and "https://" in info.value.message


def test_an_unknown_sign_in_kind_is_refused():
    with pytest.raises(CubeError) as info:
        SsasBackend("bi-server", "CORP\\dana", "x", sign_in="digest")
    assert info.value.category == "bad-request"


def test_basic_over_https_is_accepted_without_any_network_use():
    backend = SsasBackend("https://bi.corp.local/olap/msmdpump.dll", "CORP\\dana", "x", sign_in="basic")
    backend.close()


def test_kerberos_over_tcp_reaches_the_named_instance_with_its_spn(monkeypatch):
    seen = {}

    def fake_client_context(username, password, host, **kwargs):
        seen.update(kwargs, host=host)
        return FakeKerberosClient(f"{kwargs['service']}/{host}:{kwargs['instance']}@CORP.LOCAL")

    monkeypatch.setattr(kerberos, "client_context", fake_client_context)
    with FakeSsas(dict(RESPONSES), acceptor=FakeKerberosAcceptor) as server:
        backend = SsasBackend("127.0.0.1\\TAB", "CORP\\dana", "pw", sign_in="kerberos", lookup=lambda host, instance: server.port)
        assert len(backend.catalogs()) == 2
        assert isinstance(backend._session, Session) and backend._session.sign_in == "kerberos"
        backend.close()
    assert seen == {"service": "MSOLAPSvc.3", "instance": "TAB", "host": "127.0.0.1"}
    assert server.contexts[0].target == "MSOLAPSvc.3/127.0.0.1:TAB@CORP.LOCAL"


def test_kerberos_problems_reach_the_caller_as_they_are(monkeypatch):
    def unavailable(*_args, **_kwargs):
        raise CubeError("auth-mechanism", "Windows Authentication (Kerberos) is not installed on this Mac.")

    monkeypatch.setattr(kerberos, "client_context", unavailable)
    with FakeSsas(dict(RESPONSES), acceptor=FakeKerberosAcceptor) as server:
        with pytest.raises(CubeError) as info:
            SsasBackend(f"127.0.0.1:{server.port}", "CORP\\dana", "pw", sign_in="kerberos").catalogs()
    assert info.value.category == "auth-mechanism" and "not installed" in info.value.message


def test_the_server_version_comes_over_the_open_session(ntlm_users):
    import xml_samples as x

    user, password = ntlm_users
    responses = dict(RESPONSES, DISCOVER_PROPERTIES=x.rowset([{"PropertyName": "DBMSVersion", "Value": "16.0.43.211"}]))
    with FakeSsas(responses) as server:
        backend = SsasBackend(f"127.0.0.1:{server.port}", user, password)
        backend.catalogs()
        assert backend.server_version() == "16.0.43.211"
        assert server.connections == 1
        backend.close()
