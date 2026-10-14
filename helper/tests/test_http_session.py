"""XMLA over HTTP(S) against a stand-in msmdpump.dll (fake_http_xmla.py)."""

import base64
import hashlib
import socket
import ssl
import threading
import time

import pytest
import xml_samples as x
from fake_http_xmla import FakeXmlaHttp
from fake_kerberos import spnego_acceptor
from test_kerberos import FQDN, REALM, Kdc, kdc  # noqa: F401 - kdc is a fixture

from cubebridge import certs
from cubebridge.errors import CubeError
from cubebridge.ssas import http_session, kerberos, resolve
from cubebridge.ssas.http_session import HttpSession, parse_challenges

CATALOGS = x.rowset([{"CATALOG_NAME": "Sales", "DESCRIPTION": "demo"}, {"CATALOG_NAME": "Ops", "DESCRIPTION": ""}])
ONE_CELL = x.cellset(
    axes=[x.axis("Axis0", [[x.member("[Measures]", "[Measures].[Rides]", "Rides", 0)]])],
    cells=[x.cell(0, "42", "xsd:int", "#,##0")],
)
RESPONSES = {"DBSCHEMA_CATALOGS": CATALOGS, "SELECT": ONE_CELL}


def session_for(server, user="CORP\\dana", password="correct horse battery", **kwargs) -> HttpSession:
    kwargs.setdefault("connect_timeout", 5)
    kwargs.setdefault("io_timeout", 5)
    session = HttpSession(resolve.parse_target(server.url), user, password, **kwargs)
    session.open()
    return session


def catalogs(session) -> list[str]:
    return [r["CATALOG_NAME"] for r in session.discover("DBSCHEMA_CATALOGS")]


# -- the requests ----------------------------------------------------------------------
def test_soap_requests_with_sessions_and_plain_xml_answers():
    with FakeXmlaHttp(RESPONSES, offers=()) as server:
        session = session_for(server)
        assert catalogs(session) == ["Sales", "Ops"]
        assert session.execute("SELECT {[Measures].[Rides]} ON 0 FROM [Sales]", "Sales").cells[0].value == 42
        session.close()
    first, second = server.requests
    assert first["headers"]["content-type"] == "text/xml"
    assert first["headers"]["soapaction"] == '"urn:schemas-microsoft-com:xml-analysis:Discover"'
    assert second["headers"]["soapaction"] == '"urn:schemas-microsoft-com:xml-analysis:Execute"'
    assert first["body"].startswith(b'<?xml version="1.0" encoding="utf-8"?><Envelope')
    assert "<BeginSession" in first["text"]
    assert 'SessionId="S-1"' in second["text"] and "<BeginSession" not in second["text"]
    assert "authorization" not in first["headers"]  # nobody asked for a sign-in
    assert {r["connection"] for r in server.requests} == {1}


def test_a_soap_fault_is_a_server_error_and_the_session_stays_usable():
    responses = dict(RESPONSES, SELECT=x.soap_fault("Query (1, 8) The syntax for 'x' is incorrect."))
    with FakeXmlaHttp(responses, offers=()) as server:
        session = session_for(server)
        with pytest.raises(CubeError) as info:
            session.execute("SELECT x", "Sales")
        assert info.value.category == "server" and "syntax" in info.value.detail
        assert session.is_open and catalogs(session) == ["Sales", "Ops"]
        session.close()


def test_a_three_megabyte_answer_arrives_whole():
    big = x.rowset([{"N": str(i), "TEXT": "row " * 30} for i in range(20000)])
    with FakeXmlaHttp({"BIG": big}, offers=()) as server:
        session = session_for(server)
        rows = session.discover("BIG")
        assert len(rows) == 20000 and rows[-1]["N"] == "19999"
        session.close()


def test_an_answer_over_the_limit_is_refused(monkeypatch):
    monkeypatch.setattr(http_session, "MAX_RESPONSE", 100)
    with FakeXmlaHttp(RESPONSES, offers=()) as server:
        session = session_for(server)
        with pytest.raises(CubeError) as info:
            catalogs(session)
        assert info.value.category == "server" and "MB" in info.value.detail


# -- NTLM --------------------------------------------------------------------------
def test_ntlm_signs_the_connection_in_once(ntlm_users):
    user, password = ntlm_users
    with FakeXmlaHttp(RESPONSES, offers=("NTLM",)) as server:
        session = session_for(server, user, password)
        assert catalogs(session) == ["Sales", "Ops"]
        assert catalogs(session) == ["Sales", "Ops"]
        session.close()
    # anonymous 401, NTLM negotiate -> 401 + challenge, authenticate -> 200, then just the request
    assert [r.get("scheme") for r in server.requests] == [None, "NTLM", "NTLM", None]
    assert {r["connection"] for r in server.requests} == {1}
    assert session.sign_in_note == "NTLM (NTLM)"


def test_ntlm_goes_under_negotiate_when_that_is_all_the_server_offers(ntlm_users):
    user, password = ntlm_users
    with FakeXmlaHttp(RESPONSES, offers=("Negotiate",)) as server:
        session = session_for(server, user, password)
        assert catalogs(session) == ["Sales", "Ops"]
        session.close()
    assert [r.get("scheme") for r in server.requests] == [None, "Negotiate", "Negotiate"]
    token = base64.b64decode(server.requests[1]["headers"]["authorization"].split()[1])
    assert token.startswith(b"NTLMSSP\x00")


def test_a_wrong_ntlm_password_is_an_auth_error_and_never_shown(ntlm_users):
    user, _ = ntlm_users
    with FakeXmlaHttp(RESPONSES, offers=("NTLM",)) as server:
        session = session_for(server, user, "not-the-password-123")
        with pytest.raises(CubeError) as info:
            catalogs(session)
    assert info.value.category == "auth" and "refused" in info.value.message
    assert "not-the-password-123" not in info.value.message + (info.value.detail or "")
    assert not session.is_open


def test_a_server_that_wants_a_sign_in_on_every_request_gets_one(ntlm_users):
    user, password = ntlm_users
    with FakeXmlaHttp(RESPONSES, offers=("NTLM",)) as server:
        server.per_request_sign_in = True
        session = session_for(server, user, password)
        assert catalogs(session) == ["Sales", "Ops"]
        assert catalogs(session) == ["Sales", "Ops"]
        session.close()
    assert len(server.served) == 2 and len(server.contexts) == 2


def test_an_idle_connection_the_server_closed_is_opened_and_signed_in_again(ntlm_users):
    user, password = ntlm_users
    with FakeXmlaHttp(RESPONSES, offers=("NTLM",)) as server:
        server.close_after_response = True
        session = session_for(server, user, password)
        assert catalogs(session) == ["Sales", "Ops"]
        assert catalogs(session) == ["Sales", "Ops"]
        session.close()
    assert server.connections >= 2
    assert 'SessionId="S-1"' in server.served[1]["text"]  # the XMLA session outlives the connection


def test_a_server_that_closes_the_connection_mid_handshake_is_explained(ntlm_users):
    user, password = ntlm_users
    with FakeXmlaHttp(RESPONSES, offers=("NTLM",)) as server:
        server.close_after_401 = True
        session = session_for(server, user, password)
        with pytest.raises(CubeError) as info:
            catalogs(session)
    assert info.value.category == "auth" and "middle of the sign-in" in info.value.message


# -- Basic -----------------------------------------------------------------------------
@pytest.fixture
def tls(tmp_path):
    """A server certificate for localhost and 127.0.0.1, and a client that trusts it."""
    paths = certs.generate(tmp_path / "certs")
    server = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    server.load_cert_chain(str(paths.cert), str(paths.key))
    client = ssl.create_default_context(cafile=str(paths.ca))
    return server, client, paths


def test_basic_over_https_waits_to_be_asked_then_goes_with_every_request(tls):
    server_tls, client_tls, _ = tls
    with FakeXmlaHttp(RESPONSES, offers=("Basic",), basic_users={"CORP\\dana": "pw"}, tls_context=server_tls) as server:
        session = session_for(server, "CORP\\dana", "pw", sign_in="basic", ssl_context=client_tls)
        assert catalogs(session) == ["Sales", "Ops"]
        assert catalogs(session) == ["Sales", "Ops"]
        session.close()
    assert [r.get("scheme") for r in server.requests] == [None, "Basic", "Basic"]
    assert session.sign_in_note == "Basic, over https"


def test_a_wrong_basic_password_is_an_auth_error(tls):
    server_tls, client_tls, _ = tls
    with FakeXmlaHttp(RESPONSES, offers=("Basic",), basic_users={"CORP\\dana": "pw"}, tls_context=server_tls) as server:
        session = session_for(server, "CORP\\dana", "nope", sign_in="basic", ssl_context=client_tls)
        with pytest.raises(CubeError) as info:
            catalogs(session)
    assert info.value.category == "auth" and "user name or password" in info.value.message
    assert len(server.requests) == 2  # asked once, refused once: no retrying the password


def test_basic_is_refused_over_plain_http_before_anything_is_sent():
    with FakeXmlaHttp(RESPONSES, offers=("Basic",)) as server:
        with pytest.raises(CubeError) as info:
            HttpSession(resolve.parse_target(server.url), "u", "p", sign_in="basic")
    assert info.value.category == "bad-request" and "https://" in info.value.message
    assert server.requests == []


@pytest.mark.parametrize(
    "sign_in, offers, expected",
    [
        ("ntlm", ("Basic",), "User name and password over HTTPS (Basic)"),
        ("basic", ("Negotiate", "NTLM"), "User name and password (NTLM)"),
        ("kerberos", ("NTLM",), "User name and password (NTLM)"),
    ],
)
def test_a_sign_in_the_server_does_not_offer_names_the_choice_that_would_work(tls, sign_in, offers, expected):
    server_tls, client_tls, _ = tls
    with FakeXmlaHttp(RESPONSES, offers=offers, tls_context=server_tls) as server:
        session = session_for(server, sign_in=sign_in, ssl_context=client_tls)
        with pytest.raises(CubeError) as info:
            catalogs(session)
    assert info.value.category == "auth-mechanism" and expected in info.value.message
    # Nothing was sent another way: no request carried any credentials at all.
    assert all("authorization" not in r["headers"] for r in server.requests)


# -- Kerberos ----------------------------------------------------------------------
def test_kerberos_uses_negotiate_with_a_ticket_for_the_web_server(kdc):  # noqa: F811
    kdc.services.add(f"HTTP/{FQDN}@{REALM}")

    def factory(kind, username, password, _host, bindings):
        assert kind == "kerberos"
        return kerberos.client_context(username, password, "bi", service="HTTP", channel_bindings=bindings)

    with FakeXmlaHttp(RESPONSES, offers=("Negotiate", "NTLM"), acceptor=lambda scheme: spnego_acceptor("HTTP")) as server:
        session = session_for(server, sign_in="kerberos", context_factory=factory)
        assert catalogs(session) == ["Sales", "Ops"]
        assert catalogs(session) == ["Sales", "Ops"]
        session.close()
    assert [r.get("scheme") for r in server.requests] == [None, "Negotiate", None]
    assert server.contexts[0].fake.target == f"HTTP/{FQDN}@{REALM}"
    assert "Kerberos" in session.sign_in_note


def test_kerberos_channel_bindings_over_https_are_the_certificates_hash(tls, kdc):  # noqa: F811
    server_tls, client_tls, paths = tls
    kdc.services.add(f"HTTP/{FQDN}@{REALM}")
    seen = []

    def factory(kind, username, password, _host, bindings):
        seen.append(bindings)
        return kerberos.client_context(username, password, "bi", service="HTTP", channel_bindings=bindings)

    with FakeXmlaHttp(RESPONSES, acceptor=lambda scheme: spnego_acceptor("HTTP"), tls_context=server_tls) as server:
        session = session_for(server, sign_in="kerberos", context_factory=factory, ssl_context=client_tls)
        assert catalogs(session) == ["Sales", "Ops"]
        session.close()
    der = ssl.PEM_cert_to_DER_cert(paths.cert.read_text())
    assert seen[0].application_data == b"tls-server-end-point:" + hashlib.sha256(der).digest()


# -- NTLM over https with Extended Protection ----------------------------------------
def test_ntlm_over_https_carries_the_channel_binding_iis_checks(tls, ntlm_users):
    from spnego.channel_bindings import GssChannelBindings

    server_tls, client_tls, paths = tls
    user, password = ntlm_users
    der = ssl.PEM_cert_to_DER_cert(paths.cert.read_text())
    expected = GssChannelBindings(application_data=b"tls-server-end-point:" + hashlib.sha256(der).digest())
    wrong = GssChannelBindings(application_data=b"tls-server-end-point:" + b"\x00" * 32)

    def acceptor(bindings):
        return lambda scheme: __import__("spnego").server(protocol="ntlm", channel_bindings=bindings)

    with FakeXmlaHttp(RESPONSES, offers=("NTLM",), acceptor=acceptor(expected), tls_context=server_tls) as server:
        session = session_for(server, user, password, ssl_context=client_tls)
        assert catalogs(session) == ["Sales", "Ops"]
        session.close()
    with FakeXmlaHttp(RESPONSES, offers=("NTLM",), acceptor=acceptor(wrong), tls_context=server_tls) as server:
        session = session_for(server, user, password, ssl_context=client_tls)
        with pytest.raises(CubeError) as info:
            catalogs(session)
        assert info.value.category == "auth" and server.refusals == ["BadBindingsError"]


# -- TLS -------------------------------------------------------------------------------
def test_an_untrusted_certificate_stops_before_anything_is_sent(tls):
    server_tls, _client_tls, _ = tls
    untrusting = ssl.create_default_context()  # Python's public authorities only
    with FakeXmlaHttp(RESPONSES, offers=(), tls_context=server_tls) as server:
        with pytest.raises(CubeError) as info:
            session_for(server, ssl_context=untrusting)
    assert info.value.category == "unreachable" and "certificate" in info.value.message
    assert server.requests == []


def test_the_default_tls_context_verifies_and_uses_the_system_trust_store():
    context = http_session.default_ssl_context()
    assert context.verify_mode == ssl.CERT_REQUIRED and context.check_hostname
    assert context.minimum_version >= ssl.TLSVersion.TLSv1_2
    assert type(context).__module__.startswith("truststore")


# -- cancel, errors, proxies ---------------------------------------------------------
def test_abort_from_another_thread_cancels_a_running_query():
    with FakeXmlaHttp(RESPONSES, offers=()) as server:
        session = session_for(server)
        server.delay = 5.0
        threading.Timer(0.3, session.abort).start()
        started = time.monotonic()
        with pytest.raises(CubeError) as info:
            session.execute("SELECT slow", "Sales")
        assert info.value.category == "cancelled" and time.monotonic() - started < 3
        assert not session.is_open


def test_abort_also_cancels_over_https(tls):
    server_tls, client_tls, _ = tls
    with FakeXmlaHttp(RESPONSES, offers=(), tls_context=server_tls) as server:
        session = session_for(server, ssl_context=client_tls)
        server.delay = 5.0
        threading.Timer(0.3, session.abort).start()
        started = time.monotonic()
        with pytest.raises(CubeError) as info:
            session.execute("SELECT slow", "Sales")
        assert info.value.category == "cancelled" and time.monotonic() - started < 3


def test_a_read_timeout_is_not_marked_for_retrying():
    with FakeXmlaHttp(RESPONSES, offers=()) as server:
        session = session_for(server, io_timeout=0.4)
        server.delay = 2.0
        with pytest.raises(CubeError) as info:
            catalogs(session)
    assert info.value.category == "unreachable" and info.value.retry is False


@pytest.mark.parametrize(
    "status, headers, category, fragment",
    [
        (404, {}, "unreachable", "msmdpump.dll"),
        (302, {"Location": "https://bi.corp.local/olap2/msmdpump.dll?x=secret"}, "unreachable", "https://bi.corp.local/olap2/msmdpump.dll)"),
        (403, {}, "auth", "403"),
        (503, {}, "unreachable", "503"),
        (418, {}, "server", "418"),
    ],
)
def test_http_errors_are_plain_sentences(status, headers, category, fragment):
    with FakeXmlaHttp(RESPONSES, offers=()) as server:
        server.status_override = (status, {"Content-Type": "text/html", **headers}, b"<html>nope</html>")
        session = session_for(server)
        with pytest.raises(CubeError) as info:
            catalogs(session)
    assert info.value.category == category and fragment in info.value.message
    assert "secret" not in info.value.message


def test_an_answer_that_is_not_xml_is_a_protocol_error():
    with FakeXmlaHttp(RESPONSES, offers=()) as server:
        server.status_override = (200, {"Content-Type": "text/html"}, b"hello")
        session = session_for(server)
        with pytest.raises(CubeError) as info:
            catalogs(session)
    assert info.value.category == "server" and "not XML" in info.value.detail


def test_nothing_listening_and_unknown_names_are_unreachable():
    probe = socket.socket()
    probe.bind(("127.0.0.1", 0))
    port = probe.getsockname()[1]
    probe.close()
    with pytest.raises(CubeError) as info:
        HttpSession(resolve.parse_target(f"http://127.0.0.1:{port}/olap/msmdpump.dll"), "u", "p", connect_timeout=2).open()
    assert info.value.category == "unreachable"
    with pytest.raises(CubeError) as info:
        HttpSession(resolve.parse_target("https://no-such-host.invalid/olap/msmdpump.dll"), "u", "p", connect_timeout=2).open()
    assert info.value.category == "unreachable" and "no-such-host.invalid" in info.value.message


def test_no_proxy_is_used_whatever_the_environment_says(monkeypatch, ntlm_users):
    user, password = ntlm_users
    proxy = socket.socket()
    proxy.bind(("127.0.0.1", 0))
    proxy.listen(8)
    proxy.settimeout(0.2)
    address = f"http://127.0.0.1:{proxy.getsockname()[1]}"
    for name in ("HTTP_PROXY", "HTTPS_PROXY", "ALL_PROXY", "http_proxy", "https_proxy", "all_proxy"):
        monkeypatch.setenv(name, address)
    monkeypatch.delenv("NO_PROXY", raising=False)
    monkeypatch.delenv("no_proxy", raising=False)
    try:
        with FakeXmlaHttp(RESPONSES, offers=("NTLM",)) as server:
            session = session_for(server, user, password)
            assert catalogs(session) == ["Sales", "Ops"]
            session.close()
        with pytest.raises(TimeoutError):
            proxy.accept()  # nothing ever connected to the proxy
    finally:
        proxy.close()


# -- WWW-Authenticate ------------------------------------------------------------------
@pytest.mark.parametrize(
    "values, expected",
    [
        (["Negotiate", "NTLM"], {"negotiate": "", "ntlm": ""}),
        (["Negotiate, NTLM"], {"negotiate": "", "ntlm": ""}),
        (['Basic realm="a, b", charset="UTF-8"'], {"basic": 'realm="a, b"'}),
        (["Negotiate oYG3MIG0oAMKAQ=="], {"negotiate": "oYG3MIG0oAMKAQ=="}),
        (["NTLM TlRMTVNTUAACAAAA", 'Basic realm="x"'], {"ntlm": "TlRMTVNTUAACAAAA", "basic": 'realm="x"'}),
        ([], {}),
    ],
)
def test_www_authenticate_offers_are_read(values, expected):
    assert parse_challenges(values) == expected
