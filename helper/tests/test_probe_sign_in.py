"""The connection test with a sign-in kind: web addresses, Kerberos, and the question."""

from fake_http_xmla import FakeXmlaHttp
from fake_kerberos import FakeKerberosAcceptor, FakeKerberosClient
from fake_ssas import FakeSsas
from test_probe import responses

from cubebridge import probe
from cubebridge.errors import CubeError
from cubebridge.ssas import kerberos


def test_a_web_address_is_tested_step_by_step(ntlm_users):
    user, password = ntlm_users
    lines = []
    with FakeXmlaHttp(responses(), offers=("Negotiate", "NTLM")) as server:
        assert probe.run(server.url, user, password, lines.append) is True
    text = "\n".join(lines)
    for expected in ("Web address", "OK (http, port", "User name and password (NTLM)", "Name lookup",
                     "Sign-in", "OK (NTLM (NTLM))", "16.0.43.211", "Archive, Sales", "Test query"):
        assert expected in text, expected
    assert password not in text


def test_kerberos_steps_are_reported(monkeypatch):
    monkeypatch.setattr(kerberos, "check_installed", lambda: None)
    monkeypatch.setattr(kerberos, "existing_ticket", lambda: None)

    def fake_client_context(username, password, host, **kwargs):
        context = FakeKerberosClient(f"MSOLAPSvc.3/{host}@CORP.LOCAL")
        context.cube_choice = kerberos.Choice("password", "CORP.LOCAL", f"MSOLAPSvc.3/{host}@CORP.LOCAL")
        return context

    monkeypatch.setattr(kerberos, "client_context", fake_client_context)
    lines = []
    with FakeSsas(responses(), acceptor=FakeKerberosAcceptor) as server:
        assert probe.run(f"127.0.0.1:{server.port}", "CORP\\dana", "pw-secret", lines.append, sign_in="kerberos") is True
    text = "\n".join(lines)
    assert "Windows Authentication (Kerberos)" in text
    assert "Kerberos on this Mac" in text and "none yet; the password gets one" in text
    assert "OK (Kerberos, ticket got with the password (MSOLAPSvc.3/127.0.0.1@CORP.LOCAL))" in text
    assert "pw-secret" not in text


def test_kerberos_that_is_not_installed_stops_before_any_network_use(monkeypatch):
    def missing():
        raise CubeError("auth-mechanism", 'Windows Authentication (Kerberos) is not installed on this Mac. Choose "User name and password (NTLM)".')

    monkeypatch.setattr(kerberos, "check_installed", missing)
    lines = []
    assert probe.run("no-such-host.invalid", "CORP\\dana", "x", lines.append, sign_in="kerberos") is False
    text = "\n".join(lines)
    assert "NOT INSTALLED" in text and "NTLM" in text and "Name lookup" not in text


def test_basic_for_a_server_name_is_not_valid():
    lines = []
    assert probe.run("bi-server", "CORP\\dana", "x", lines.append, sign_in="basic") is False
    assert "NOT VALID" in "\n".join(lines) and "https://" in "\n".join(lines)


def test_the_sign_in_question_offers_basic_only_for_https():
    assert probe.sign_in_choices("bi-server\\TAB") == ["ntlm", "kerberos"]
    assert probe.sign_in_choices("http://bi/olap/msmdpump.dll") == ["ntlm", "kerberos"]
    assert probe.sign_in_choices("https://bi/olap/msmdpump.dll") == ["ntlm", "kerberos", "basic"]


def test_the_sign_in_question_defaults_to_ntlm_and_insists_on_a_valid_number():
    shown = []
    answers = iter(["", ])
    assert probe.ask_sign_in("bi", lambda _prompt: next(answers), shown.append) == "ntlm"
    answers = iter(["9", "x", "3"])
    assert probe.ask_sign_in("https://bi/olap/msmdpump.dll", lambda _prompt: next(answers), shown.append) == "basic"
    assert "  2  Windows Authentication (Kerberos)" in shown
    assert sum("Type the number" in line for line in shown) == 2
