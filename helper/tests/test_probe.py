import socket

import xml_samples as x
from fake_ssas import FakeSsas
from test_ssas_backend import RESPONSES

from cubebridge import probe

SECRET = "correct horse battery"


def responses():
    canned = dict(RESPONSES)
    canned["DISCOVER_PROPERTIES"] = x.rowset([{"PropertyName": "DBMSVersion", "Value": "16.0.43.211"}])
    canned["SELECT NON EMPTY {[Measures].DefaultMember}"] = x.cellset(
        axes=[x.axis("Axis0", [[x.member("[Measures]", "[Measures].[Rides]", "Rides", 0)]])],
        cells=[x.cell(0, "42", "xsd:int")],
    )
    return canned


def test_every_step_passes_against_a_working_server(ntlm_users):
    user, password = ntlm_users
    lines = []
    with FakeSsas(responses()) as server:
        assert probe.run(f"127.0.0.1:{server.port}", user, password, lines.append) is True
    text = "\n".join(lines)
    for expected in ("Name lookup", "OK", "Sign-in", "16.0.43.211", "Archive, Sales", "Sales Cube", "Test query"):
        assert expected in text
    assert password not in text


def test_wrong_password_stops_at_sign_in_and_never_prints_it(ntlm_users):
    user, _ = ntlm_users
    lines = []
    with FakeSsas(responses()) as server:
        assert probe.run(f"127.0.0.1:{server.port}", user, "wrong-" + SECRET, lines.append) is False
    text = "\n".join(lines)
    assert "Sign-in" in text and "FAILED" in text and "username" in text.lower()
    assert SECRET not in text


def test_closed_port_is_explained():
    listener = socket.socket()
    listener.bind(("127.0.0.1", 0))
    port = listener.getsockname()[1]
    listener.close()
    lines = []
    assert probe.run(f"127.0.0.1:{port}", "CORP\\dana", "x", lines.append) is False
    text = "\n".join(lines)
    assert f"Port {port}" in text and "CLOSED" in text


def test_unknown_host_and_bad_server_text():
    lines = []
    assert probe.run("no-such-host.invalid", "CORP\\dana", "x", lines.append) is False
    assert "could not be found" in "\n".join(lines)
    lines.clear()
    assert probe.run("https://bi/olap", "CORP\\dana", "x", lines.append) is False
    assert "NOT VALID" in "\n".join(lines)


# -- review fixes --------------------------------------------------------------------
def test_verbose_output_never_contains_sign_in_tokens(ntlm_users):
    import io
    import logging

    user, password = ntlm_users
    stream = io.StringIO()
    undo = probe.enable_verbose(stream)
    try:
        with FakeSsas(responses()) as server:
            assert probe.run(f"127.0.0.1:{server.port}", user, password, lambda _line: None) is True
    finally:
        undo()
    text = stream.getvalue()
    assert "sign-in complete" in text  # our own step logging is on
    assert "TlRMTVNTUA" not in text  # base64 of "NTLMSSP": the library's token dump is not
    assert password not in text
    assert logging.getLogger().level != logging.DEBUG
    assert not logging.getLogger("spnego").isEnabledFor(logging.DEBUG)
