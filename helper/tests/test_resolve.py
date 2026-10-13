import pytest
import xml_samples as x

from cubebridge.errors import CubeError
from cubebridge.ssas import resolve


@pytest.mark.parametrize(
    "text, expected",
    [
        ("bi-server", ("bi-server", None, None)),
        ("  BI-Server.corp.local  ", ("BI-Server.corp.local", None, None)),
        ("10.20.30.40", ("10.20.30.40", None, None)),
        ("bi-server:2384", ("bi-server", 2384, None)),
        ("bi-server\\TABULAR", ("bi-server", None, "TABULAR")),
        ("bi_server\\INST$01", ("bi_server", None, "INST$01")),
    ],
)
def test_parse_accepts_the_forms_windows_excel_accepts(text, expected):
    target = resolve.parse_server(text)
    assert (target.host, target.port, target.instance) == expected


@pytest.mark.parametrize(
    "text",
    ["", "   ", "bi server", "bi-server:0", "bi-server:70000", "bi-server:abc", "host\\", "\\inst", "a\\b\\c", "host;drop", "x" * 300],
)
def test_parse_rejects_junk(text):
    with pytest.raises(CubeError) as info:
        resolve.parse_server(text)
    assert info.value.category == "bad-request"


@pytest.mark.parametrize("text", ["https://bi.corp.local/olap/msmdpump.dll", "powerbi://api.powerbi.com/v1.0/myorg/Sales", "asazure://westeurope.asazure.windows.net/x"])
def test_web_addresses_get_a_specific_explanation(text):
    with pytest.raises(CubeError) as info:
        resolve.parse_server(text)
    assert info.value.category == "bad-request"
    assert "web address" in info.value.message.lower()


def test_default_instance_uses_2383_and_explicit_port_wins():
    def no_lookup(host, instance):
        raise AssertionError("no lookup expected")

    assert resolve.endpoint(resolve.parse_server("srv"), lookup=no_lookup) == ("srv", 2383)
    assert resolve.endpoint(resolve.parse_server("srv:2384"), lookup=no_lookup) == ("srv", 2384)


def test_named_instance_is_looked_up():
    calls = []

    def lookup(host, instance):
        calls.append((host, instance))
        return 50123

    assert resolve.endpoint(resolve.parse_server("srv\\TAB"), lookup=lookup) == ("srv", 50123)
    assert calls == [("srv", "TAB")]


def test_instance_port_is_read_from_the_redirector_reply():
    reply = x.rowset(
        [
            {"INSTANCE_NAME": "OTHER", "INSTANCE_PORT_NUMBER": "50001"},
            {"INSTANCE_NAME": "tab", "INSTANCE_PORT_NUMBER": "50123"},
        ],
        session_id=None,
    )
    assert resolve.port_from_reply(reply, "TAB") == 50123


def test_instance_missing_from_the_reply_is_an_instance_error():
    reply = x.rowset([{"INSTANCE_NAME": "OTHER", "INSTANCE_PORT_NUMBER": "50001"}], session_id=None)
    with pytest.raises(CubeError) as info:
        resolve.port_from_reply(reply, "TAB")
    assert info.value.category == "instance"
    assert "srv:port" in info.value.message or ":port" in info.value.message


def test_lookup_failure_becomes_an_instance_error_with_the_workaround():
    # Port 1 on localhost: nothing listens, so the redirector cannot be reached.
    with pytest.raises(CubeError) as info:
        resolve.instance_port("127.0.0.1", "TAB", redirector_port=1, timeout=1)
    assert info.value.category == "instance"
