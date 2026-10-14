"""Sign-in kinds through the API and the settings file: saved, validated, passed on."""

import json

import pytest
from test_api import PORT, call

from cubebridge.api import App
from cubebridge.errors import CubeError
from cubebridge.store import MemorySecrets, Store


class SignInBackend:
    """Records how it was made; takes the sign-in kind as the real backend does."""

    made: list = []

    def __init__(self, server, username, password, sign_in="ntlm"):
        self.args = (server, username, password, sign_in)
        SignInBackend.made.append(self)

    def catalogs(self):
        return [{"name": "Sales", "description": ""}]

    def cancel(self):
        pass

    def close(self):
        pass


@pytest.fixture
def app(tmp_path):
    SignInBackend.made = []
    store = Store(tmp_path / "app", secrets=MemorySecrets())
    return App(store, port=PORT, pane_dir=None, backend_factory=SignInBackend, cert_path=None)


def add(app, **fields):
    body = {"server": "bi-server", "username": "CORP\\dana", "password": "pw-1", **fields}
    return call(app, "POST", "/api/connections", body)


@pytest.mark.parametrize(
    "fields, sign_in",
    [
        ({}, "ntlm"),
        ({"signIn": "ntlm"}, "ntlm"),
        ({"signIn": "kerberos"}, "kerberos"),
        ({"signIn": "kerberos", "server": "https://bi.corp.local/olap/msmdpump.dll"}, "kerberos"),
        ({"signIn": "basic", "server": "https://bi.corp.local/olap/msmdpump.dll"}, "basic"),
    ],
)
def test_the_sign_in_kind_is_tested_saved_and_reported(app, fields, sign_in):
    status, created = add(app, **fields)
    assert status == 200 and created["signIn"] == sign_in
    assert SignInBackend.made[0].args[3] == sign_in
    saved = json.loads((app.store.directory / "config.json").read_text())["connections"][0]
    assert saved["signIn"] == sign_in and "pw-1" not in json.dumps(saved)
    listed = call(app, "GET", "/api/connections")[1]
    assert listed[1]["signIn"] == sign_in and "signIn" not in listed[0]  # the sample cube has none


@pytest.mark.parametrize(
    "fields, fragment",
    [
        ({"signIn": "basic"}, "https://"),
        ({"signIn": "basic", "server": "http://bi.corp.local/olap/msmdpump.dll"}, "https://"),
        ({"signIn": "digest"}, "signIn"),
        ({"signIn": 3}, "signIn"),
        ({"server": "ftp://bi/olap"}, "web address"),
    ],
)
def test_bad_sign_in_settings_are_refused_before_anything_is_sent(app, fields, fragment):
    status, body = add(app, **fields)
    assert status == 400 and body["error"]["category"] == "bad-request"
    assert fragment in body["error"]["message"]
    assert SignInBackend.made == [] and app.store.connections()[1:] == []


def test_a_saved_connection_reopens_with_its_sign_in_kind(app, tmp_path):
    _, created = add(app, signIn="kerberos")
    restarted = App(app.store, port=PORT, pane_dir=None, backend_factory=SignInBackend, cert_path=None)
    assert call(restarted, "GET", f"/api/connections/{created['id']}/catalogs")[0] == 200
    assert SignInBackend.made[-1].args == ("bi-server", "CORP\\dana", "pw-1", "kerberos")


def test_a_connection_saved_before_sign_in_kinds_existed_signs_in_with_ntlm(tmp_path):
    directory = tmp_path / "app"
    directory.mkdir()
    old = {"id": "abc123", "name": "BI", "kind": "ssas", "server": "bi-server", "username": "CORP\\dana"}
    (directory / "config.json").write_text(json.dumps({"token": "t" * 43, "connections": [old]}))
    secrets = MemorySecrets()
    secrets.set("abc123", "pw-old")
    store = Store(directory, secrets=secrets)
    assert store.get("abc123")["signIn"] == "ntlm"
    SignInBackend.made = []
    app = App(store, port=PORT, pane_dir=None, backend_factory=SignInBackend, cert_path=None)
    assert call(app, "GET", "/api/connections/abc123/catalogs")[0] == 200
    assert SignInBackend.made[0].args == ("bi-server", "CORP\\dana", "pw-old", "ntlm")
    # Reading it changed nothing in the file.
    assert "signIn" not in json.loads((directory / "config.json").read_text())["connections"][0]


def test_a_three_argument_backend_factory_still_works_for_ntlm(tmp_path):
    made = []

    def old_factory(server, username, password):
        made.append((server, username, password))
        return SignInBackend(server, username, password)

    app = App(Store(tmp_path / "app", secrets=MemorySecrets()), port=PORT, pane_dir=None, backend_factory=old_factory)
    assert add(app)[0] == 200 and made == [("bi-server", "CORP\\dana", "pw-1")]


def test_error_messages_name_the_choices_in_the_panes_own_words():
    from pathlib import Path

    from cubebridge.ssas import auth

    pane = (Path(__file__).resolve().parents[2] / "addin" / "src" / "state" / "signIn.ts").read_text(encoding="utf-8")
    for kind, label in auth.LABELS.items():
        assert f'{kind}: "{label}"' in pane


def test_the_real_backend_refuses_basic_for_a_server_name():
    from cubebridge.backends.ssas_backend import SsasBackend

    with pytest.raises(CubeError):
        SsasBackend("bi-server", "u", "p", sign_in="basic")
