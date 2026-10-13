import json
import ssl
import threading
import urllib.error
import urllib.request

import pytest

from cubebridge import certs
from cubebridge.api import App, serve
from cubebridge.backends.demo_backend import CATALOG, CUBE
from cubebridge.errors import CubeError
from cubebridge.store import MemorySecrets, Store

PORT = 47613
HOST = f"localhost:{PORT}"
ORIGIN = f"https://{HOST}"


class FakeBackend:
    """Stands in for a real server connection."""

    instances: list = []

    def __init__(self, server, username, password):
        self.args = (server, username, password)
        self.closed = False
        self.cancelled = 0
        FakeBackend.instances.append(self)
        if password == "wrong":
            raise CubeError("auth", "The server refused the sign-in.")

    def catalogs(self):
        return [{"name": "Sales", "description": ""}]

    def cancel(self):
        self.cancelled += 1

    def close(self):
        self.closed = True


@pytest.fixture
def app(tmp_path):
    FakeBackend.instances = []
    pane = tmp_path / "pane"
    (pane / "assets").mkdir(parents=True)
    (pane / "index.html").write_text('<meta name="cube-token" content="__CUBE_TOKEN__"><div id="root"></div>')
    (pane / "assets" / "app.js").write_text("console.log('pane')")
    (tmp_path / "secret.txt").write_text("outside the pane folder")
    store = Store(tmp_path / "app", secrets=MemorySecrets())
    return App(store, port=PORT, pane_dir=pane, backend_factory=FakeBackend, cert_path=None)


def call(app, method, path, body=None, *, token="valid", host=HOST, origin=None):
    headers = {"Host": host}
    if token == "valid":
        headers["X-Cube-Token"] = app.store.token
    elif token:
        headers["X-Cube-Token"] = token
    if origin:
        headers["Origin"] = origin
    raw = json.dumps(body).encode() if body is not None else b""
    status, response_headers, payload = app.handle(method, path, headers, raw)
    content_type = response_headers.get("Content-Type", "")
    parsed = json.loads(payload) if content_type.startswith("application/json") else payload
    return status, parsed


def field(fields, unique_name):
    for dimension in fields["dimensions"]:
        for h in dimension["hierarchies"]:
            if h["uniqueName"] == unique_name:
                return {
                    "hierarchy": h["uniqueName"],
                    "caption": h["caption"],
                    "allMember": h["allMember"],
                    "expanded": [],
                    "members": [],
                    "memberCaptions": [],
                }
    raise AssertionError(unique_name)


# -- access control ----------------------------------------------------------------
def test_api_needs_the_token(app):
    status, body = call(app, "GET", "/api/status", token=None)
    assert status == 403 and body["error"]["category"] == "forbidden"
    assert call(app, "GET", "/api/status", token="not-the-token")[0] == 403


@pytest.mark.parametrize("host", ["evil.example:47613", "localhost:9999", "localhost", ""])
def test_requests_for_another_host_are_refused_even_with_the_token(app, host):
    assert call(app, "GET", "/api/status", host=host)[0] == 403
    assert call(app, "GET", "/", host=host, token=None)[0] == 403


def test_foreign_origin_is_refused_and_own_origin_is_accepted(app):
    assert call(app, "POST", "/api/connections/demo/cancel", {}, origin="https://evil.example")[0] == 403
    assert call(app, "POST", "/api/connections/demo/cancel", {}, origin=ORIGIN)[0] == 200
    assert call(app, "GET", "/api/status", host=f"127.0.0.1:{PORT}", origin=f"https://127.0.0.1:{PORT}")[0] == 200


def test_no_cors_headers_are_ever_sent(app):
    _, headers, _ = app.handle("GET", "/api/status", {"Host": HOST, "X-Cube-Token": app.store.token}, b"")
    assert not any(name.lower().startswith("access-control") for name in headers)


# -- static files ------------------------------------------------------------------
def test_index_gets_the_token_injected_and_is_not_cached(app):
    status, headers, body = app.handle("GET", "/", {"Host": HOST}, b"")
    assert status == 200 and headers["Content-Type"].startswith("text/html")
    assert app.store.token.encode() in body and b"__CUBE_TOKEN__" not in body
    assert headers["Cache-Control"] == "no-store"
    assert app.handle("GET", "/index.html?x=1", {"Host": HOST}, b"")[2] == body


def test_assets_are_served_with_a_type(app):
    status, headers, body = app.handle("GET", "/assets/app.js", {"Host": HOST}, b"")
    assert status == 200 and body == b"console.log('pane')"
    assert headers["Content-Type"].startswith("text/javascript")


@pytest.mark.parametrize("path", ["/../secret.txt", "/assets/../../secret.txt", "/%2e%2e/secret.txt", "//etc/passwd", "/assets/nope.js"])
def test_nothing_outside_the_pane_folder_is_served(app, path):
    status, _, body = app.handle("GET", path, {"Host": HOST}, b"")
    assert status == 404 and b"outside the pane folder" not in body


# -- routes ------------------------------------------------------------------------
def test_status(app):
    status, body = call(app, "GET", "/api/status")
    assert status == 200 and body["ok"] is True and body["version"] and body["certDaysLeft"] is None


def test_sample_cube_end_to_end(app):
    assert call(app, "GET", "/api/connections")[1] == [
        {"id": "demo", "name": "Sample cube", "kind": "demo", "server": "(built in)", "username": ""}
    ]
    assert call(app, "GET", "/api/connections/demo/catalogs")[1][0]["name"] == CATALOG
    assert call(app, "GET", f"/api/connections/demo/cubes?catalog={CATALOG}")[1] == [{"name": CUBE, "caption": CUBE}]

    status, fields = call(app, "GET", f"/api/connections/demo/fields?catalog={CATALOG}&cube=Ride%20Sales")
    assert status == 200 and fields["measures"] and fields["dimensions"]

    status, members = call(
        app,
        "POST",
        "/api/connections/demo/members",
        {"catalog": CATALOG, "cube": CUBE, "hierarchy": "[Geography].[City]", "limit": 3},
    )
    assert status == 200 and len(members["members"]) == 3 and members["truncated"] is True

    layout = {
        "rows": [field(fields, "[Geography].[City]")],
        "columns": [],
        "values": [{"measure": "[Measures].[Rides]", "caption": "Rides"}],
        "filters": [],
        "options": {"grandTotals": True, "subtotals": False, "nonEmpty": True},
    }
    status, grid = call(app, "POST", "/api/connections/demo/query", {"catalog": CATALOG, "cube": CUBE, "layout": layout})
    assert status == 200
    assert set(grid) == {
        "filterRows", "rowFieldCaptions", "colHeaderRows", "rowHeaders", "values",
        "formatPalette", "cellFormats", "rowKinds", "colKinds", "stats",
    }
    assert grid["rowHeaders"][-1][0]["text"] == "Grand Total"
    assert grid["stats"]["mdx"].startswith("SELECT") and grid["stats"]["elapsedMs"] >= 0


def test_errors_have_the_contract_shape_and_status(app):
    status, body = call(app, "GET", "/api/connections/nope/catalogs")
    assert status == 404 and body == {"error": {"category": "not-found", "message": body["error"]["message"]}}

    status, body = call(app, "POST", "/api/connections/demo/query", {"catalog": CATALOG, "cube": CUBE, "layout": {"rows": "x"}})
    assert status == 400 and body["error"]["category"] == "bad-request"

    status, body = call(app, "GET", "/api/connections/demo/cubes")
    assert status == 400 and "catalog" in body["error"]["message"].lower()

    assert call(app, "GET", "/api/nope")[0] == 404
    assert call(app, "PUT", "/api/status")[0] == 404


def test_too_large_result_is_413(app):
    app.cell_limit = 3
    status, fields = call(app, "GET", f"/api/connections/demo/fields?catalog={CATALOG}&cube=Ride%20Sales")
    layout = {"rows": [field(fields, "[Geography].[City]")], "values": [{"measure": "[Measures].[Rides]", "caption": "Rides"}]}
    status, body = call(app, "POST", "/api/connections/demo/query", {"catalog": CATALOG, "cube": CUBE, "layout": layout})
    assert status == 413 and body["error"]["category"] == "too-large"


def test_malformed_json_body_is_a_bad_request_not_a_crash(app):
    status, _, payload = app.handle(
        "POST", "/api/connections/demo/query", {"Host": HOST, "X-Cube-Token": app.store.token}, b"{nope"
    )
    assert status == 400 and json.loads(payload)["error"]["category"] == "bad-request"


def test_unexpected_exception_becomes_internal_without_leaking_details(app, monkeypatch):
    def boom(*_a, **_k):
        raise RuntimeError("secret internal path /Users/x")

    monkeypatch.setattr(app, "_status", boom)
    status, body = call(app, "GET", "/api/status")
    assert status == 500 and body["error"]["category"] == "internal"
    assert "secret internal path" not in json.dumps(body)


# -- connections -------------------------------------------------------------------
def test_adding_a_connection_tests_it_first_and_keeps_the_password_out_of_responses(app):
    new = {"name": "BI", "server": "bi-server", "username": "CORP\\dana", "password": "pw-123"}
    status, created = call(app, "POST", "/api/connections", new)
    assert status == 200
    assert created == {"id": created["id"], "name": "BI", "kind": "ssas", "server": "bi-server", "username": "CORP\\dana"}
    assert FakeBackend.instances[0].args == ("bi-server", "CORP\\dana", "pw-123")
    assert app.store.password(created["id"]) == "pw-123"
    assert "pw-123" not in json.dumps(call(app, "GET", "/api/connections")[1])

    # the tested backend is reused, not reconnected
    assert call(app, "GET", f"/api/connections/{created['id']}/catalogs")[1] == [{"name": "Sales", "description": ""}]
    assert len(FakeBackend.instances) == 1

    status, _ = call(app, "DELETE", f"/api/connections/{created['id']}")
    assert status == 200 and FakeBackend.instances[0].closed
    assert [c["id"] for c in call(app, "GET", "/api/connections")[1]] == ["demo"]


def test_a_connection_that_fails_its_test_is_not_saved(app):
    status, body = call(app, "POST", "/api/connections", {"server": "bi", "username": "u", "password": "wrong"})
    assert status == 401 and body["error"]["category"] == "auth"
    assert [c["id"] for c in app.store.connections()] == ["demo"]


@pytest.mark.parametrize(
    "body",
    [{}, {"server": "", "username": "u", "password": "p"}, {"server": "s", "username": "", "password": "p"}, {"server": "s", "username": "u", "password": ""}, {"server": 1, "username": "u", "password": "p"}],
)
def test_incomplete_connection_details_are_rejected(app, body):
    status, response = call(app, "POST", "/api/connections", body)
    assert status == 400 and response["error"]["category"] == "bad-request"
    assert FakeBackend.instances == []


def test_saved_connection_is_opened_with_its_keychain_password_after_restart(app, tmp_path):
    _, created = call(app, "POST", "/api/connections", {"server": "bi", "username": "u", "password": "pw-9"})
    restarted = App(app.store, port=PORT, pane_dir=tmp_path / "pane", backend_factory=FakeBackend, cert_path=None)
    assert call(restarted, "GET", f"/api/connections/{created['id']}/catalogs")[0] == 200
    assert FakeBackend.instances[-1].args == ("bi", "u", "pw-9")


def test_cancel_reaches_the_backend(app):
    _, created = call(app, "POST", "/api/connections", {"server": "bi", "username": "u", "password": "p"})
    assert call(app, "POST", f"/api/connections/{created['id']}/cancel", {})[1] == {"ok": True}
    assert FakeBackend.instances[0].cancelled == 1


# -- over a real HTTPS socket ------------------------------------------------------
def test_served_over_https_with_the_generated_certificate(tmp_path, app):
    paths = certs.generate(tmp_path / "certs")
    app.port = 0  # let the OS pick; Host header must then match the real port
    server = serve(app, certfile=paths.cert, keyfile=paths.key, port=0)
    port = server.server_address[1]
    app.port = port
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        context = ssl.create_default_context(cafile=str(paths.ca))
        # Go direct. A Mac with a system-wide proxy (a VPN client, for instance) would
        # otherwise send even https://localhost through it; Excel's web view does not.
        opener = urllib.request.build_opener(
            urllib.request.ProxyHandler({}), urllib.request.HTTPSHandler(context=context)
        )
        request = urllib.request.Request(
            f"https://localhost:{port}/api/status", headers={"X-Cube-Token": app.store.token}
        )
        with opener.open(request, timeout=10) as response:
            assert response.status == 200
            assert json.loads(response.read())["ok"] is True

        with pytest.raises(urllib.error.HTTPError) as info:
            opener.open(f"https://localhost:{port}/api/status", timeout=10)
        assert info.value.code == 403

        with opener.open(f"https://localhost:{port}/", timeout=10) as page:
            assert app.store.token.encode() in page.read()
    finally:
        server.shutdown()
        server.server_close()


# -- the real backend class behind the API, against the loopback stand-in server ------
def test_real_backend_through_the_api(tmp_path, ntlm_users):
    from fake_ssas import FakeSsas
    from test_ssas_backend import RESPONSES

    user, password = ntlm_users
    real = App(Store(tmp_path / "real", secrets=MemorySecrets()), port=PORT, pane_dir=None)
    with FakeSsas(dict(RESPONSES)) as server:
        address = f"127.0.0.1:{server.port}"
        status, body = call(real, "POST", "/api/connections", {"server": address, "username": user, "password": "wrong"})
        assert status == 401 and body["error"]["category"] == "auth"

        status, created = call(real, "POST", "/api/connections", {"server": address, "username": user, "password": password})
        assert status == 200 and created["server"] == address
        base = f"/api/connections/{created['id']}"

        assert [c["name"] for c in call(real, "GET", f"{base}/catalogs")[1]] == ["Archive", "Sales"]
        assert call(real, "GET", f"{base}/cubes?catalog=Sales")[1][0]["name"] == "Finance View"
        status, fields = call(real, "GET", f"{base}/fields?catalog=Sales&cube=Sales")
        assert status == 200
        layout = {"rows": [field(fields, "[Geo].[City]")], "values": [{"measure": "[Measures].[Rides]", "caption": "Rides"}]}
        status, grid = call(real, "POST", f"{base}/query", {"catalog": "Sales", "cube": "Sales", "layout": layout})
        assert status == 200
        assert [line[0]["text"] for line in grid["rowHeaders"]] == ["Tehran", "Grand Total"]
        assert grid["values"] == [[120], [200]]
        status, members = call(real, "POST", f"{base}/members", {"catalog": "Sales", "cube": "Sales", "hierarchy": "[Geo].[City]"})
        assert status == 200 and [m["caption"] for m in members["members"]] == ["Tehran", "Shiraz", "Tabriz"]
        real.close()
