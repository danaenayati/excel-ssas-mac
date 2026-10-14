"""The cube formulas' routes: access rules (Origin, token, CORS) and answers."""

import http.client
import json
import ssl
import threading

import pytest

from cubebridge import certs
from cubebridge.api import App, serve
from cubebridge.backends.demo_backend import CATALOG, CUBE
from cubebridge.store import MemorySecrets, Store

PORT = 47613
HOST = f"localhost:{PORT}"
ORIGIN = f"https://{HOST}"
TEHRAN = "[Geography].[City].[City].&[Tehran]"
RIDES = "[Measures].[Rides]"


class TwoDatabases:
    """A server connection with two databases, each with two cubes."""

    def __init__(self, server, username, password):
        self.asked = []

    def catalogs(self):
        return [{"name": "Sales", "description": ""}, {"name": "Finance", "description": ""}]

    def cubes(self, catalog):
        return [{"name": "Model", "caption": "Model"}, {"name": "Other", "caption": "Other"}]

    def cells(self, catalog, cube, tuples):
        self.asked.append((catalog, cube, tuples))
        return [{"value": 1} for _ in tuples]

    def cancel(self):
        pass

    def close(self):
        pass


@pytest.fixture
def app(tmp_path):
    pane = tmp_path / "pane"
    (pane / "assets").mkdir(parents=True)
    (pane / "index.html").write_text('<meta name="cube-token" content="__CUBE_TOKEN__">')
    (pane / "functions.html").write_text('<meta name="cube-token" content="__CUBE_TOKEN__"><script src="functions.js"></script>')
    (pane / "functions.js").write_text('var t="__CUBE_TOKEN__";CustomFunctions.associate("VALUE",f)')
    (pane / "functions.json").write_text('{"functions": []}')
    (pane / "assets" / "functions.js").write_text('var t="__CUBE_TOKEN__"')
    (pane / "assets" / "app.js").write_text('var t="__CUBE_TOKEN__"')
    store = Store(tmp_path / "app", secrets=MemorySecrets())
    return App(store, port=PORT, pane_dir=pane, backend_factory=TwoDatabases, cert_path=None)


def send(app, method, path, body=None, *, header_token=False, body_token=False, origin=None, extra=None):
    headers = {"Host": HOST, "Content-Type": "text/plain;charset=UTF-8"}
    if header_token:
        headers["X-Cube-Token"] = app.store.token
    if origin is not None:
        headers["Origin"] = origin
    headers.update(extra or {})
    if body is not None and body_token:
        body = {**body, "token": app.store.token}
    raw = json.dumps(body).encode() if body is not None else b""
    status, response_headers, payload = app.handle(method, path, headers, raw)
    parsed = json.loads(payload) if response_headers.get("Content-Type", "").startswith("application/json") else payload
    return status, response_headers, parsed


CELLS = {"catalog": CATALOG, "cube": CUBE, "tuples": [[TEHRAN, RIDES], ["Tehran"]]}


# -- answers -----------------------------------------------------------------------
def test_cells_from_the_sample_cube(app):
    status, _, body = send(app, "POST", "/api/connections/demo/cells", CELLS, header_token=True)
    assert status == 200
    first, second = body["results"]
    assert isinstance(first["value"], int) and first["value"] > 0
    assert second["error"]["category"] == "bad-request"


def test_the_only_database_and_cube_are_used_when_none_is_named(app):
    status, _, body = send(app, "POST", "/api/connections/demo/cells", {"tuples": [[TEHRAN]]}, body_token=True)
    assert status == 200 and body["results"][0]["value"] > 0


def test_a_connection_with_several_databases_must_name_one(app):
    app.store.add(name="Sales Cube", server="bi", username="u", password="p")
    connection = [c for c in app.store.connections() if c["name"] == "Sales Cube"][0]
    path = f"/api/connections/{connection['id']}/cells"
    status, _, body = send(app, "POST", path, {"tuples": [[TEHRAN]]}, body_token=True)
    assert status == 400
    message = body["error"]["message"]
    assert "Sales, Finance" in message and '"Sales Cube/Database/Cube"' in message
    status, _, body = send(app, "POST", path, {"catalog": "Sales", "cube": "Model", "tuples": [[TEHRAN]]}, body_token=True)
    assert status == 200 and body == {"results": [{"value": 1}]}


def test_lookups_from_the_sample_cube(app):
    items = [
        {"kind": "member", "members": [TEHRAN]},
        {"kind": "count", "set": "[Geography].[City].Members"},
        {"kind": "ranked", "set": "{" + TEHRAN + "}", "rank": 1},
        {"kind": "property", "member": TEHRAN, "property": "Key"},
    ]
    status, _, body = send(app, "POST", "/api/connections/demo/lookup", {"items": items}, body_token=True)
    assert status == 200
    assert [r["value"] for r in body["results"]] == ["Tehran", 10, "Tehran", "Tehran"]


def test_connection_names_for_formulas(app):
    status, _, body = send(app, "POST", "/api/formulas/connections", {}, body_token=True)
    assert status == 200 and body == {"connections": [{"id": "demo", "name": "Sample cube"}]}


@pytest.mark.parametrize(
    "body",
    [
        {"tuples": "x"},
        {"tuples": [[1]]},
        {"tuples": [[]] * 5001},
        {"catalog": 5, "tuples": []},
    ],
)
def test_malformed_cells_requests_are_bad_requests(app, body):
    status, headers, response = send(app, "POST", "/api/connections/demo/cells", body, body_token=True, origin="null")
    assert status == 400 and response["error"]["category"] == "bad-request"
    # Errors carry the CORS header too, so that the formula can show the message.
    assert headers["Access-Control-Allow-Origin"] == "null"


def test_an_unknown_connection_is_not_found(app):
    status, _, _ = send(app, "POST", "/api/connections/nope/cells", {"tuples": []}, body_token=True)
    assert status == 404


# -- access ------------------------------------------------------------------------
@pytest.mark.parametrize("origin", [None, "null", ORIGIN, f"https://127.0.0.1:{PORT}"])
def test_the_functions_runtime_may_send_no_origin_null_or_ours(app, origin):
    status, headers, _ = send(app, "POST", "/api/connections/demo/cells", CELLS, body_token=True, origin=origin)
    assert status == 200
    if origin is None:
        assert not any(name.lower().startswith("access-control") for name in headers)
    else:
        assert headers["Access-Control-Allow-Origin"] == origin and headers["Vary"] == "Origin"


@pytest.mark.parametrize("origin", ["https://evil.example", "http://localhost:47613", "https://localhost:9999", "Null "])
def test_other_origins_are_refused_even_with_the_token(app, origin):
    status, headers, _ = send(app, "POST", "/api/connections/demo/cells", CELLS, header_token=True, body_token=True, origin=origin)
    assert status == 403
    assert "Access-Control-Allow-Origin" not in headers


def test_the_token_is_required_on_every_formula_request(app):
    for path, body in (("/api/connections/demo/cells", CELLS), ("/api/connections/demo/lookup", {"items": []}), ("/api/formulas/connections", {})):
        assert send(app, "POST", path, body, origin="null")[0] == 403
        assert send(app, "POST", path, {**body, "token": "wrong"}, origin="null")[0] == 403
        assert send(app, "POST", path, {**body, "token": ["not text"]})[0] == 403
        assert send(app, "POST", path, {**body, "token": "\ud800"})[0] == 403
    status, _, _ = app.handle("POST", "/api/connections/demo/cells", {"Host": HOST}, b"{not json")
    assert status == 403
    # The header, when there is one, is the token that counts.
    status, _, _ = send(app, "POST", "/api/connections/demo/cells", CELLS, body_token=True, extra={"X-Cube-Token": "wrong"})
    assert status == 403


def test_a_preflight_is_answered_without_the_token(app):
    status, headers, body = send(app, "OPTIONS", "/api/connections/demo/cells", origin="null")
    assert status == 204 and body == b""
    assert headers["Access-Control-Allow-Origin"] == "null"
    assert headers["Access-Control-Allow-Methods"] == "POST"
    assert headers["Access-Control-Allow-Headers"] == "Content-Type, X-Cube-Token"
    assert send(app, "OPTIONS", "/api/connections/demo/cells", origin="https://evil.example")[0] == 403
    assert send(app, "OPTIONS", "/api/connections/demo/cells", origin="null", extra={"Host": "evil.example:47613"})[0] == 403


def test_the_other_routes_stay_as_strict_as_before(app):
    # Origin "null", a token in the body, or a preflight: none of it opens the pane's routes.
    assert send(app, "GET", "/api/status", header_token=True, origin="null")[0] == 403
    body = {"catalog": CATALOG, "cube": CUBE, "layout": {}}
    assert send(app, "POST", "/api/connections/demo/query", body, body_token=True)[0] == 403
    assert send(app, "OPTIONS", "/api/connections/demo/query", origin="null")[0] == 403
    _, headers, _ = send(app, "GET", "/api/status", header_token=True)
    assert not any(name.lower().startswith("access-control") for name in headers)


# -- the functions script ----------------------------------------------------------
def test_the_functions_files_get_the_token_and_are_not_cached(app):
    for path in ("/functions.js", "/functions.html"):
        status, headers, body = app.handle("GET", path, {"Host": HOST}, b"")
        assert status == 200 and app.store.token.encode() in body and b"__CUBE_TOKEN__" not in body
        assert headers["Cache-Control"] == "no-store"
    status, headers, body = app.handle("GET", "/functions.json", {"Host": HOST}, b"")
    assert status == 200 and headers["Content-Type"].startswith("application/json")
    # Nowhere else: not another script, not one of the same name in another folder.
    for path in ("/assets/app.js", "/assets/functions.js"):
        assert app.handle("GET", path, {"Host": HOST}, b"")[2] == b'var t="__CUBE_TOKEN__"'


@pytest.mark.parametrize(
    "headers",
    [
        {"Sec-Fetch-Site": "same-origin", "Sec-Fetch-Dest": "script"},
        {"Sec-Fetch-Site": "none"},
        {"Referer": f"{ORIGIN}/functions.html"},
        {},
    ],
)
def test_excel_may_load_the_functions_script(app, headers):
    status, _, body = app.handle("GET", "/functions.js", {"Host": HOST, **headers}, b"")
    assert status == 200 and app.store.token.encode() in body


@pytest.mark.parametrize(
    "headers",
    [
        {"Sec-Fetch-Site": "cross-site", "Sec-Fetch-Dest": "script"},
        {"Sec-Fetch-Site": "same-site"},  # another port on localhost
        {"Referer": "https://evil.example/page.html"},
        {"Referer": "http://localhost:47613/"},
        {"Referer": "http://[bad"},
    ],
)
def test_another_site_cannot_load_the_functions_script(app, headers):
    status, _, body = app.handle("GET", "/functions.js", {"Host": HOST, **headers}, b"")
    assert status == 403 and app.store.token.encode() not in body


# -- over a real HTTPS socket ------------------------------------------------------
def test_preflight_and_simple_request_over_https(tmp_path, app):
    paths = certs.generate(tmp_path / "certs")
    server = serve(app, certfile=paths.cert, keyfile=paths.key, port=0)
    port = server.server_address[1]
    app.port = port
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        context = ssl.create_default_context(cafile=str(paths.ca))
        connection = http.client.HTTPSConnection("localhost", port, context=context, timeout=10)
        connection.request("OPTIONS", "/api/connections/demo/cells", headers={"Origin": "null", "Access-Control-Request-Method": "POST"})
        response = connection.getresponse()
        assert response.status == 204 and response.read() == b""
        assert response.getheader("Access-Control-Allow-Origin") == "null"
        body = json.dumps({"token": app.store.token, "tuples": [[TEHRAN]]})
        connection.request("POST", "/api/connections/demo/cells", body=body, headers={"Origin": "null", "Content-Type": "text/plain"})
        response = connection.getresponse()
        assert response.status == 200 and json.loads(response.read())["results"][0]["value"] > 0
        connection.close()
    finally:
        server.shutdown()
        server.server_close()
