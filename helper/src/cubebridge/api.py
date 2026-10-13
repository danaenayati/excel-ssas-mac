"""The local HTTPS service the task pane talks to.

It listens on 127.0.0.1 only and serves two things from one origin: the built task
pane (static files) and the JSON API under /api.

Three checks protect it from web pages open in the user's browser:

- `Host` must be localhost:<port> or 127.0.0.1:<port>. A page that re-points its own
  domain at 127.0.0.1 (DNS rebinding) still sends its own host name.
- `Origin`, when a browser sends one, must be this service's own origin.
- Every /api request must carry the install token, which only the served pane knows.

No CORS headers are sent, so no other origin can read a response either.

`App.handle` is plain data in, plain data out, so it is tested without sockets; the
HTTP server at the bottom is a thin wrapper.
"""

from __future__ import annotations

import hmac
import json
import logging
import re
import ssl
import threading
import time
from collections.abc import Callable
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlsplit

from . import __version__, certs
from .backends.base import DEFAULT_MEMBER_LIMIT, MAX_MEMBER_LIMIT
from .backends.demo_backend import DemoBackend
from .backends.ssas_backend import SsasBackend
from .errors import CubeError
from .pivot.grid import DEFAULT_CELL_LIMIT, build_grid
from .pivot.model import Layout
from .pivot.names import validate_unique_name
from .store import DEMO_ID, Store

log = logging.getLogger("cubebridge")

TOKEN_PLACEHOLDER = b"__CUBE_TOKEN__"
MAX_BODY = 5 * 1024 * 1024
MAX_TEXT = 400

_CONTENT_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".mjs": "text/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".map": "application/json; charset=utf-8",
    ".svg": "image/svg+xml",
    ".png": "image/png",
    ".ico": "image/x-icon",
    ".woff2": "font/woff2",
    ".txt": "text/plain; charset=utf-8",
}

_ROUTES: list[tuple[str, re.Pattern, str]] = [
    ("GET", re.compile(r"^/api/status$"), "_status"),
    ("GET", re.compile(r"^/api/connections$"), "_list_connections"),
    ("POST", re.compile(r"^/api/connections$"), "_add_connection"),
    ("DELETE", re.compile(r"^/api/connections/(?P<id>[A-Za-z0-9_-]{1,64})$"), "_delete_connection"),
    ("GET", re.compile(r"^/api/connections/(?P<id>[A-Za-z0-9_-]{1,64})/catalogs$"), "_catalogs"),
    ("GET", re.compile(r"^/api/connections/(?P<id>[A-Za-z0-9_-]{1,64})/cubes$"), "_cubes"),
    ("GET", re.compile(r"^/api/connections/(?P<id>[A-Za-z0-9_-]{1,64})/fields$"), "_fields"),
    ("POST", re.compile(r"^/api/connections/(?P<id>[A-Za-z0-9_-]{1,64})/members$"), "_members"),
    ("POST", re.compile(r"^/api/connections/(?P<id>[A-Za-z0-9_-]{1,64})/query$"), "_query"),
    ("POST", re.compile(r"^/api/connections/(?P<id>[A-Za-z0-9_-]{1,64})/cancel$"), "_cancel"),
]

Response = tuple[int, dict[str, str], bytes]


def _bad(message: str) -> CubeError:
    return CubeError("bad-request", message)


def _text(source: dict, name: str, *, required: bool = True, limit: int = MAX_TEXT) -> str:
    value = source.get(name)
    if value is None or value == "":
        if required:
            raise _bad(f"The request is missing '{name}'.")
        return ""
    if not isinstance(value, str) or len(value) > limit:
        raise _bad(f"'{name}' is not valid.")
    return value


class App:
    def __init__(
        self,
        store: Store,
        *,
        port: int,
        pane_dir: Path | None,
        backend_factory: Callable = SsasBackend,
        cert_path: Path | None = None,
        cell_limit: int = DEFAULT_CELL_LIMIT,
    ) -> None:
        self.store = store
        self.port = port
        self.pane_dir = Path(pane_dir) if pane_dir else None
        self.cell_limit = cell_limit
        self._backend_factory = backend_factory
        self._cert_path = cert_path
        self._backends: dict[str, object] = {}
        self._lock = threading.Lock()

    # -- entry point -------------------------------------------------------------
    def handle(self, method: str, target: str, headers: dict[str, str], body: bytes) -> Response:
        headers = {name.lower(): value for name, value in headers.items()}
        try:
            try:
                parts = urlsplit(target)
                path = unquote(parts.path)
            except ValueError as exc:
                raise _bad("The request address is not valid.") from exc
            self._check_host(headers)
            if not path.startswith("/api/"):
                return self._static(method, path)
            self._check_origin_and_token(headers)
            for route_method, pattern, handler_name in _ROUTES:
                match = pattern.match(path)
                if match and route_method == method:
                    query = {k: v[-1] for k, v in parse_qs(parts.query, keep_blank_values=True).items()}
                    payload = self._json_body(body) if method == "POST" else {}
                    result = getattr(self, handler_name)(query=query, body=payload, **match.groupdict())
                    return self._json(200, result)
            raise CubeError("not-found", "The helper does not know that request.")
        except CubeError as exc:
            # Category and detail go to the log so a failure can be diagnosed from it.
            # Neither ever holds a password or a sign-in token.
            if exc.category != "cancelled":
                log.warning("%s request failed: %s: %s%s", method, exc.category, exc.message, f" [{exc.detail}]" if exc.detail else "")
            return self._json(exc.http_status, {"error": exc.to_json()})
        except Exception:
            log.exception("unhandled error in %s request", method)
            error = CubeError("internal", "Something went wrong inside the helper. See its log for details.")
            return self._json(500, {"error": error.to_json()})

    # -- access control ----------------------------------------------------------
    def _allowed_hosts(self) -> set[str]:
        return {f"localhost:{self.port}", f"127.0.0.1:{self.port}"}

    def _check_host(self, headers: dict[str, str]) -> None:
        if headers.get("host", "").lower() not in self._allowed_hosts():
            raise CubeError("forbidden", "This helper only answers requests made to localhost.")

    def _check_origin_and_token(self, headers: dict[str, str]) -> None:
        origin = headers.get("origin")
        if origin is not None and origin.lower() not in {f"https://{h}" for h in self._allowed_hosts()}:
            raise CubeError("forbidden", "This helper only answers requests from its own task pane.")
        supplied = headers.get("x-cube-token", "")
        if not supplied or not hmac.compare_digest(supplied.encode(), self.store.token.encode()):
            raise CubeError("forbidden", "The request did not come from the Cube task pane.")

    # -- responses ---------------------------------------------------------------
    @staticmethod
    def _json(status: int, data: object) -> Response:
        # backslashreplace: text that is not valid Unicode (a lone surrogate echoed
        # back in an error message) must not stop a response from being sent.
        payload = json.dumps(data, ensure_ascii=False, allow_nan=False, separators=(",", ":")).encode(
            "utf-8", "backslashreplace"
        )
        return status, {"Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store"}, payload

    @staticmethod
    def _json_body(body: bytes) -> dict:
        if not body:
            return {}
        try:
            data = json.loads(body)
        except ValueError as exc:
            raise _bad("The request body is not valid JSON.") from exc
        if not isinstance(data, dict):
            raise _bad("The request body must be a JSON object.")
        return data

    def _static(self, method: str, path: str) -> Response:
        missing = (404, {"Content-Type": "text/plain; charset=utf-8", "Cache-Control": "no-store"}, b"Not found")
        if method != "GET" or self.pane_dir is None:
            return missing
        relative = path.lstrip("/") or "index.html"
        root = self.pane_dir.resolve()
        try:
            candidate = (root / relative).resolve()
        except (OSError, ValueError):
            return missing
        if not candidate.is_relative_to(root) or not candidate.is_file():
            return missing
        data = candidate.read_bytes()
        headers = {
            "Content-Type": _CONTENT_TYPES.get(candidate.suffix.lower(), "application/octet-stream"),
            "Cache-Control": "no-cache",
            "X-Content-Type-Options": "nosniff",
            "Referrer-Policy": "no-referrer",
        }
        if candidate.suffix.lower() == ".html":
            data = data.replace(TOKEN_PLACEHOLDER, self.store.token.encode())
            headers["Cache-Control"] = "no-store"
        return 200, headers, data

    # -- backends ----------------------------------------------------------------
    def _backend(self, connection_id: str):
        with self._lock:
            backend = self._backends.get(connection_id)
            if backend is not None:
                return backend
            if connection_id == DEMO_ID:
                backend = DemoBackend()
            else:
                connection = self.store.get(connection_id)
                backend = self._backend_factory(
                    connection["server"], connection["username"], self.store.password(connection_id)
                )
            self._backends[connection_id] = backend
            return backend

    def close(self) -> None:
        with self._lock:
            backends = list(self._backends.values())
            self._backends.clear()
        for backend in backends:  # outside the lock: closing may wait on the network
            backend.close()

    # -- routes ------------------------------------------------------------------
    def _status(self, **_):
        return {
            "ok": True,
            "version": __version__,
            "certDaysLeft": certs.days_left(self._cert_path) if self._cert_path else None,
        }

    def _list_connections(self, **_):
        return self.store.connections()

    def _add_connection(self, body: dict, **_):
        server = _text(body, "server").strip()
        username = _text(body, "username").strip()
        password = _text(body, "password", limit=1024)
        name = _text(body, "name", required=False)
        backend = self._backend_factory(server, username, password)
        try:
            backend.catalogs()  # proves the server is reachable and the sign-in works
        except BaseException:
            backend.close()
            raise
        connection = self.store.add(name=name, server=server, username=username, password=password)
        with self._lock:
            self._backends[connection["id"]] = backend
        return connection

    def _delete_connection(self, id: str, **_):
        self.store.delete(id)
        with self._lock:
            backend = self._backends.pop(id, None)
        if backend is not None:
            backend.close()
        return {"ok": True}

    def _catalogs(self, id: str, **_):
        return self._backend(id).catalogs()

    def _cubes(self, id: str, query: dict, **_):
        return self._backend(id).cubes(_text(query, "catalog"))

    def _fields(self, id: str, query: dict, **_):
        return self._backend(id).fields(_text(query, "catalog"), _text(query, "cube"))

    def _members(self, id: str, body: dict, **_):
        limit = body.get("limit", DEFAULT_MEMBER_LIMIT)
        if not isinstance(limit, int) or isinstance(limit, bool):
            raise _bad("'limit' must be a whole number.")
        parent = _text(body, "parent", required=False, limit=1000)
        return self._backend(id).members(
            _text(body, "catalog"),
            _text(body, "cube"),
            validate_unique_name(body.get("hierarchy")),
            validate_unique_name(parent) if parent else None,
            _text(body, "search", required=False, limit=200) or None,
            max(1, min(limit, MAX_MEMBER_LIMIT)),
        )

    def _query(self, id: str, body: dict, **_):
        catalog = _text(body, "catalog")
        cube = _text(body, "cube")
        layout = Layout.from_json(body.get("layout"))
        started = time.monotonic()
        cellset, text = self._backend(id).query(catalog, cube, layout)
        elapsed = int((time.monotonic() - started) * 1000)
        return build_grid(cellset, layout, elapsed_ms=elapsed, mdx=text, cell_limit=self.cell_limit)

    def _cancel(self, id: str, **_):
        with self._lock:
            backend = self._backends.get(id)
        if backend is not None:
            backend.cancel()
        return {"ok": True}


class _Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    server_version = "CubeBridge"
    sys_version = ""
    timeout = 120  # an idle or stalled connection is dropped, never held forever

    def _dispatch(self) -> None:
        started = time.monotonic()
        try:
            length = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            length = -1
        if length < 0 or length > MAX_BODY:
            self.send_error(413)
            return
        body = self.rfile.read(length) if length else b""
        status, headers, payload = self.server.app.handle(self.command, self.path, dict(self.headers.items()), body)
        self.send_response(status)
        for name, value in headers.items():
            self.send_header(name, value)
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)
        # Path only: the query string can carry database and cube names.
        log.info("%s %s -> %s in %d ms", self.command, urlsplit(self.path).path, status, (time.monotonic() - started) * 1000)

    do_GET = do_POST = do_DELETE = do_PUT = do_PATCH = do_OPTIONS = do_HEAD = _dispatch

    def log_message(self, format: str, *args) -> None:  # noqa: A002 - signature is fixed by the base class
        # Silenced: the base class logs the full request line, query string included.
        # `_dispatch` logs each request by path instead.
        pass


class _Server(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True
    app: App


def serve(app: App, *, certfile: Path, keyfile: Path, port: int) -> _Server:
    """Bind to loopback with TLS. The caller runs `serve_forever()`."""
    server = _Server(("127.0.0.1", port), _Handler)
    server.app = app
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.minimum_version = ssl.TLSVersion.TLSv1_2
    context.load_cert_chain(certfile=str(certfile), keyfile=str(keyfile))
    # Handshake in the request thread, so one slow client cannot stall accept().
    server.socket = context.wrap_socket(server.socket, server_side=True, do_handshake_on_connect=False)
    return server
