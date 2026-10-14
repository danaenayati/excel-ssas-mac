"""A loopback stand-in for msmdpump.dll behind IIS.

SOAP over HTTP/1.1 with keep-alive, and the sign-in IIS asks for: Basic, NTLM and
Negotiate, each in a 401's WWW-Authenticate. NTLM is played by pyspnego's own
acceptor, which signs in a connection, not a request, as IIS does. Answers come from
canned XML, as in FakeSsas. Optionally served over TLS.
"""

from __future__ import annotations

import base64
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import spnego
import xml_samples as x


class _Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    timeout = 30

    def setup(self) -> None:
        super().setup()
        self.fake: FakeXmlaHttp = self.server.fake
        with self.fake.lock:
            self.fake.connections += 1
            self.connection_number = self.fake.connections
        self.signed_in = False
        self.context = None

    def log_message(self, *_args) -> None:
        pass

    def do_POST(self) -> None:  # noqa: N802 - name fixed by the base class
        fake = self.fake
        body = self.rfile.read(int(self.headers.get("Content-Length") or 0))
        record = {
            "connection": self.connection_number,
            "path": self.path,
            "headers": {k.lower(): v for k, v in self.headers.items()},
            "body": body,
            "text": body.decode("utf-8", "replace"),
        }
        with fake.lock:
            fake.requests.append(record)
        if fake.status_override is not None:
            status, headers, payload = fake.status_override
            self._send(status, payload, headers)
            return
        if self.path != fake.path:
            self._send(404, b"<html>Not Found</html>", {"Content-Type": "text/html"})
            return
        extra = {}
        if fake.offers and not self.signed_in:
            outcome = self._authenticate(record)
            if outcome is None:
                return  # a 401 went out
            extra = outcome
        with fake.lock:
            fake.served.append(record)
        if fake.delay:
            time.sleep(fake.delay)
        payload = fake.respond(record["text"])
        status = 500 if b":Fault" in payload else 200
        headers = {"Content-Type": "text/xml; charset=utf-8", **extra}
        if fake.close_after_response:
            headers["Connection"] = "close"
        self._send(status, payload, headers)
        if fake.per_request_sign_in:
            self.signed_in = False
            self.context = None

    def _send(self, status: int, payload: bytes, headers: dict[str, str] | list[tuple[str, str]]) -> None:
        self.send_response(status)
        items = headers.items() if isinstance(headers, dict) else headers
        for name, value in items:
            self.send_header(name, value)
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)
        if any(name.lower() == "connection" and value.lower() == "close" for name, value in items):
            self.close_connection = True

    def _challenge(self, token: tuple[str, bytes] | None = None) -> None:
        if token is not None:
            offers = [("WWW-Authenticate", f"{token[0]} {base64.b64encode(token[1]).decode()}")]
        else:
            offers = [
                ("WWW-Authenticate", 'Basic realm="bi.corp.example.com"' if o.lower() == "basic" else o)
                for o in self.fake.offers
            ]
        if self.fake.close_after_401:
            offers.append(("Connection", "close"))
        self._send(401, b"<html>401 - Unauthorized</html>", [("Content-Type", "text/html"), *offers])

    def _authenticate(self, record: dict) -> dict | None:
        """Headers for the answer when the request is signed in; None after sending a 401."""
        fake = self.fake
        header = self.headers.get("Authorization")
        if not header:
            self._challenge()
            return None
        scheme, _, token = header.partition(" ")
        record["scheme"] = scheme
        if scheme.lower() not in (o.lower() for o in fake.offers):
            self._challenge()
            return None
        if scheme.lower() == "basic":
            user, _, password = base64.b64decode(token).decode("utf-8").partition(":")
            if fake.basic_users.get(user) == password:
                return {}  # Basic signs in the request, not the connection
            self._challenge()
            return None
        if self.context is None:
            self.context = fake.acceptor(scheme)
            with fake.lock:
                fake.contexts.append(self.context)
        try:
            out = self.context.step(base64.b64decode(token))
        except Exception as exc:  # noqa: BLE001 - a refused sign-in, as IIS answers it
            fake.refusals.append(type(exc).__name__)
            self.context = None
            self._challenge()
            return None
        if self.context.complete:
            self.signed_in = True
            return {"WWW-Authenticate": f"{scheme} {base64.b64encode(out).decode()}"} if out else {}
        self._challenge((scheme, out))
        return None


class _Server(ThreadingHTTPServer):
    daemon_threads = True
    fake: FakeXmlaHttp

    def handle_error(self, request, client_address) -> None:
        # A client that cancels, or refuses the certificate, drops the connection on
        # purpose in these tests; that is not worth a traceback.
        pass


class FakeXmlaHttp:
    """Start with `with FakeXmlaHttp(...) as server:`; its address is `server.url`."""

    def __init__(
        self,
        responses=None,
        *,
        offers=("Negotiate", "NTLM"),
        basic_users=None,
        acceptor=None,
        tls_context=None,
        path: str = "/olap/msmdpump.dll",
    ) -> None:
        self.responses = responses or {}
        self.offers = list(offers)  # what a 401 offers; empty: anyone may ask
        self.basic_users = dict(basic_users or {})
        self.acceptor = acceptor or (lambda scheme: spnego.server(protocol="ntlm"))
        self.path = path
        self.delay = 0.0
        self.close_after_401 = False
        self.close_after_response = False
        self.per_request_sign_in = False
        self.status_override: tuple[int, dict, bytes] | None = None
        self.lock = threading.Lock()
        self.requests: list[dict] = []  # every request, 401s included
        self.served: list[dict] = []  # the requests that got an answer
        self.contexts: list = []
        self.refusals: list[str] = []
        self.connections = 0
        self._server = _Server(("127.0.0.1", 0), _Handler)
        self._server.fake = self
        if tls_context is not None:
            self._server.socket = tls_context.wrap_socket(self._server.socket, server_side=True)
        self.port = self._server.server_address[1]
        scheme = "https" if tls_context is not None else "http"
        self.url = f"{scheme}://127.0.0.1:{self.port}{path}"
        self._thread = threading.Thread(target=self._server.serve_forever, kwargs={"poll_interval": 0.05}, daemon=True)

    def __enter__(self) -> FakeXmlaHttp:
        self._thread.start()
        return self

    def __exit__(self, *exc) -> None:
        self._server.shutdown()
        self._server.server_close()

    def respond(self, text: str) -> bytes:
        for marker, response in self.responses.items():
            if marker in text:
                return response(text) if callable(response) else response
        return x.soap_fault(f"The fake server has no canned response for this request: {text[:200]}")
