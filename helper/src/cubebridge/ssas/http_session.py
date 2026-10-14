"""XMLA over HTTP(S): an msmdpump.dll endpoint, as Windows Excel reaches it by address.

Every request is one SOAP 1.1 POST (Content-Type text/xml and a SOAPAction header)
with the same envelopes as the TCP binding, BeginSession and Session headers included,
and every answer is plain XML, read by xmla.py. Nothing is sealed: https protects the
connection (http:// sends XML in the clear, as Windows Excel would).

Sign-in follows the server's 401 answer and its WWW-Authenticate offers:

- "basic": only over https, and only once the server has asked for Basic; from then on
  the header goes with every request.
- "ntlm": the NTLM handshake in Authorization / WWW-Authenticate, under "NTLM", or
  under "Negotiate" when that is all the server offers. It signs in the TCP connection,
  not a request, so the handshake and the requests after it stay on one connection.
- "kerberos": "Negotiate" with SPNEGO tokens that offer Kerberos only (kerberos.py).

A kind the server does not offer is an error naming the choice that would work; the
password is never sent another way. Over https, NTLM and Kerberos tokens carry the
certificate's channel binding (IIS "Extended Protection"), as Windows sends it.

No proxy is ever used: http.client connects directly, whatever the environment or
System Settings say (a VPN client may set a system proxy), as the TCP binding does.

TLS is verified against the system trust store through truststore (the macOS Keychain,
so a company root certificate that IT installed is trusted as in Safari). Without
truststore, Python's own list of public authorities is used.

The interface is the TCP Session's: open, close, abort, is_open, discover, execute and
execute_rowset. `abort` shuts the socket down from another thread to cancel.
"""

from __future__ import annotations

import base64
import binascii
import http.client
import logging
import re
import select
import socket
import ssl
import threading
from collections.abc import Callable, Mapping
from urllib.parse import urlsplit

from .. import __version__
from ..errors import CubeError, ProtocolError
from ..pivot.model import CellSet
from . import auth, envelopes, xmla
from .resolve import HttpTarget

log = logging.getLogger("cubebridge.ssas")

SERVICE_CLASS = "HTTP"
MAX_RESPONSE = 256 * 1024 * 1024
MAX_HANDSHAKE_ROUNDS = 10
_READ_CHUNK = 1024 * 1024
_ACTIONS = {
    "discover": '"urn:schemas-microsoft-com:xml-analysis:Discover"',
    "execute": '"urn:schemas-microsoft-com:xml-analysis:Execute"',
}
_DECLARATION = b'<?xml version="1.0" encoding="utf-8"?>'
_SCHEME = re.compile(r"[A-Za-z][A-Za-z0-9!#$%&'*+.^_`|~-]*")
_SCHEME_NAMES = {"ntlm": "NTLM", "negotiate": "Negotiate", "kerberos": "Kerberos", "basic": "Basic"}


# -- WWW-Authenticate ----------------------------------------------------------------
def parse_challenges(values: list[str]) -> dict[str, str]:
    """{scheme in lower case: what follows it}, from every WWW-Authenticate header.

    "Negotiate, NTLM" and separate headers both list two schemes; "Negotiate <token>"
    carries a handshake token; Basic's realm="..." parameters are kept as text.
    """
    found: dict[str, str] = {}
    for value in values:
        for piece in _split_outside_quotes(value):
            piece = piece.strip()
            if not piece:
                continue
            word, _, rest = piece.partition(" ")
            if "=" in word and not rest:
                continue  # a parameter of the scheme before it (realm="x", charset=...)
            if _SCHEME.fullmatch(word):
                found.setdefault(word.lower(), rest.strip())
    return found


def _split_outside_quotes(value: str) -> list[str]:
    parts, current, quoted = [], [], False
    for ch in value:
        if ch == '"':
            quoted = not quoted
        if ch == "," and not quoted:
            parts.append("".join(current))
            current = []
        else:
            current.append(ch)
    parts.append("".join(current))
    return parts


def _decode_token(text: str) -> bytes:
    try:
        return base64.b64decode("".join(text.split()), validate=True)
    except (binascii.Error, ValueError) as exc:
        raise ProtocolError("the server's sign-in token is not valid base64") from exc


# -- TLS -------------------------------------------------------------------------------
def default_ssl_context() -> ssl.SSLContext:
    try:
        import truststore

        context = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    except ImportError:
        context = ssl.create_default_context()
    context.minimum_version = ssl.TLSVersion.TLSv1_2
    return context


def channel_bindings(sock: ssl.SSLSocket):
    """RFC 5929 tls-server-end-point: the hash of the server's certificate, which IIS
    checks under Extended Protection. SHA-256 unless the certificate is signed with a
    stronger hash."""
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes
    from spnego.channel_bindings import GssChannelBindings

    der = sock.getpeercert(binary_form=True)
    if not der:
        return None
    try:
        algorithm = x509.load_der_x509_certificate(der).signature_hash_algorithm
    except Exception:  # noqa: BLE001 - an unusual signature: fall back to SHA-256
        algorithm = None
    if algorithm is None or algorithm.name in ("md5", "sha1"):
        algorithm = hashes.SHA256()
    digest = hashes.Hash(algorithm)
    digest.update(der)
    return GssChannelBindings(application_data=b"tls-server-end-point:" + digest.finalize())


def _default_context(kind: str, username: str, password: str, host: str, bindings):
    if kind == "kerberos":
        from . import kerberos

        return kerberos.client_context(username, password, host, service=SERVICE_CLASS, channel_bindings=bindings)
    return auth.ntlm_context(username, password, host, SERVICE_CLASS, bindings)


def _dropped(sock: socket.socket) -> bool:
    """True when the server has closed this idle connection.

    Peeks at the raw bytes (below TLS, which they stay in the buffer for): nothing at
    all means the server closed it. TLS records that are waiting (TLS 1.3 sends session
    tickets after the handshake) do not mean that.
    """
    try:
        readable, _, _ = select.select([sock], [], [], 0)
        if not readable:
            return False
        return socket.socket.recv(sock, 1, socket.MSG_PEEK | socket.MSG_DONTWAIT) == b""
    except BlockingIOError:
        return False
    except (OSError, ValueError):
        return True


class _Reply:
    def __init__(self, status: int, headers: http.client.HTTPMessage, data: bytes, will_close: bool) -> None:
        self.status = status
        self.headers = headers
        self.data = data
        self.will_close = will_close

    def challenges(self) -> dict[str, str]:
        return parse_challenges(self.headers.get_all("WWW-Authenticate") or [])


class HttpSession:
    def __init__(
        self,
        target: HttpTarget,
        username: str,
        password: str,
        *,
        sign_in: str = auth.DEFAULT_SIGN_IN,
        connect_timeout: float = 15.0,
        io_timeout: float = 60.0,
        query_timeout: float = 1800.0,
        context_factory: Callable | None = None,
        ssl_context: ssl.SSLContext | None = None,
    ) -> None:
        auth.check_allowed(sign_in, target)
        self.target = target
        self.host = target.host
        self.port = target.port
        self.sign_in = sign_in
        self._username = username
        self._password = password
        self._connect_timeout = connect_timeout
        self._io_timeout = io_timeout
        self._query_timeout = query_timeout
        self._context_factory = context_factory or _default_context
        self._ssl_context = ssl_context
        self._lock = threading.Lock()
        self._conn: http.client.HTTPConnection | None = None
        self._sock: socket.socket | None = None  # kept: http.client drops its own on "Connection: close"
        self._bindings = None
        self._signed_in_connection = False  # NTLM / Kerberos signed this TCP connection in
        self._basic_asked = False  # the server asked for Basic: send it with every request
        self._aborted = False
        self._session_id: str | None = None
        self.offered: list[str] = []  # what the server's last 401 offered, as it named them
        self.sign_in_note = ""

    def __repr__(self) -> str:  # never include credentials
        return f"<HttpSession {self.target.scheme}://{self.host}:{self.port} open={self.is_open}>"

    @property
    def is_open(self) -> bool:
        return self._conn is not None and not self._aborted

    # -- lifecycle ---------------------------------------------------------------
    def open(self) -> None:
        with self._lock:
            self._close()
            self._aborted = False
            self._connect()

    def close(self) -> None:
        self.abort()
        with self._lock:
            self._close()

    def abort(self) -> None:
        """Cancel whatever request is running. Safe to call from any thread."""
        self._aborted = True
        sock = self._sock
        if sock is not None:
            try:
                # The plain socket's shutdown, also for TLS: SSLSocket.shutdown would
                # first drop its TLS state under a read running in another thread.
                socket.socket.shutdown(sock, socket.SHUT_RDWR)
            except OSError:
                pass

    def _close(self) -> None:
        if self._conn is not None:
            try:
                self._conn.close()
            except OSError:
                pass
        self._conn = None
        self._sock = None
        self._bindings = None
        self._signed_in_connection = False
        self._basic_asked = False
        self._session_id = None

    def _connect(self) -> None:
        if self.target.secure:
            conn: http.client.HTTPConnection = http.client.HTTPSConnection(
                self.host, self.port, timeout=self._connect_timeout, context=self._ssl_context or default_ssl_context()
            )
        else:
            conn = http.client.HTTPConnection(self.host, self.port, timeout=self._connect_timeout)
        conn.auto_open = 0  # never reconnect behind our back: the sign-in belongs to one connection
        try:
            conn.connect()
        except Exception as exc:
            conn.close()
            raise self._connect_error(exc) from exc
        sock = conn.sock
        try:
            sock.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
        except OSError:
            pass
        self._conn = conn
        self._sock = sock
        self._signed_in_connection = False
        self._bindings = channel_bindings(sock) if self.target.secure else None
        log.debug("connected to %s port %d", self.target.scheme, self.port)

    def _connect_error(self, exc: Exception) -> CubeError:
        if isinstance(exc, CubeError):
            return exc
        if isinstance(exc, (socket.gaierror, UnicodeError)):
            return CubeError(
                "unreachable",
                f"The server name '{self.host}' could not be found. Check the name, and that the VPN is connected.",
                str(exc),
            )
        if isinstance(exc, ssl.SSLCertVerificationError):
            text = getattr(exc, "verify_message", "") or str(exc)
            if getattr(exc, "verify_code", None) == 62 or "hostname" in text.lower() or "match" in text.lower():
                return CubeError(
                    "unreachable",
                    f"The server's certificate is not for the name '{self.host}', so nothing was sent to it. "
                    "Use the name the certificate is for (ask IT).",
                    text,
                )
            return CubeError(
                "unreachable",
                "This Mac does not trust the server's certificate, so nothing was sent to it. If the "
                "company uses its own certificates, ask IT to install the company's root certificate on this Mac.",
                text,
            )
        if isinstance(exc, ssl.SSLError):
            return CubeError("unreachable", "The secure (https) connection to the server failed.", str(exc))
        if isinstance(exc, TimeoutError):
            return CubeError(
                "unreachable",
                f"Can't reach the server on port {self.port}. Are you on the VPN or office network?",
                "connection timed out",
            )
        if isinstance(exc, OSError):
            return CubeError(
                "unreachable",
                f"Can't reach the server on port {self.port}. Are you on the VPN or office network?",
                str(exc),
            )
        return CubeError("unreachable", "The connection to the server failed.", type(exc).__name__)

    # -- requests ----------------------------------------------------------------
    def discover(
        self,
        request_type: str,
        restrictions: Mapping[str, str] | None = None,
        catalog: str | None = None,
    ) -> list[dict[str, str]]:
        def build(session_id: str | None) -> bytes:
            return envelopes.discover(request_type, restrictions, catalog, session_id)

        return self._roundtrip(build, "discover", self._io_timeout).rows

    def execute(self, mdx: str, catalog: str | None) -> CellSet:
        def build(session_id: str | None) -> bytes:
            return envelopes.execute(mdx, catalog, session_id)

        return self._roundtrip(build, "execute", self._query_timeout).cellset

    def execute_rowset(self, statement: str, catalog: str | None) -> xmla.Parsed:
        def build(session_id: str | None) -> bytes:
            return envelopes.execute(statement, catalog, session_id, tabular=True)

        return self._roundtrip(build, "execute", self._query_timeout)

    def _roundtrip(self, build: Callable[[str | None], bytes], action: str, timeout: float) -> xmla.Parsed:
        with self._lock:
            if self._conn is None or self._aborted:
                self._close()
                raise CubeError("unreachable", "The connection to the server is closed.", retry=True)
            try:
                body = _DECLARATION + build(self._session_id)
                reply = self._exchange(body, _ACTIONS[action], timeout)
                parsed = self._read(reply)
            except CubeError as exc:
                log.debug("request failed: %s (%s)", exc.category, exc.detail or exc.message)
                self._close()
                raise
            # A server fault leaves the session usable; keep the session id first.
            if parsed.session_id and not self._session_id:
                self._session_id = parsed.session_id
            if parsed.error is not None:
                raise parsed.error
            return parsed

    def _exchange(self, body: bytes, action: str, timeout: float) -> _Reply:
        conn = self._usable_connection()
        assert self._sock is not None
        self._sock.settimeout(timeout)
        authorization = self._basic_header() if self._basic_asked else None
        reply = self._send(conn, body, action, authorization)
        if reply.status == 401:
            reply = self._sign_in(body, action, reply)
        return reply

    def _usable_connection(self) -> http.client.HTTPConnection:
        conn = self._conn
        assert conn is not None
        if conn.sock is None or _dropped(conn.sock):
            # Idle connections are closed by the server (IIS after two minutes): open a
            # new one. The XMLA session lives on the server, so its id stays valid.
            log.debug("the server closed the connection; connecting again")
            conn.close()
            self._connect()
            conn = self._conn
            assert conn is not None
        return conn

    def _send(self, conn: http.client.HTTPConnection, body: bytes, action: str, authorization: str | None) -> _Reply:
        headers = {
            "Content-Type": "text/xml",
            "SOAPAction": action,
            "Accept": "text/xml, application/xml",
            "Accept-Encoding": "identity",
            "User-Agent": f"CubeExplorer/{__version__} (Mac)",
        }
        if authorization:
            headers["Authorization"] = authorization
        try:
            if conn.sock is None:
                raise ConnectionResetError("the server closed the connection")
            conn.request("POST", self.target.path, body=body, headers=headers)
            response = conn.getresponse()
            declared = response.getheader("Content-Length")
            if declared and declared.isdigit() and int(declared) > MAX_RESPONSE:
                raise ProtocolError(f"the server announced a response of more than {MAX_RESPONSE // (1024 * 1024)} MB")
            chunks, total = [], 0
            while True:
                chunk = response.read(_READ_CHUNK)
                if not chunk:
                    break
                total += len(chunk)
                if total > MAX_RESPONSE:
                    raise ProtocolError(f"the server sent more than {MAX_RESPONSE // (1024 * 1024)} MB")
                chunks.append(chunk)
            log.debug("HTTP %d, %d bytes", response.status, total)
            return _Reply(response.status, response.msg, b"".join(chunks), response.will_close)
        except CubeError:
            raise
        except Exception as exc:
            raise self._io_error(exc) from exc

    def _io_error(self, exc: Exception) -> CubeError:
        if self._aborted:
            return CubeError("cancelled", "The query was cancelled.")
        if isinstance(exc, TimeoutError):
            return CubeError("unreachable", "The server took too long to answer.", "read timed out")
        if isinstance(exc, http.client.IncompleteRead):
            return CubeError("unreachable", "The connection to the server was lost.", "response cut short", retry=True)
        if isinstance(exc, (ConnectionError, http.client.RemoteDisconnected, http.client.NotConnected)):
            return CubeError("unreachable", "The connection to the server was lost.", str(exc) or type(exc).__name__, retry=True)
        if isinstance(exc, http.client.HTTPException):
            return ProtocolError(f"not an HTTP answer ({type(exc).__name__})")
        if isinstance(exc, OSError):
            return CubeError("unreachable", "The connection to the server was lost.", str(exc), retry=True)
        return CubeError("unreachable", "The connection to the server failed.", type(exc).__name__)

    # -- the answer --------------------------------------------------------------
    def _read(self, reply: _Reply) -> xmla.Parsed:
        content_type = (reply.headers.get("Content-Type") or "").lower()
        encoding = (reply.headers.get("Content-Encoding") or "identity").strip().lower()
        if encoding not in ("", "identity"):
            raise ProtocolError(f"the server compressed its answer ({encoding}) although it was asked not to")
        if "xpress" in content_type or "application/sx" in content_type:
            raise ProtocolError(f"the server answered in binary or compressed XML ({content_type})")
        is_xml = reply.data.lstrip(b"\xef\xbb\xbf \t\r\n").startswith(b"<")
        if reply.status in (200, 500) and is_xml:
            parsed = xmla.parse(reply.data, raise_errors=False)
            if reply.status == 500 and parsed.error is None:
                raise CubeError("server", "The web server reported an error (HTTP 500).", reply.data[:300].decode("utf-8", "replace"))
            return parsed
        if reply.status == 200:
            raise ProtocolError("the answer is not XML")
        raise self._status_error(reply)

    def _status_error(self, reply: _Reply) -> CubeError:
        status = reply.status
        if 300 <= status < 400:
            location = reply.headers.get("Location") or ""
            try:
                parts = urlsplit(location)
                shown = f"{parts.scheme}://{parts.netloc}{parts.path}" if parts.netloc else parts.path
            except ValueError:
                shown = ""
            return CubeError(
                "unreachable",
                "The web server sends this address somewhere else"
                + (f" ({shown})" if shown else "")
                + ". Enter that address instead.",
                f"HTTP {status}",
            )
        if status == 401:
            return CubeError("auth", "The web server refused the sign-in (HTTP 401).", f"HTTP {status}")
        if status == 403:
            return CubeError(
                "auth",
                "The web server refused access (HTTP 403). Your account may not be allowed to use this address; ask IT.",
                f"HTTP {status}",
            )
        if status == 404:
            return CubeError(
                "unreachable",
                "There is no cube at that web address (HTTP 404). Check it; it usually ends in "
                "/olap/msmdpump.dll.",
                f"HTTP {status}",
            )
        if status == 407:
            return CubeError(
                "unreachable",
                "A proxy between this Mac and the server asks for a sign-in (HTTP 407). The add-in "
                "connects directly; ask IT whether the server can be reached without the proxy.",
                f"HTTP {status}",
            )
        if status in (502, 503, 504):
            return CubeError("unreachable", f"The web server could not reach the cube just now (HTTP {status}).", f"HTTP {status}")
        return CubeError("server", f"The web server answered with an error (HTTP {status}).", f"HTTP {status}")

    # -- sign-in -----------------------------------------------------------------
    def _basic_header(self) -> str:
        pair = f"{self._username}:{self._password}".encode("utf-8")
        return "Basic " + base64.b64encode(pair).decode("ascii")

    def _sign_in(self, body: bytes, action: str, refused: _Reply) -> _Reply:
        offers = refused.challenges()
        self.offered = [_SCHEME_NAMES.get(name, name) for name in offers]
        log.debug("the server asks for sign-in; it offers %s", ", ".join(self.offered) or "nothing")
        if self.sign_in == "basic":
            return self._basic(body, action, offers)
        if self._signed_in_connection:
            # The server wants a sign-in on every request (or forgot this connection's).
            log.debug("signing the connection in again")
            self._signed_in_connection = False
        return self._handshake(body, action, offers)

    def _not_offered(self, offers: Mapping[str, str]) -> CubeError:
        mine = auth.LABELS[self.sign_in]
        if not offers:
            return CubeError(
                "auth-mechanism",
                "The web server refused access without saying how to sign in (HTTP 401).",
                "no WWW-Authenticate header",
            )
        works = []
        if "negotiate" in offers or "ntlm" in offers:
            works.append(auth.LABELS["ntlm"])
        if "negotiate" in offers or "kerberos" in offers:
            works.append(auth.LABELS["kerberos"])
        if "basic" in offers and self.target.secure:
            works.append(auth.LABELS["basic"])
        names = ", ".join(_SCHEME_NAMES.get(name, name) for name in offers)
        if works:
            choose = " or ".join(f'"{label}"' for label in works if label != mine)
            return CubeError(
                "auth-mechanism",
                f"The web server does not accept {mine}. Choose {choose}.",
                f"the server offers: {names}",
            )
        if "basic" in offers:
            return CubeError(
                "auth-mechanism",
                "The web server accepts only Basic sign-in, which sends the password itself. This "
                "add-in does that only over https: use the https:// address.",
                f"the server offers: {names}",
            )
        return CubeError(
            "auth-mechanism",
            f"The web server asks for a kind of sign-in this add-in does not have ({names}).",
            f"the server offers: {names}",
        )

    def _basic(self, body: bytes, action: str, offers: Mapping[str, str]) -> _Reply:
        if "basic" not in offers:
            raise self._not_offered(offers)
        if self._basic_asked:
            raise CubeError("auth", "The web server refused the user name or password.", "Basic: HTTP 401")
        self._basic_asked = True
        reply = self._send(self._usable_connection(), body, action, self._basic_header())
        if reply.status == 401:
            self._basic_asked = False
            raise CubeError("auth", "The web server refused the user name or password.", "Basic: HTTP 401")
        self.sign_in_note = "Basic, over https"
        return reply

    def _handshake(self, body: bytes, action: str, offers: Mapping[str, str]) -> _Reply:
        kerberos = self.sign_in == "kerberos"
        if kerberos:
            scheme = "Negotiate" if "negotiate" in offers else "Kerberos" if "kerberos" in offers else None
        else:
            scheme = "NTLM" if "ntlm" in offers else "Negotiate" if "negotiate" in offers else None
        if scheme is None:
            raise self._not_offered(offers)
        # After a 401 with "Connection: close" this is a new connection: the handshake
        # starts there, and everything after it must stay on it.
        conn = self._usable_connection()
        try:
            context = self._context_factory(self.sign_in, self._username, self._password, self.host, self._bindings)
        except CubeError:
            raise
        except Exception as exc:
            raise CubeError("auth-mechanism", "Password sign-in could not be started on this Mac.", type(exc).__name__) from exc
        choice = getattr(context, "cube_choice", None)
        if choice is not None:
            how = "ticket already on this Mac" if choice.credential == "ticket" else "ticket got with the password"
            self.sign_in_note = f"Kerberos ({scheme}), {how} ({choice.service})"
        else:
            self.sign_in_note = f"{'Kerberos' if kerberos else 'NTLM'} ({scheme})"

        in_token: bytes | None = None
        for round_number in range(1, MAX_HANDSHAKE_ROUNDS + 1):
            try:
                out_token = context.step(in_token)
            except Exception as exc:
                # The exception text can carry the account and domain: do not pass it on.
                raise CubeError(
                    "auth",
                    "The web server did not accept the Kerberos sign-in. Ask IT whether it accepts "
                    f'Kerberos from outside Windows, or choose "{auth.LABELS["ntlm"]}".'
                    if kerberos
                    else "The sign-in could not be completed. Check the user name (DOMAIN\\user) and password.",
                    type(exc).__name__,
                ) from exc
            if not out_token:
                raise CubeError("auth", "The sign-in stopped before it was complete.")
            if conn.sock is None:
                raise CubeError(
                    "auth",
                    "The web server closed the connection in the middle of the sign-in. NTLM needs one "
                    "connection kept open; a proxy or load balancer in between may prevent that.",
                    "Connection: close during the handshake",
                )
            log.debug("sign-in round %d: sending a %d-byte token (%s)", round_number, len(out_token), scheme)
            reply = self._send(conn, body, action, f"{scheme} {base64.b64encode(out_token).decode('ascii')}")
            challenge = reply.challenges().get(scheme.lower(), "")
            if reply.status == 401:
                if not challenge:
                    raise CubeError(
                        "auth",
                        "The web server refused the Kerberos sign-in. Ask IT whether the server accepts "
                        f'Kerberos from outside Windows, or choose "{auth.LABELS["ntlm"]}".'
                        if kerberos
                        else "The web server refused the user name or password.",
                        f"{scheme}: HTTP 401",
                    )
                in_token = _decode_token(challenge)
                continue
            if challenge and not context.complete:
                # Mutual authentication: the server proves who it is in its last token.
                try:
                    context.step(_decode_token(challenge))
                except CubeError:
                    raise
                except Exception as exc:
                    raise CubeError("auth", "The web server's sign-in answer could not be verified.", type(exc).__name__) from exc
            self._signed_in_connection = True
            log.debug("sign-in complete after %d round(s); HTTP %d", round_number, reply.status)
            return reply
        raise CubeError("auth", "The sign-in did not finish.", f"more than {MAX_HANDSHAKE_ROUNDS} rounds")
