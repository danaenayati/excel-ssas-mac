"""An authenticated conversation with one Analysis Services instance.

Sequence: connect, authenticate (NTLM tokens carried in Authenticate SOAP messages,
in the clear), then sealed Discover and Execute requests.

A session is used by one request at a time (the lock), except `abort`, which another
thread calls to cancel. After any transport failure the session is closed: the NTLM
key stream is continuous, so a half-read message cannot be recovered from.

Handshake logic derived from oluies/ssas-xmla-tcp (Apache-2.0); see THIRD_PARTY_NOTICES.md.
"""

from __future__ import annotations

import base64
import threading
from collections.abc import Callable, Mapping

from ..errors import CubeError, IncompleteMessage, ProtocolError
from ..pivot.model import CellSet
from . import dime, envelopes, sealing, xmla
from .transport import MessageReader, SocketChannel, send_message

SERVICE_CLASS = "MSOLAPSvc.3"
MAX_HANDSHAKE_ROUNDS = 10


def ntlm_context(username: str, password: str, host: str):
    import spnego

    return spnego.client(
        username=username,
        password=password,
        hostname=host,
        service=SERVICE_CLASS,
        protocol="ntlm",
    )


class Session:
    def __init__(
        self,
        host: str,
        port: int,
        username: str,
        password: str,
        *,
        connect_timeout: float = 15.0,
        io_timeout: float = 60.0,
        query_timeout: float = 1800.0,
        context_factory: Callable = ntlm_context,
    ) -> None:
        self.host = host
        self.port = port
        self._username = username
        self._password = password
        self._connect_timeout = connect_timeout
        self._io_timeout = io_timeout
        self._query_timeout = query_timeout
        self._context_factory = context_factory
        self._lock = threading.Lock()
        self._channel: SocketChannel | None = None
        self._reader: MessageReader | None = None
        self._context = None
        self._session_id: str | None = None

    def __repr__(self) -> str:  # never include credentials
        return f"<Session {self.host}:{self.port} open={self.is_open}>"

    @property
    def is_open(self) -> bool:
        return self._channel is not None and self._context is not None

    # -- lifecycle ---------------------------------------------------------------
    def open(self) -> None:
        with self._lock:
            self._close()
            channel = SocketChannel(self.host, self.port, self._connect_timeout)
            channel.set_timeout(self._io_timeout)
            self._channel = channel
            self._reader = MessageReader(channel)
            try:
                self._context = self._authenticate(channel, self._reader)
            except BaseException:
                self._close()
                raise

    def close(self) -> None:
        with self._lock:
            self._close()

    def abort(self) -> None:
        """Cancel whatever request is running. Safe to call from any thread."""
        channel = self._channel
        if channel is not None:
            channel.abort()

    def _close(self) -> None:
        if self._channel is not None:
            self._channel.close()
        self._channel = None
        self._reader = None
        self._context = None
        self._session_id = None

    # -- requests ----------------------------------------------------------------
    def discover(
        self,
        request_type: str,
        restrictions: Mapping[str, str] | None = None,
        catalog: str | None = None,
    ) -> list[dict[str, str]]:
        def build(session_id: str | None) -> bytes:
            return envelopes.discover(request_type, restrictions, catalog, session_id)

        return self._roundtrip(build, self._io_timeout).rows

    def execute(self, mdx: str, catalog: str | None) -> CellSet:
        def build(session_id: str | None) -> bytes:
            return envelopes.execute(mdx, catalog, session_id)

        return self._roundtrip(build, self._query_timeout).cellset

    def _roundtrip(self, build: Callable[[str | None], bytes], timeout: float) -> xmla.Parsed:
        with self._lock:
            if self._channel is None or self._reader is None or self._context is None:
                raise CubeError("unreachable", "The connection to the server is closed.")
            try:
                self._channel.set_timeout(timeout)
                payload = build(self._session_id)
                send_message(self._channel, sealing.seal_message(self._context, payload), dime.OPTIONS_NEGOTIATED)
                raw = self._reader.read_message()
                try:
                    plain = sealing.unseal_message(self._context, raw)
                except IncompleteMessage as exc:
                    raise ProtocolError(f"sealed response is truncated: {exc}") from exc
                except CubeError:
                    raise
                except Exception as exc:
                    raise ProtocolError(f"could not decrypt the response ({type(exc).__name__})") from exc
                if not plain.lstrip(b"\xef\xbb\xbf \t\r\n").startswith(b"<"):
                    raise ProtocolError("the decrypted response is not XML")
            except CubeError:
                # Transport-level trouble: the stream position and cipher state are
                # no longer trustworthy.
                self._close()
                raise
            parsed = xmla.parse(plain)  # a server fault leaves the session usable
            if parsed.session_id and not self._session_id:
                self._session_id = parsed.session_id
            return parsed

    # -- authentication ----------------------------------------------------------
    def _authenticate(self, channel: SocketChannel, reader: MessageReader):
        try:
            context = self._context_factory(self._username, self._password, self.host)
        except Exception as exc:
            raise CubeError(
                "auth-mechanism",
                "Password sign-in could not be started on this Mac.",
                type(exc).__name__,
            ) from exc
        for name in ("wrap_winrm", "unwrap_winrm"):
            if not callable(getattr(context, name, None)):
                raise CubeError("auth-mechanism", "The sign-in method cannot protect messages.", name)

        in_token: bytes | None = None
        sent_any = False
        for _ in range(MAX_HANDSHAKE_ROUNDS):
            try:
                out_token = context.step(in_token)
            except Exception as exc:
                # The exception text can carry the account and domain: do not pass it on.
                raise CubeError(
                    "auth",
                    "The sign-in could not be completed. Check the username (DOMAIN\\user) and password.",
                    type(exc).__name__,
                ) from exc
            if not out_token:
                if context.complete:
                    return context
                raise CubeError("auth", "The sign-in stopped before it was complete.")
            # NEGO stays clear on the very first record and is set on every later one.
            options = dime.OPTIONS_NEGOTIATED if sent_any else dime.OPTIONS_FIRST
            send_message(
                channel,
                sealing.BOM + envelopes.authenticate(base64.b64encode(out_token).decode("ascii")),
                options,
            )
            try:
                reply = reader.read_message()
            except CubeError as exc:
                if exc.category == "unreachable" and sent_any:
                    raise CubeError(
                        "auth",
                        "The server ended the sign-in. This usually means a wrong username or "
                        "password, or a server that does not accept password (NTLM) sign-in.",
                        exc.detail,
                    ) from exc
                raise
            sent_any = True
            # Every reply is fault-checked, including the last one: for NTLM the client
            # side is complete as soon as it sends its final token, so a "Logon failure"
            # in the reply would otherwise be missed.
            parsed = xmla.parse(reply, during_auth=True)
            if context.complete:
                return context
            in_token = parsed.handshake_token or None
        raise CubeError("auth", "The sign-in did not finish.", f"more than {MAX_HANDSHAKE_ROUNDS} rounds")
