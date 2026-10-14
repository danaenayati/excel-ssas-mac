"""A loopback stand-in for an Analysis Services instance.

It speaks the same wire layers a real server does (DIME records, an NTLM acceptor,
sealed frames after the handshake) but answers from canned XML. The framing here is
written out independently with `struct` on purpose: if it imported the client's own
dime/sealing modules, a mistake in those would cancel itself out.
"""

from __future__ import annotations

import base64
import re
import socket
import struct
import threading
import time

import spnego
import xml_samples as x

BOM = b"\xef\xbb\xbf"


def _pad(n: int) -> int:
    return (-n) % 4


class _Wire:
    def __init__(self, conn: socket.socket) -> None:
        self.conn = conn
        self.buf = b""

    def _exact(self, n: int) -> bytes:
        while len(self.buf) < n:
            chunk = self.conn.recv(65536)
            if not chunk:
                raise ConnectionError("client went away")
            self.buf += chunk
        out, self.buf = self.buf[:n], self.buf[n:]
        return out

    def read_message(self) -> tuple[bytes, int]:
        """Return (payload, first OPTIONS byte of the first record)."""
        data = b""
        first_options = None
        while True:
            flags, _type_t, opt_len, id_len, type_len, data_len = struct.unpack(">BBHHHI", self._exact(12))
            assert flags >> 3 == 1, "DIME version must be 1"
            options = self._exact(opt_len + _pad(opt_len))
            self._exact(id_len + _pad(id_len))
            self._exact(type_len + _pad(type_len))
            data += self._exact(data_len + _pad(data_len))[:data_len]
            if first_options is None:
                first_options = options[0] if options else 0
            if flags & 0x02:
                return data, first_options

    def send_message(self, payload: bytes, record_size: int) -> None:
        parts = [payload[i : i + record_size] for i in range(0, len(payload), record_size)] or [b""]
        out = bytearray()
        for index, part in enumerate(parts):
            first, last = index == 0, index == len(parts) - 1
            flags = (1 << 3) | (int(first) << 2) | (int(last) << 1) | int(not last)
            options = bytes([1, 0, 0, 0]) if first else b""
            type_ = b"text/xml" if first else b""
            out += struct.pack(">BBHHHI", flags, (1 if first else 0) << 4, len(options), 0, len(type_), len(part))
            out += options + type_ + part + b"\x00" * _pad(len(part))
        self.conn.sendall(bytes(out))


def _seal(ctx, payload: bytes, frame_size: int) -> bytes:
    out = bytearray()
    for piece in [BOM] + [payload[i : i + frame_size] for i in range(0, len(payload), frame_size)]:
        wrapped = ctx.wrap_winrm(piece)
        out += struct.pack("<HH", len(wrapped.data), len(wrapped.header)) + wrapped.data + wrapped.header
    return bytes(out)


def _unseal(ctx, blob: bytes, token_sizes: list[int] | None = None) -> bytes:
    out, offset = b"", 0
    while offset < len(blob):
        data_size, token_size = struct.unpack_from("<HH", blob, offset)
        offset += 4
        data = blob[offset : offset + data_size]
        token = blob[offset + data_size : offset + data_size + token_size]
        offset += data_size + token_size
        if token_sizes is not None:
            token_sizes.append(token_size)
        out += ctx.unwrap_winrm(token, data)
    return out[3:] if out.startswith(BOM) else out


class FakeSsas:
    """Start with `with FakeSsas(...) as server:` and connect to `server.port`."""

    def __init__(self, responses=None, record_size: int = 60000, frame_size: int = 16000, acceptor=None) -> None:
        # substring of the request -> bytes, or callable(request_text) -> bytes
        self.responses = responses or {}
        self.record_size = record_size
        self.frame_size = frame_size
        # Makes the server's security context for each connection; NTLM by default.
        self.acceptor = acceptor or (lambda: spnego.server(protocol="ntlm"))
        self.contexts: list = []
        self.client_token_sizes: list[int] = []  # security token size of every frame received
        self.delay = 0.0
        self.wrap_tokens = False  # CRLF every 64 characters, as Windows' encoder emits
        self.final_token_text = ""  # text placed in the last reply's SspiHandshake
        self.drop_after: int | None = None  # close the connection after this many sealed requests
        self.requests: list[str] = []
        self.handshake_options: list[int] = []
        self.sealed_options: list[int] = []
        self.connections = 0
        self._sock = socket.socket()
        self._sock.bind(("127.0.0.1", 0))
        self._sock.listen(8)
        self.port = self._sock.getsockname()[1]
        self._stop = False
        self._thread = threading.Thread(target=self._accept, daemon=True)

    def __enter__(self) -> FakeSsas:
        self._thread.start()
        return self

    def __exit__(self, *exc) -> None:
        self._stop = True
        try:
            self._sock.close()
        except OSError:
            pass

    def _accept(self) -> None:
        while not self._stop:
            try:
                conn, _ = self._sock.accept()
            except OSError:
                return
            self.connections += 1
            threading.Thread(target=self._serve, args=(conn,), daemon=True).start()

    def _serve(self, conn: socket.socket) -> None:
        wire = _Wire(conn)
        try:
            ctx = self.acceptor()
            self.contexts.append(ctx)
            while not ctx.complete:
                payload, options = wire.read_message()
                self.handshake_options.append(options)
                text = payload.decode("utf-8-sig")
                assert "analysisservices/2003/ext" in text, "Authenticate must use the ext namespace"
                token = base64.b64decode(re.search(r"<SspiHandshake>(.*?)</SspiHandshake>", text, re.S).group(1))
                try:
                    out = ctx.step(token)
                except Exception:
                    wire.send_message(x.soap_fault("Logon failure: unknown user name or bad password."), self.record_size)
                    return
                text = base64.b64encode(out).decode() if out else self.final_token_text
                if out and self.wrap_tokens:
                    text = "".join(text[i : i + 64] + "\r\n" for i in range(0, len(text), 64))
                wire.send_message(x.authenticate_response(text), self.record_size)
            served = 0
            while True:
                payload, options = wire.read_message()
                self.sealed_options.append(options)
                text = _unseal(ctx, payload, self.client_token_sizes).decode("utf-8")
                self.requests.append(text)
                if self.delay:
                    time.sleep(self.delay)
                if self.drop_after is not None and served >= self.drop_after:
                    return
                served += 1
                wire.send_message(_seal(ctx, self._respond(text), self.frame_size), self.record_size)
        except (ConnectionError, OSError):
            pass
        finally:
            try:
                conn.close()
            except OSError:
                pass

    def _respond(self, text: str) -> bytes:
        for marker, response in self.responses.items():
            if marker in text:
                return response(text) if callable(response) else response
        return x.soap_fault(f"The fake server has no canned response for this request: {text[:200]}")
