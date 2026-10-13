"""The TCP socket and whole-message reads and writes over it.

`Channel` is the seam: the real one wraps a socket, tests supply scripted bytes.
"""

from __future__ import annotations

import socket
from typing import Protocol

from ..errors import CubeError, ProtocolError
from . import dime

MAX_MESSAGE = 256 * 1024 * 1024


class Channel(Protocol):
    def send(self, data: bytes) -> None: ...

    def recv(self, size: int) -> bytes: ...

    def close(self) -> None: ...


class SocketChannel:
    def __init__(self, host: str, port: int, connect_timeout: float) -> None:
        self._aborted = False
        try:
            self._sock = socket.create_connection((host, port), timeout=connect_timeout)
        except socket.gaierror as exc:
            raise CubeError(
                "unreachable",
                f"The server name '{host}' could not be found. Check the name, and that the VPN is connected.",
                str(exc),
            ) from exc
        except TimeoutError as exc:
            raise CubeError(
                "unreachable",
                f"Can't reach the server on port {port}. Are you on the VPN or office network?",
                "connection timed out",
            ) from exc
        except OSError as exc:
            raise CubeError(
                "unreachable",
                f"Can't reach the server on port {port}. Are you on the VPN or office network?",
                str(exc),
            ) from exc
        self._sock.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)

    def set_timeout(self, seconds: float) -> None:
        self._sock.settimeout(seconds)

    def send(self, data: bytes) -> None:
        try:
            self._sock.sendall(data)
        except OSError as exc:
            raise self._io_error("The connection to the server was lost while sending.", exc) from exc

    def recv(self, size: int) -> bytes:
        try:
            return self._sock.recv(size)
        except TimeoutError as exc:
            raise CubeError("unreachable", "The server took too long to answer.", "read timed out") from exc
        except OSError as exc:
            raise self._io_error("The connection to the server was lost.", exc) from exc

    def abort(self) -> None:
        """Unblock a recv() running in another thread. Used to cancel a query."""
        self._aborted = True
        try:
            self._sock.shutdown(socket.SHUT_RDWR)
        except OSError:
            pass

    def close(self) -> None:
        try:
            self._sock.close()
        except OSError:
            pass

    @property
    def aborted(self) -> bool:
        return self._aborted

    def _io_error(self, message: str, exc: OSError) -> CubeError:
        if self._aborted:
            return CubeError("cancelled", "The query was cancelled.")
        return CubeError("unreachable", message, str(exc))


def send_message(channel: Channel, payload: bytes, options: bytes) -> None:
    channel.send(dime.encode_record(payload, options=options))


class MessageReader:
    """Reads whole DIME messages, driven by the lengths each record header declares.

    Bytes left over after a message stay buffered for the next one, since a peer may
    pack several messages into one segment.
    """

    def __init__(self, channel: Channel, max_message: int = MAX_MESSAGE) -> None:
        self._channel = channel
        self._buf = bytearray()
        self._pos = 0
        self._max = max_message

    def read_message(self) -> bytes:
        chunks: list[bytes] = []
        total = 0
        first = True
        while True:
            header = dime.decode_header(self._read_exact(dime.HEADER_LEN))
            total += header.data_len
            if total > self._max:
                raise ProtocolError(
                    f"the server announced a response of more than {self._max // (1024 * 1024)} MB"
                )
            options = self._read_padded(header.options_len)
            self._read_padded(header.id_len)
            type_ = self._read_padded(header.type_len)
            if first:
                if not header.mb:
                    raise ProtocolError("the first DIME record of a message does not set MB")
                dime.check_negotiated(options, type_)
                first = False
            chunks.append(self._read_padded(header.data_len))
            if header.me:
                return chunks[0] if len(chunks) == 1 else b"".join(chunks)

    def _read_padded(self, length: int) -> bytes:
        data = self._read_exact(length)
        padding = dime.pad(length)
        if padding:
            self._read_exact(padding)
        return data

    def _read_exact(self, n: int) -> bytes:
        while len(self._buf) - self._pos < n:
            chunk = self._channel.recv(max(65536, n - (len(self._buf) - self._pos)))
            if not chunk:
                if getattr(self._channel, "aborted", False):
                    raise CubeError("cancelled", "The query was cancelled.")
                raise CubeError(
                    "unreachable",
                    "The server closed the connection.",
                    "connection closed before a complete message arrived",
                )
            self._buf += chunk
        out = bytes(self._buf[self._pos : self._pos + n])
        self._pos += n
        if self._pos == len(self._buf):
            self._buf.clear()
            self._pos = 0
        elif self._pos > 4 * 1024 * 1024:
            del self._buf[: self._pos]
            self._pos = 0
        return out
