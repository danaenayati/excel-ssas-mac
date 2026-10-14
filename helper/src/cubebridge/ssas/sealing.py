"""The post-authentication frame: how a sealed message is laid out on the wire.

Every message after the handshake is sealed with the negotiated security context.
[MS-SSAS] does not document this layer. The layout below is the one a live SQL Server
2022 instance accepts, as established by the ssas-xmla-tcp project:

    uint16  dataSize    little-endian, ciphertext length
    uint16  tokenSize   little-endian, security token length
    bytes   ciphertext
    bytes   token

Ciphertext comes first and the token second, the inverse of the usual GSS-API order.
Getting it backwards is silently fatal: the server drops the connection.

The two parts are what SSPI's EncryptMessage produces as its DATA and TOKEN buffers,
which is what pyspnego's `wrap_winrm` returns for any mechanism: for NTLM a 16-byte
signature, for Kerberos (RFC 4121, AES) the wrap token's header with its trailer folded
in (60 bytes for AES-256). The token's size is read from each frame, never assumed.

The UTF-8 BOM is sealed as its own frame before the body, and the body is chunked at
2888 bytes (NTLM's maximum token size, which the reference client reuses as its data
chunk size; Kerberos takes the same frames). The key stream and sequence numbers are
continuous, so frames must be sealed and unsealed in wire order.

Confirmed against a live server for NTLM only. A mechanism that pads its plaintext
(Kerberos with the old RC4 cipher) cannot use this frame, because neither length field
carries the unpadded size.

Derived from oluies/ssas-xmla-tcp (Apache-2.0); see THIRD_PARTY_NOTICES.md.
"""

from __future__ import annotations

import struct
from typing import Protocol

from ..errors import IncompleteMessage, ProtocolError

BOM = b"\xef\xbb\xbf"
_HEADER = struct.Struct("<HH")
MAX_CHUNK = 2888


class SealingContext(Protocol):
    def wrap_winrm(self, data: bytes) -> tuple[bytes, bytes, int]: ...

    def unwrap_winrm(self, header: bytes, data: bytes) -> bytes: ...


class PaddingNotSupported(ProtocolError):
    """The negotiated sign-in pads what it encrypts, which the SSAS frame cannot carry."""

    def __init__(self) -> None:
        super().__init__(
            "the sign-in mechanism pads its data (Kerberos with RC4), and the SSAS frame has no "
            "field for the unpadded length"
        )
        self.category = "auth-mechanism"
        self.message = (
            "The server's Kerberos uses the old RC4 cipher, which this add-in cannot use to protect "
            'messages to Analysis Services. Ask IT to allow AES for the server, or choose "User name '
            'and password (NTLM)".'
        )
        self.args = (self.message,)


def seal_frame(context: SealingContext, payload: bytes) -> bytes:
    # (header, data, padding length): pyspnego's WinRMWrapResult, or a plain tuple.
    token, ciphertext, padding_length = context.wrap_winrm(payload)
    if padding_length:
        raise PaddingNotSupported()
    if len(ciphertext) > 0xFFFF or len(token) > 0xFFFF:
        raise ProtocolError("sealed frame exceeds the 16-bit size fields")
    return _HEADER.pack(len(ciphertext), len(token)) + ciphertext + token


def seal_message(context: SealingContext, payload: bytes) -> bytes:
    out = [seal_frame(context, BOM)]
    for start in range(0, len(payload), MAX_CHUNK):
        out.append(seal_frame(context, payload[start : start + MAX_CHUNK]))
    return b"".join(out)


def unseal_message(context: SealingContext, blob: bytes) -> bytes:
    """Decrypt every frame of a sealed message, in order, and join the plaintext."""
    view = memoryview(blob)
    out: list[bytes] = []
    offset = 0
    total = len(blob)
    while offset < total:
        if offset + _HEADER.size > total:
            raise IncompleteMessage("sealed message ends inside a frame header")
        data_size, token_size = _HEADER.unpack_from(view, offset)
        offset += _HEADER.size
        end_data = offset + data_size
        end_token = end_data + token_size
        if end_token > total:
            raise IncompleteMessage("sealed frame runs past the end of the message")
        out.append(context.unwrap_winrm(bytes(view[end_data:end_token]), bytes(view[offset:end_data])))
        offset = end_token
    plain = b"".join(out)
    return plain[len(BOM) :] if plain.startswith(BOM) else plain
