"""The post-authentication frame: how a sealed message is laid out on the wire.

Every message after the handshake is sealed with the negotiated security context.
[MS-SSAS] does not document this layer. The layout below is the one a live SQL Server
2022 instance accepts, as established by the ssas-xmla-tcp project:

    uint16  dataSize    little-endian, ciphertext length
    uint16  tokenSize   little-endian, security token length (16 for NTLM)
    bytes   ciphertext
    bytes   token

Ciphertext comes first and the token second, the inverse of the usual GSS-API order.
Getting it backwards is silently fatal: the server drops the connection.

The UTF-8 BOM is sealed as its own frame before the body, and the body is chunked at
2888 bytes (NTLM's maximum token size, which the reference client reuses as its data
chunk size). The NTLM key stream is continuous, so frames must be sealed and unsealed
in wire order.

Confirmed for NTLM only. A mechanism that pads its plaintext cannot use this frame,
because neither length field carries the unpadded size.

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


def seal_frame(context: SealingContext, payload: bytes) -> bytes:
    token, ciphertext, padding_length = context.wrap_winrm(payload)
    if padding_length:
        raise ProtocolError(
            "the sign-in mechanism pads its data, and the SSAS frame has no field for "
            "the unpadded length; only NTLM is supported"
        )
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
