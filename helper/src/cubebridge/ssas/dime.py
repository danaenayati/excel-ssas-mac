"""DIME record framing for the Analysis Services TCP binding.

Every field is specified in [MS-SSAS] "TCP":
https://learn.microsoft.com/en-us/openspecs/sql_server_protocols/ms-ssas/f172a52f-f69e-4051-8b3a-627433e978fb

Record header, 12 bytes, big-endian:

    byte 0       VERSION (5 bits, MUST be 1) | MB | ME | CF
    byte 1       TYPE_T (4 bits) | RESERVED (4 bits, MUST be 0)
    bytes 2-3    OPTIONS_LENGTH
    bytes 4-5    ID_LENGTH
    bytes 6-7    TYPE_LENGTH
    bytes 8-11   DATA_LENGTH

then OPTIONS, ID, TYPE and DATA, each padded to a 4-byte boundary. The length fields
exclude their own padding.

Binary XML and XPRESS compression are optional and negotiated through the OPTIONS
bits. This client asks for neither and refuses a server that selects them.

Derived from oluies/ssas-xmla-tcp (Apache-2.0); see THIRD_PARTY_NOTICES.md.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass

from ..errors import ProtocolError

VERSION = 1
HEADER_LEN = 12
_HEADER = struct.Struct(">BBHHHI")

TYPE_TEXT_XML = b"text/xml"

# OPTIONS: 4 bytes, only the first is used. Bits from least significant upward.
OPT_NEGO = 0x01
OPT_REQ_SX = 0x02
OPT_REQ_XPRESS = 0x04
OPT_RESP_SX = 0x08
OPT_RESP_XPRESS = 0x10

# NEGO is clear on the very first record of a connection and set on every later one.
# RESP_XPRESS must stay clear: setting it makes the server answer with compressed XML.
OPTIONS_FIRST = bytes(4)
OPTIONS_NEGOTIATED = bytes([OPT_NEGO, 0, 0, 0])


def pad(n: int) -> int:
    """Bytes of padding needed to reach the next 4-byte boundary."""
    return (-n) % 4


@dataclass(frozen=True)
class Header:
    mb: bool
    me: bool
    cf: bool
    type_t: int
    options_len: int
    id_len: int
    type_len: int
    data_len: int


def encode_record(
    data: bytes,
    *,
    options: bytes = OPTIONS_FIRST,
    type_: bytes = TYPE_TEXT_XML,
    mb: bool = True,
    me: bool = True,
    cf: bool = False,
    type_t: int = 1,
) -> bytes:
    flags = (VERSION << 3) | (int(mb) << 2) | (int(me) << 1) | int(cf)
    header = _HEADER.pack(flags, (type_t & 0x0F) << 4, len(options), 0, len(type_), len(data))
    out = [header]
    for field in (options, b"", type_, data):
        out.append(field)
        out.append(b"\x00" * pad(len(field)))
    return b"".join(out)


def decode_header(buf: bytes) -> Header:
    flags, type_byte, options_len, id_len, type_len, data_len = _HEADER.unpack(buf)
    version = flags >> 3
    if version != VERSION:
        raise ProtocolError(f"unsupported DIME version {version}; expected {VERSION}")
    return Header(
        mb=bool(flags & 0x04),
        me=bool(flags & 0x02),
        cf=bool(flags & 0x01),
        type_t=type_byte >> 4,
        options_len=options_len,
        id_len=id_len,
        type_len=type_len,
        data_len=data_len,
    )


def check_negotiated(options: bytes, type_: bytes) -> None:
    """Fail loudly if the server selected an encoding this client does not implement.

    Compressed or binary XML would otherwise arrive as convincing-looking garbage.
    """
    first = options[0] if options else 0
    if first & (OPT_REQ_SX | OPT_RESP_SX):
        raise ProtocolError("the server selected binary XML, which this add-in does not support")
    if first & (OPT_REQ_XPRESS | OPT_RESP_XPRESS):
        raise ProtocolError("the server selected compressed XML, which this add-in does not support")
    if type_ and type_ != TYPE_TEXT_XML:
        raise ProtocolError(f"the server replied with content type {type_!r}; expected text/xml")
