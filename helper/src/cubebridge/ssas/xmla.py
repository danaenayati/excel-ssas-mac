"""Parse XMLA responses: faults, rowsets, cellsets and handshake tokens.

One streaming pass over the response. Rows and cells are detached from the tree as
they are read, so a large result does not keep a second copy alive as XML elements.
"""

from __future__ import annotations

import base64
import binascii
import io
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

from ..errors import CubeError, ProtocolError
from ..pivot.model import AxisMember, Cell, CellSet

_XSI_TYPE = "{http://www.w3.org/2001/XMLSchema-instance}type"
_ENCODED_CHAR = re.compile(r"_x([0-9A-Fa-f]{4})_")

_INT_TYPES = frozenset(
    {"int", "long", "short", "byte", "integer", "unsignedInt", "unsignedLong", "unsignedShort", "unsignedByte"}
)
_FLOAT_TYPES = frozenset({"double", "float", "decimal"})

_DENIED_MARKERS = ("does not have access", "permission", "not authorized", "access is denied")

CELL_ERROR = "#VALUE!"


@dataclass
class Parsed:
    rows: list[dict[str, str]] = field(default_factory=list)
    cellset: CellSet = field(default_factory=CellSet)
    session_id: str | None = None
    handshake_token: bytes | None = None
    error: CubeError | None = None


def parse(data: bytes, *, during_auth: bool = False, raise_errors: bool = True) -> Parsed:
    """Parse one response.

    A server-reported error is raised, or with `raise_errors=False` returned as
    `Parsed.error` so the caller can first read what else the response carried.
    """
    parsed = Parsed()
    axes: dict[str, list[tuple[AxisMember, ...]]] = {}
    errors: list[str] = []
    stack: list[ET.Element] = []
    try:
        for event, element in ET.iterparse(io.BytesIO(data), events=("start", "end")):
            if event == "start":
                stack.append(element)
                continue
            stack.pop()
            parent = stack[-1] if stack else None
            tag = _local(element.tag)
            parent_tag = _local(parent.tag) if parent is not None else ""

            if tag == "row" and parent_tag == "root":
                parsed.rows.append({_decode_name(_local(c.tag)): (c.text or "") for c in element})
                parent.remove(element)
            elif tag == "Cell" and parent_tag == "CellData":
                ordinal = element.get("CellOrdinal")
                if ordinal is not None:
                    if not ordinal.isdigit():
                        raise ProtocolError(f"cell ordinal {ordinal[:20]!r} is not a number")
                    parsed.cellset.cells[int(ordinal)] = _cell(element)
                parent.remove(element)
            elif tag == "Axis" and parent_tag == "Axes":
                axes[element.get("name", "")] = _axis(element)
                element.clear()
            elif tag == "faultstring":
                errors.append(element.text or "")
            elif tag == "Error" and parent_tag == "Messages":
                errors.append(element.get("Description") or "")
            elif tag == "Session" and element.get("SessionId"):
                parsed.session_id = element.get("SessionId")
            elif tag == "SspiHandshake":
                parsed.handshake_token = _token(element.text)
    except ET.ParseError as exc:
        raise ProtocolError(f"the response is not valid XML ({exc})") from exc

    if errors:
        parsed.error = _server_error(next((e for e in errors if e), "The server reported an error."), during_auth)
        if raise_errors:
            raise parsed.error
        return parsed

    index = 0
    while f"Axis{index}" in axes:
        parsed.cellset.axes.append(axes[f"Axis{index}"])
        index += 1
    return parsed


def _server_error(text: str, during_auth: bool) -> CubeError:
    detail = text.strip()[:1000]
    lowered = detail.lower()
    if during_auth:
        return CubeError(
            "auth",
            "The server refused the sign-in. Check the username (DOMAIN\\user) and password.",
            detail,
        )
    if any(marker in lowered for marker in _DENIED_MARKERS):
        return CubeError("server", "Your account does not have access to this database or cube.", detail)
    return CubeError("server", "The cube rejected the request.", detail)


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _decode_name(name: str) -> str:
    return _ENCODED_CHAR.sub(lambda m: chr(int(m.group(1), 16)), name)


def _token(text: str | None) -> bytes:
    if not text or not text.strip():
        return b""
    try:
        return base64.b64decode(text.strip(), validate=True)
    except (binascii.Error, ValueError) as exc:
        raise ProtocolError("the server's sign-in token is not valid base64") from exc


def _axis(element: ET.Element) -> list[tuple[AxisMember, ...]]:
    tuples = []
    for tuple_el in element.iter():
        if _local(tuple_el.tag) != "Tuple":
            continue
        members = []
        for member_el in tuple_el:
            if _local(member_el.tag) != "Member":
                continue
            fields = {_local(child.tag): (child.text or "") for child in member_el}
            members.append(
                AxisMember(
                    unique_name=fields.get("UName", ""),
                    caption=fields.get("Caption", ""),
                    hierarchy=member_el.get("Hierarchy", ""),
                    level=_int(fields.get("LNum")),
                    # Low 16 bits: estimated child count. Higher bits are display flags.
                    child_count=_int(fields.get("DisplayInfo")) & 0xFFFF,
                )
            )
        tuples.append(tuple(members))
    return tuples


def _int(text: str | None) -> int:
    try:
        return int(text or 0)
    except ValueError:
        return 0


def _cell(element: ET.Element) -> Cell:
    value: object = None
    format_string = ""
    formatted = ""
    for child in element:
        tag = _local(child.tag)
        if tag == "Value":
            value = _value(child)
        elif tag == "FormatString":
            format_string = child.text or ""
        elif tag == "FmtValue":
            formatted = child.text or ""
    return Cell(value=value, format_string=format_string, formatted=formatted)


def _value(element: ET.Element) -> object:
    if len(element):  # <Value><Error>…</Error></Value>: the cell failed to calculate
        return CELL_ERROR
    text = element.text
    if text is None:
        return None
    xsi_type = (element.get(_XSI_TYPE) or "").rsplit(":", 1)[-1]
    try:
        if xsi_type in _INT_TYPES:
            return int(text)
        if xsi_type in _FLOAT_TYPES:
            return float(text)
    except ValueError:
        return text
    if xsi_type == "boolean":
        return text.strip().lower() in ("true", "1")
    return text
