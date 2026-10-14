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
_XSD_NS = "{http://www.w3.org/2001/XMLSchema}"
_SQL_FIELD = "{urn:schemas-microsoft-com:xml-sql}field"
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
    handshake_text: str | None = None
    error: CubeError | None = None
    # Rowset columns in schema order, as their decoded names, and their XSD types.
    columns: list[str] = field(default_factory=list)
    column_types: dict[str, str] = field(default_factory=dict)

    @property
    def handshake_token(self) -> bytes | None:
        """The server's sign-in token, decoded only when asked for.

        The reply to our last token carries nothing we need, so it is never decoded;
        only a token the handshake is about to use can fail here.
        """
        return None if self.handshake_text is None else decode_token(self.handshake_text)


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
                parsed.handshake_text = element.text or ""
            elif element.tag == _XSD_NS + "element" and (element.get("type") or "").startswith("xsd:"):
                # A column of a typed rowset (DRILLTHROUGH). sql:field holds the readable
                # name; the element name is its XML-encoded form, used in the rows.
                name = _decode_name(element.get("name", ""))
                if name:
                    parsed.columns.append(name)
                    parsed.column_types[name] = element.get("type", "")
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


def typed_rows(parsed: Parsed) -> list[list[object]]:
    """Rows in column order, converted by the schema's types; missing cells are None.

    Without a schema, columns are taken in first-seen order and left as text.
    """
    columns = list(parsed.columns)
    if not columns:
        for row in parsed.rows:
            for name in row:
                if name not in columns:
                    columns.append(name)
    out = []
    for row in parsed.rows:
        line: list[object] = []
        for name in columns:
            text = row.get(name)
            line.append(None if text is None else _convert(text, parsed.column_types.get(name, "")))
        out.append(line)
    return out


def _convert(text: str, xsd_type: str) -> object:
    kind = xsd_type.rsplit(":", 1)[-1]
    try:
        if kind in _INT_TYPES:
            return int(text)
        if kind in _FLOAT_TYPES:
            return float(text)
    except ValueError:
        return text
    if kind == "boolean":
        return text.strip().lower() in ("true", "1")
    return text


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


def decode_token(text: str) -> bytes:
    """Decode a base64 sign-in token as a real server sends it.

    Windows' own encoder breaks base64 into lines (CR LF every 64 characters and at
    the end), so white space inside the token is normal and is removed first. What
    is left must be strict base64. If it is not, the error describes the characters
    that do not belong, never the token itself.
    """
    compact = "".join(text.split())
    if not compact:
        return b""
    try:
        return base64.b64decode(compact + "=" * (-len(compact) % 4), validate=True)
    except (binascii.Error, ValueError) as exc:
        strangers: dict[str, int] = {}
        for ch in compact:
            if not (ch.isascii() and (ch.isalnum() or ch in "+/=")):
                code = f"U+{ord(ch):04X}"
                strangers[code] = strangers.get(code, 0) + 1
        found = ", ".join(f"{code} x{count}" for code, count in sorted(strangers.items())[:12]) or "none (bad length or padding)"
        raise ProtocolError(
            f"the server's sign-in token is not valid base64: {len(compact)} characters, "
            f"not in the base64 alphabet: {found}"
        ) from exc


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
                    # Excel shows a member with an empty caption as "(blank)".
                    caption=fields.get("Caption", "") or "(blank)",
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
