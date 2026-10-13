"""SOAP envelopes for the three operations this client sends.

Read-only by construction: there is no builder for Create, Alter, Delete, Process or
Cancel, and no parameter through which a request could become one. `execute` wraps a
statement, and the only statements the helper produces are SELECTs from pivot/mdx.py.

Derived from oluies/ssas-xmla-tcp (Apache-2.0); see THIRD_PARTY_NOTICES.md.
"""

from __future__ import annotations

import re
from collections.abc import Mapping
from xml.sax.saxutils import escape

_SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/"
_XMLA_NS = "urn:schemas-microsoft-com:xml-analysis"
# Authenticate is NOT in the XMLA namespace. A live server rejects it there with
# "The Authenticate element ... cannot appear under Envelope/Body".
_EXT_NS = "http://schemas.microsoft.com/analysisservices/2003/ext"

# Restriction and property names become element names, which cannot be made safe by
# escaping: they are validated instead.
_ELEMENT_NAME = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def authenticate(token_b64: str) -> bytes:
    return (
        f'<Envelope xmlns="{_SOAP_NS}"><Body>'
        f'<Authenticate xmlns="{_EXT_NS}"><SspiHandshake>{escape(token_b64)}</SspiHandshake></Authenticate>'
        f"</Body></Envelope>"
    ).encode("utf-8")


def discover(
    request_type: str,
    restrictions: Mapping[str, str] | None = None,
    catalog: str | None = None,
    session_id: str | None = None,
) -> bytes:
    properties = {"Catalog": catalog} if catalog else {}
    return (
        f'<Envelope xmlns="{_SOAP_NS}">{_session_header(session_id)}<Body>'
        f'<Discover xmlns="{_XMLA_NS}"><RequestType>{escape(request_type)}</RequestType>'
        f"<Restrictions><RestrictionList>{_elements(restrictions)}</RestrictionList></Restrictions>"
        f"<Properties><PropertyList>{_elements(properties)}</PropertyList></Properties>"
        f"</Discover></Body></Envelope>"
    ).encode("utf-8")


def execute(statement: str, catalog: str | None = None, session_id: str | None = None) -> bytes:
    properties = {}
    if catalog:
        properties["Catalog"] = catalog
    # A cellset with one Tuple per axis position, without the XSD schema block.
    properties.update({"Format": "Multidimensional", "AxisFormat": "TupleFormat", "Content": "Data"})
    return (
        f'<Envelope xmlns="{_SOAP_NS}">{_session_header(session_id)}<Body>'
        f'<Execute xmlns="{_XMLA_NS}"><Command><Statement>{escape(statement)}</Statement></Command>'
        f"<Properties><PropertyList>{_elements(properties)}</PropertyList></Properties>"
        f"</Execute></Body></Envelope>"
    ).encode("utf-8")


def _session_header(session_id: str | None) -> str:
    # [MS-SSAS] "Initialization for Non-HTTP Transport": the first request begins a
    # session, and every later request must carry the SessionId the server returned.
    if session_id:
        return f'<Header><Session xmlns="{_XMLA_NS}" mustUnderstand="1" SessionId="{escape(session_id, {chr(34): "&quot;"})}"/></Header>'
    return f'<Header><BeginSession xmlns="{_XMLA_NS}" mustUnderstand="1"/></Header>'


def _elements(values: Mapping[str, str] | None) -> str:
    if not values:
        return ""
    parts = []
    for name, value in values.items():
        if not _ELEMENT_NAME.match(name):
            raise ValueError(f"invalid XMLA element name: {name!r}")
        parts.append(f"<{name}>{escape(str(value))}</{name}>")
    return "".join(parts)
