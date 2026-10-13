"""Canned XMLA responses in the shape SQL Server Analysis Services sends them."""

SOAP = "http://schemas.xmlsoap.org/soap/envelope/"
XMLA = "urn:schemas-microsoft-com:xml-analysis"


def envelope(body: str, session_id: str | None = "S-1") -> bytes:
    header = (
        f'<soap:Header><Session xmlns="{XMLA}" SessionId="{session_id}"/></soap:Header>'
        if session_id
        else ""
    )
    return (
        f'<soap:Envelope xmlns:soap="{SOAP}">{header}<soap:Body>{body}</soap:Body></soap:Envelope>'
    ).encode("utf-8")


def rowset(rows: list[dict[str, str]], session_id: str | None = "S-1") -> bytes:
    body_rows = "".join(
        "<row>" + "".join(f"<{k}>{_esc(v)}</{k}>" for k, v in row.items()) + "</row>" for row in rows
    )
    return envelope(
        f'<DiscoverResponse xmlns="{XMLA}"><return>'
        f'<root xmlns="{XMLA}:rowset" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xmlns:xsd="http://www.w3.org/2001/XMLSchema">'
        f'<xsd:schema targetNamespace="{XMLA}:rowset" elementFormDefault="qualified">'
        f'<xsd:element name="root"><xsd:complexType><xsd:sequence>'
        f'<xsd:element name="row" type="row" minOccurs="0" maxOccurs="unbounded"/>'
        f"</xsd:sequence></xsd:complexType></xsd:element></xsd:schema>"
        f"{body_rows}</root></return></DiscoverResponse>",
        session_id,
    )


def member(hierarchy: str, uname: str, caption: str, lnum: int, display_info: int = 0) -> str:
    return (
        f'<Member Hierarchy="{_esc(hierarchy)}"><UName>{_esc(uname)}</UName>'
        f"<Caption>{_esc(caption)}</Caption><LName>{_esc(hierarchy)}.[L{lnum}]</LName>"
        f"<LNum>{lnum}</LNum><DisplayInfo>{display_info}</DisplayInfo></Member>"
    )


def axis(name: str, tuples: list[list[str]]) -> str:
    inner = "".join("<Tuple>" + "".join(members) + "</Tuple>" for members in tuples)
    return f'<Axis name="{name}"><Tuples>{inner}</Tuples></Axis>'


def cell(ordinal: int, value: str, xsi_type: str = "xsd:double", fmt: str = "#,##0.00", formatted: str = "") -> str:
    return (
        f'<Cell CellOrdinal="{ordinal}"><Value xsi:type="{xsi_type}">{_esc(value)}</Value>'
        f"<FmtValue>{_esc(formatted or value)}</FmtValue><FormatString>{_esc(fmt)}</FormatString></Cell>"
    )


def error_cell(ordinal: int, description: str) -> str:
    return (
        f'<Cell CellOrdinal="{ordinal}"><Value><Error><ErrorCode>3238658065</ErrorCode>'
        f"<Description>{_esc(description)}</Description></Error></Value></Cell>"
    )


def cellset(axes: list[str], cells: list[str], session_id: str | None = "S-1") -> bytes:
    return envelope(
        f'<ExecuteResponse xmlns="{XMLA}"><return>'
        f'<root xmlns="{XMLA}:mddataset" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        f'xmlns:xsd="http://www.w3.org/2001/XMLSchema">'
        f"<OlapInfo><CubeInfo><Cube><CubeName>Sales</CubeName></Cube></CubeInfo>"
        f'<AxesInfo><AxisInfo name="Axis0"/></AxesInfo>'
        f'<CellInfo><Value name="VALUE"/><FmtValue name="FORMATTED_VALUE"/>'
        f'<FormatString name="FORMAT_STRING"/></CellInfo></OlapInfo>'
        f"<Axes>{''.join(axes)}"
        f'<Axis name="SlicerAxis"><Tuples><Tuple>'
        f"{member('[Product].[Category]', '[Product].[Category].[All]', 'All', 0)}"
        f"</Tuple></Tuples></Axis></Axes>"
        f"<CellData>{''.join(cells)}</CellData></root></return></ExecuteResponse>",
        session_id,
    )


def soap_fault(message: str, code: str = "XMLAnalysisError.0xc10f0007") -> bytes:
    return (
        f'<soap:Envelope xmlns:soap="{SOAP}"><soap:Body><soap:Fault xmlns="{SOAP}">'
        f"<faultcode>{code}</faultcode><faultstring>{_esc(message)}</faultstring>"
        f'<detail><Error ErrorCode="3239116807" Description="{_esc(message)}" '
        f'Source="Microsoft SQL Server 2022 Analysis Services" HelpFile=""/></detail>'
        f"</soap:Fault></soap:Body></soap:Envelope>"
    ).encode("utf-8")


def inband_error(message: str) -> bytes:
    return envelope(
        f'<ExecuteResponse xmlns="{XMLA}"><return><root xmlns="{XMLA}:empty">'
        f'<Exception xmlns="{XMLA}:exception"/><Messages xmlns="{XMLA}:exception">'
        f'<Error ErrorCode="3238658121" Description="{_esc(message)}" '
        f'Source="Microsoft SQL Server 2022 Analysis Services" HelpFile=""/>'
        f"</Messages></root></return></ExecuteResponse>"
    )


def authenticate_response(token_b64: str) -> bytes:
    inner = f"<SspiHandshake>{token_b64}</SspiHandshake>" if token_b64 else "<SspiHandshake/>"
    return (
        f'<soap:Envelope xmlns:soap="{SOAP}"><soap:Body>'
        f'<AuthenticateResponse xmlns="http://schemas.microsoft.com/analysisservices/2003/ext">'
        f"<return>{inner}</return></AuthenticateResponse></soap:Body></soap:Envelope>"
    ).encode("utf-8")


def _esc(text: str) -> str:
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )
