import base64

import pytest
import xml_samples as x

from cubebridge.errors import CubeError, ProtocolError
from cubebridge.ssas import xmla

MEASURES = "[Measures]"
CITY = "[Geo].[City]"


def test_rowset_rows_and_session_id():
    parsed = xmla.parse(
        x.rowset(
            [
                {"CATALOG_NAME": "Sales", "DESCRIPTION": ""},
                {"CATALOG_NAME": "Ops & <More>", "DESCRIPTION": "second"},
            ],
            session_id="ABC-9",
        )
    )
    assert parsed.rows == [
        {"CATALOG_NAME": "Sales", "DESCRIPTION": ""},
        {"CATALOG_NAME": "Ops & <More>", "DESCRIPTION": "second"},
    ]
    assert parsed.session_id == "ABC-9"


def test_rowset_column_names_are_decoded():
    parsed = xmla.parse(x.rowset([{"Display_x0020_Name": "v"}]))
    assert parsed.rows == [{"Display Name": "v"}]


def test_soap_fault_raises_server_error_with_the_servers_words():
    with pytest.raises(CubeError) as info:
        xmla.parse(x.soap_fault("The cube 'Nope' does not exist."))
    assert info.value.category == "server"
    assert "Nope" in info.value.detail


def test_fault_during_authentication_is_an_auth_error():
    with pytest.raises(CubeError) as info:
        xmla.parse(x.soap_fault("Logon failure: unknown user name or bad password."), during_auth=True)
    assert info.value.category == "auth"


def test_inband_exception_is_a_server_error_too():
    with pytest.raises(CubeError) as info:
        xmla.parse(x.inband_error("Query (1, 8) The member '[X]' was not found in the cube."))
    assert info.value.category == "server"
    assert "was not found" in info.value.detail


def test_access_denied_text_is_reported_as_permission_problem():
    with pytest.raises(CubeError) as info:
        xmla.parse(x.soap_fault("Either the user, 'CORP\\dana', does not have access to the 'HR' database, or the database does not exist."))
    assert info.value.category == "server"
    assert "access" in info.value.message.lower()


def test_two_axis_cellset():
    data = x.cellset(
        axes=[
            x.axis(
                "Axis0",
                [
                    [x.member(MEASURES, "[Measures].[Rides]", "Rides", 0)],
                    [x.member(MEASURES, "[Measures].[Revenue]", "Revenue", 0)],
                ],
            ),
            x.axis(
                "Axis1",
                [
                    [x.member(CITY, "[Geo].[City].&[1]", "Tehran", 1, display_info=3)],
                    [x.member(CITY, "[Geo].[City].&[2]", "Shiraz & Co", 1)],
                ],
            ),
        ],
        cells=[
            x.cell(0, "120", "xsd:int", "#,##0"),
            x.cell(1, "3500.5", "xsd:double", "Currency", "$3,500.50"),
            # ordinal 2 is absent: sparse
            x.cell(3, "9.25E2", "xsd:double", "Currency"),
        ],
    )
    cs = xmla.parse(data).cellset
    assert len(cs.axes) == 2  # slicer axis is not a display axis
    assert [t[0].caption for t in cs.axes[0]] == ["Rides", "Revenue"]
    tehran = cs.axes[1][0][0]
    assert (tehran.unique_name, tehran.caption, tehran.hierarchy, tehran.level, tehran.child_count) == (
        "[Geo].[City].&[1]",
        "Tehran",
        CITY,
        1,
        3,
    )
    assert cs.axes[1][1][0].caption == "Shiraz & Co"
    assert cs.cells[0].value == 120 and isinstance(cs.cells[0].value, int)
    assert cs.cells[1].value == 3500.5 and cs.cells[1].format_string == "Currency"
    assert cs.cells[1].formatted == "$3,500.50"
    assert 2 not in cs.cells
    assert cs.cells[3].value == 925.0


def test_display_info_high_bits_are_not_counted_as_children():
    data = x.cellset(
        axes=[x.axis("Axis0", [[x.member(CITY, "[Geo].[City].&[1]", "Tehran", 1, display_info=0x10000 | 5)]])],
        cells=[],
    )
    assert xmla.parse(data).cellset.axes[0][0][0].child_count == 5


def test_single_axis_and_empty_axis():
    one = xmla.parse(x.cellset(axes=[x.axis("Axis0", [[x.member(MEASURES, "[Measures].[Rides]", "Rides", 0)]])], cells=[x.cell(0, "7", "xsd:long")])).cellset
    assert len(one.axes) == 1 and one.cells[0].value == 7
    empty = xmla.parse(x.cellset(axes=[x.axis("Axis0", []), x.axis("Axis1", [])], cells=[])).cellset
    assert empty.axes == [[], []] and empty.cells == {}


@pytest.mark.parametrize(
    "text, xsi_type, expected",
    [
        ("12", "xsd:short", 12),
        ("4000000000", "xsd:unsignedInt", 4000000000),
        ("1.5", "xsd:decimal", 1.5),
        ("true", "xsd:boolean", True),
        ("false", "xsd:boolean", False),
        ("hello", "xsd:string", "hello"),
        ("2026-10-13T00:00:00", "xsd:dateTime", "2026-10-13T00:00:00"),
    ],
)
def test_cell_value_types(text, xsi_type, expected):
    data = x.cellset(axes=[x.axis("Axis0", [[x.member(MEASURES, "[Measures].[M]", "M", 0)]])], cells=[x.cell(0, text, xsi_type)])
    value = xmla.parse(data).cellset.cells[0].value
    assert value == expected and type(value) is type(expected)


def test_error_cell_becomes_an_excel_error_text():
    data = x.cellset(
        axes=[x.axis("Axis0", [[x.member(MEASURES, "[Measures].[M]", "M", 0)]])],
        cells=[x.error_cell(0, "MdxScript(Sales) (12, 4) Division by zero")],
    )
    assert xmla.parse(data).cellset.cells[0].value == "#VALUE!"


def test_handshake_token_is_extracted_and_empty_token_is_empty():
    token = base64.b64encode(b"NTLMSSP\x00challenge").decode()
    assert xmla.parse(x.authenticate_response(token)).handshake_token == b"NTLMSSP\x00challenge"
    assert xmla.parse(x.authenticate_response("")).handshake_token == b""


def test_garbage_is_a_protocol_error_not_an_empty_result():
    with pytest.raises(ProtocolError):
        xmla.parse(b"\x9f\x00\x12 not xml at all")


# -- review fixes --------------------------------------------------------------------
def test_a_non_numeric_cell_ordinal_is_a_protocol_error():
    bad = x.cellset(
        axes=[x.axis("Axis0", [[x.member(MEASURES, "[Measures].[M]", "M", 0)]])],
        cells=['<Cell CellOrdinal="abc"><Value xsi:type="xsd:int">1</Value></Cell>'],
    )
    with pytest.raises(ProtocolError):
        xmla.parse(bad)


def test_errors_can_be_returned_instead_of_raised_so_the_session_id_survives():
    parsed = xmla.parse(x.inband_error("Query (1, 1) bad"), raise_errors=False)
    assert parsed.error is not None and parsed.error.category == "server"
    assert parsed.session_id == "S-1"


# -- first contact with a real server ------------------------------------------------
def test_token_with_line_breaks_inside_is_decoded():
    raw = b"NTLMSSP\x00" + bytes(range(256)) * 2
    encoded = base64.b64encode(raw).decode()
    wrapped = "\r\n" + "".join(encoded[i : i + 64] + "\r\n" for i in range(0, len(encoded), 64)) + "  "
    assert "\r\n" in wrapped.strip()
    assert xmla.parse(x.authenticate_response(wrapped)).handshake_token == raw


def test_token_without_padding_is_decoded():
    assert xmla.parse(x.authenticate_response("QUJDRA")).handshake_token == b"ABCD"


def test_a_token_that_really_is_not_base64_is_reported_by_shape_not_content():
    parsed = xmla.parse(x.authenticate_response("secret!token*value"))  # parsing alone must not fail
    with pytest.raises(ProtocolError) as info:
        parsed.handshake_token
    detail = info.value.detail
    assert "U+0021" in detail and "U+002A" in detail and "18 characters" in detail
    assert "secret" not in detail and "token" not in detail.replace("sign-in token", "")


# -- Show Details ------------------------------------------------------------------
def test_drillthrough_rowset_columns_follow_the_schema_and_values_are_typed():
    data = x.drillthrough_rowset(
        [("[$Trips].[City]", "xsd:string"), ("[$Trips].[Rides]", "xsd:long"), ("[$Trips].[Fare]", "xsd:double"), ("[$Trips].[Paid]", "xsd:boolean")],
        [["Tehran", "12", "4.5", "true"], ["Shiraz", None, "3", "false"], ["00123", "7", None, None]],
    )
    parsed = xmla.parse(data)
    assert parsed.columns == ["[$Trips].[City]", "[$Trips].[Rides]", "[$Trips].[Fare]", "[$Trips].[Paid]"]
    assert xmla.typed_rows(parsed) == [
        ["Tehran", 12, 4.5, True],
        ["Shiraz", None, 3.0, False],
        ["00123", 7, None, None],  # a string column keeps its leading zeros
    ]


def test_a_rowset_without_schema_falls_back_to_first_seen_columns_as_text():
    parsed = xmla.parse(x.rowset([{"A": "1", "B": "x"}, {"B": "y", "C": "2"}]))
    assert xmla.typed_rows(parsed) == [["1", "x", None], [None, "y", "2"]]


def test_a_member_with_an_empty_caption_reads_blank_as_in_excel():
    data = x.cellset(axes=[x.axis("Axis0", [[x.member("[Geo].[City]", "[Geo].[City].&[]", "", 1)]])], cells=[])
    assert xmla.parse(data).cellset.axes[0][0][0].caption == "(blank)"
