import inspect
import xml.etree.ElementTree as ET

import pytest

from cubebridge.ssas import envelopes

XMLA = "{urn:schemas-microsoft-com:xml-analysis}"
SOAP = "{http://schemas.xmlsoap.org/soap/envelope/}"
EXT = "{http://schemas.microsoft.com/analysisservices/2003/ext}"


def parse(payload: bytes) -> ET.Element:
    return ET.fromstring(payload)


def test_authenticate_lives_in_the_ext_namespace_not_the_xmla_one():
    root = parse(envelopes.authenticate("QUJD"))
    token = root.find(f"{SOAP}Body/{EXT}Authenticate/{EXT}SspiHandshake")
    assert token is not None and token.text == "QUJD"


def test_first_request_begins_a_session_and_later_ones_carry_its_id():
    first = parse(envelopes.discover("DBSCHEMA_CATALOGS", session_id=None))
    assert first.find(f"{SOAP}Header/{XMLA}BeginSession") is not None
    later = parse(envelopes.discover("DBSCHEMA_CATALOGS", session_id="ABC-1"))
    session = later.find(f"{SOAP}Header/{XMLA}Session")
    assert session is not None and session.get("SessionId") == "ABC-1"
    assert later.find(f"{SOAP}Header/{XMLA}BeginSession") is None


def test_discover_restrictions_and_catalog_property():
    root = parse(
        envelopes.discover(
            "MDSCHEMA_CUBES",
            restrictions={"CATALOG_NAME": "Sales & <Ops>"},
            catalog="Sales & <Ops>",
            session_id="S",
        )
    )
    discover = root.find(f"{SOAP}Body/{XMLA}Discover")
    assert discover.find(f"{XMLA}RequestType").text == "MDSCHEMA_CUBES"
    restriction = discover.find(f"{XMLA}Restrictions/{XMLA}RestrictionList/{XMLA}CATALOG_NAME")
    assert restriction.text == "Sales & <Ops>"
    assert discover.find(f"{XMLA}Properties/{XMLA}PropertyList/{XMLA}Catalog").text == "Sales & <Ops>"


def test_restriction_names_are_validated_not_escaped():
    with pytest.raises(ValueError):
        envelopes.discover("X", restrictions={"A><B": "1"})


def test_execute_asks_for_a_multidimensional_tuple_result():
    mdx = 'SELECT {[Measures].[A & B]} ON 0 FROM [Sales] WHERE ("<x>")'
    root = parse(envelopes.execute(mdx, catalog="Sales", session_id="S"))
    execute = root.find(f"{SOAP}Body/{XMLA}Execute")
    assert execute.find(f"{XMLA}Command/{XMLA}Statement").text == mdx
    props = execute.find(f"{XMLA}Properties/{XMLA}PropertyList")
    assert props.find(f"{XMLA}Catalog").text == "Sales"
    assert props.find(f"{XMLA}Format").text == "Multidimensional"
    assert props.find(f"{XMLA}AxisFormat").text == "TupleFormat"
    assert props.find(f"{XMLA}Content").text == "Data"


def test_the_module_can_only_build_authenticate_discover_and_execute():
    public = {
        name
        for name, obj in vars(envelopes).items()
        if inspect.isfunction(obj) and obj.__module__ == envelopes.__name__ and not name.startswith("_")
    }
    assert public == {"authenticate", "discover", "execute"}
