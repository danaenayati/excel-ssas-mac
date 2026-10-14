"""Cube formulas through the real SSAS backend, against the loopback stand-in server."""

import html
import re
import threading
import time

import pytest
import xml_samples as x
from fake_ssas import FakeSsas

from cubebridge.backends.ssas_backend import SsasBackend

CITY = "[Geo].[City]"
KNOWN = {"[Geo].[City].&[1]": ("Tehran", 120), "[Geo].[City].&[2]": ("Shiraz", 80)}
_NAME = re.compile(r"\[(?:[^\]]|\]\])*\](?:\.\[(?:[^\]]|\]\])*\]|\.?&\[(?:[^\]]|\]\])*\])*")


def statement(request: str) -> str:
    return html.unescape(re.search(r"<Statement>(.*)</Statement>", request, re.S).group(1))


def values_response(request: str) -> bytes:
    """SELECT {(…), …} ON 0: a fault when any member is unknown, as a cube in Error mode."""
    text = statement(request)
    tuples = [_NAME.findall(t) for t in re.findall(r"\(([^()]*)\)", text[: text.index(" ON 0")])]
    for names in tuples:
        for name in names:
            if name.startswith("[Geo]") and name not in KNOWN:
                return x.soap_fault(f"Query (1, 9) The member '{name}' was not found in the cube.")
    positions = [[x.member(CITY, t[0], KNOWN[t[0]][0], 1), x.member("[Measures]", t[1], "Rides", 0)] for t in tuples]
    cells = [x.cell(i, str(KNOWN[t[0]][1]), "xsd:int", "#,##0") for i, t in enumerate(tuples)]
    return x.cellset(axes=[x.axis("Axis0", positions)], cells=cells)


def members_response(request: str) -> bytes:
    text = statement(request)
    names = _NAME.findall(text[text.index("ON 0, {") : text.index("ON 1")])
    rows = [[x.member(CITY, n, KNOWN[n][0], 1)] for n in names if n in KNOWN]  # missing members dropped
    return x.cellset(axes=[x.axis("Axis0", []), x.axis("Axis1", rows)], cells=[])


@pytest.fixture
def backend(ntlm_users):
    user, password = ntlm_users
    responses = {"CELL PROPERTIES VALUE": values_response, "SELECT {} ON 0, {": members_response}
    with FakeSsas(responses) as server:
        made = SsasBackend(f"127.0.0.1:{server.port}", user, password)
        made.server = server
        yield made
        made.close()


def test_values_are_asked_in_one_query_per_group(backend):
    answers = backend.cells("Sales", "Sales", [["[Measures].[Rides]", "[Geo].[City].&[1]"], ["[Geo].[City].&[2]", "[Measures].[Rides]"]])
    assert answers == [{"value": 120}, {"value": 80}]
    sent = [statement(r) for r in backend.server.requests if "<Statement>" in r]
    assert sent == [
        "SELECT {([Geo].[City].&[1], [Measures].[Rides]), ([Geo].[City].&[2], [Measures].[Rides])} "
        "ON 0 FROM [Sales] CELL PROPERTIES VALUE"
    ]
    assert "<Catalog>Sales</Catalog>" in backend.server.requests[-1]


def test_a_member_the_cube_rejects_fails_alone(backend):
    tuples = [["[Geo].[City].&[1]", "[Measures].[Rides]"], ["[Geo].[City].&[9]", "[Measures].[Rides]"], ["[Geo].[City].&[2]", "[Measures].[Rides]"]]
    answers = backend.cells("Sales", "Sales", tuples)
    assert answers[0] == {"value": 120} and answers[2] == {"value": 80}
    assert answers[1]["error"]["category"] == "not-found" and "&[9]" in answers[1]["error"]["detail"]
    # The session survives a fault, so every query went over the one connection.
    assert backend.server.connections == 1


def test_member_captions_with_missing_members_dropped_by_the_cube(backend):
    answers = backend.lookup(
        "Sales",
        "Sales",
        [{"kind": "member", "members": ["[Geo].[City].&[2]"]}, {"kind": "member", "members": ["[Geo].[City].&[7]"]}],
    )
    assert answers[0] == {"value": "Shiraz"}
    assert answers[1]["error"]["category"] == "not-found"


def test_formulas_are_not_stopped_by_the_panes_cancel(backend):
    backend.server.delay = 0.6
    threading.Timer(0.2, backend.cancel).start()
    started = time.monotonic()
    answers = backend.cells("Sales", "Sales", [["[Geo].[City].&[1]", "[Measures].[Rides]"]])
    assert answers == [{"value": 120}]
    assert time.monotonic() - started >= 0.5
    assert backend.server.connections == 1  # cancel had no query of the pane to stop
