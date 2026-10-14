import pytest

from cubebridge.errors import CubeError
from cubebridge.pivot.names import validate_expression


@pytest.mark.parametrize(
    "text",
    [
        "[Measures].[Revenue] / [Measures].[Rides]",
        "IIF([Measures].[Rides] = 0, NULL, [Measures].[Revenue] / [Measures].[Rides])",
        "Aggregate({[Geo].[City].&[1], [Geo].[City].&[2]})",
        '[Measures].[A] + Len("a; b -- c")',
        "[Odd ]] name; select].[x] * 2",
        "Sum(YTD([Date].[Calendar].CurrentMember), [Measures].[Rides])",
        "TopCount([Geo].[City].[All].Children, 5, [Measures].[Rides])",
    ],
)
def test_formulas_are_accepted(text):
    assert validate_expression(text, "The formula") == text


def test_line_breaks_become_spaces():
    assert validate_expression("[Measures].[A]\n  + 1", "The formula") == "[Measures].[A] + 1"


@pytest.mark.parametrize(
    "text, reason",
    [
        ("", "empty"),
        ("   ", "empty"),
        ("[Measures].[A]; DROP", "';'"),
        ("1 SELECT {} ON 0 FROM [Model]", "statement"),
        ("[Measures].[A] -- rest", "comments"),
        ("[Measures].[A] /* x */", "comments"),
        ("([Measures].[A]", "not closed"),
        ("[Measures].[A])", "no matching"),
        ("{[Measures].[A])", "no matching"),
        ("[Measures].[A", "'['"),
        ('"open', "quotes"),
        ("1 + 2" * 500, "longer"),
        ("a\x00b", "not allowed"),
    ],
)
def test_formulas_that_could_leave_their_place_are_refused(text, reason):
    with pytest.raises(CubeError) as info:
        validate_expression(text, "The formula")
    assert reason in info.value.message
