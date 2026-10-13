"""Validation and quoting for everything that is placed into MDX text.

Member, hierarchy and measure unique names reach the helper from the pane, which may
have loaded them from a workbook someone else built. They are accepted only in the
shape SSAS itself produces: bracketed parts joined by dots, with optional `&[key]`
parts. Inside brackets every character is literal (`]` is doubled), so a name that
passes cannot terminate the expression it is placed in.
"""

from __future__ import annotations

import re

from ..errors import CubeError

_PART = r"\[(?:[^\]\x00-\x1f]|\]\])*\]"
_UNIQUE_NAME = re.compile(rf"^{_PART}(?:\.{_PART}|\.?&{_PART})*$")
MAX_NAME = 1000


def validate_unique_name(name: object) -> str:
    if not isinstance(name, str) or not name or len(name) > MAX_NAME or not _UNIQUE_NAME.match(name):
        shown = name if isinstance(name, str) else type(name).__name__
        raise CubeError(
            "bad-request",
            "The report refers to a field or member name that is not valid.",
            f"rejected name: {shown[:120]!r}",
        )
    return name


def quote_identifier(name: str) -> str:
    """Bracket-quote a plain name such as a cube name."""
    if any(ord(ch) < 0x20 for ch in name):
        raise CubeError("bad-request", "The cube name contains characters that are not allowed.")
    return "[" + name.replace("]", "]]") + "]"


def quote_string(text: str) -> str:
    """An MDX string literal."""
    return '"' + text.replace('"', '""') + '"'
