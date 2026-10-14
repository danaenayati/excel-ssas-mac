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
# A tabular model names its blank item with a bare "&" for the key: `[T].[Column].&`.
_UNIQUE_NAME = re.compile(rf"{_PART}(?:\.{_PART}|\.?&{_PART})*(?:\.&)?")
MAX_NAME = 1000


def validate_unique_name(name: object) -> str:
    if not isinstance(name, str) or not name or len(name) > MAX_NAME or not _UNIQUE_NAME.fullmatch(name):
        shown = name if isinstance(name, str) else type(name).__name__
        raise CubeError(
            "bad-request",
            "The report refers to a field or member name that is not valid.",
            f"rejected name: {shown[:120]!r}",
        )
    return name


_BRACKETED = re.compile(_PART)


def hierarchy_of(unique_name: str) -> str:
    """The hierarchy part of a member's unique name, lower-cased for comparison."""
    parts = _BRACKETED.findall(unique_name)
    return ".".join(parts[:2]).lower() if len(parts) >= 2 else unique_name.lower()


def quote_identifier(name: str) -> str:
    """Bracket-quote a plain name such as a cube name."""
    if any(ord(ch) < 0x20 for ch in name):
        raise CubeError("bad-request", "The cube name contains characters that are not allowed.")
    return "[" + name.replace("]", "]]") + "]"


def quote_string(text: str) -> str:
    """An MDX string literal."""
    return '"' + text.replace('"', '""') + '"'


MAX_EXPRESSION = 2000
# Words that start or shape a statement. An expression never needs them, and refusing
# them keeps a formula from a shared workbook to one value inside the query it joins.
_STATEMENT_WORDS = frozenset(
    {
        "select", "from", "where", "with", "drillthrough", "create", "drop", "alter", "update",
        "call", "insert", "delete", "refresh", "clear", "scope", "freeze",
    }
)
_CLOSING = {")": "(", "}": "{"}


def validate_expression(text: object, what: str, max_length: int = MAX_EXPRESSION) -> str:
    """An MDX value or set expression typed by the user (a calculated measure, a
    calculated item, a set for a cube formula), made safe to place in a query.

    Bracketed names and string literals are skipped over; outside them, brackets must
    balance, and ';', comments and statement keywords are refused. The caller still
    wraps the result in parentheses, so a valid expression cannot end the clause it is
    placed in. Line breaks become spaces.
    """
    if not isinstance(text, str) or not text.strip():
        raise CubeError("bad-request", f"{what} is empty.")
    if len(text) > max_length:
        raise CubeError("bad-request", f"{what} is longer than {max_length:,} characters.")
    if any(ord(ch) < 0x20 and ch not in "\t\r\n" for ch in text):
        raise CubeError("bad-request", f"{what} contains characters that are not allowed.")
    flat = " ".join(text.split())

    def refuse(reason: str) -> CubeError:
        return CubeError("bad-request", f"{what} is not a formula Cube Explorer can use: {reason}.")

    stack: list[str] = []
    i, n = 0, len(flat)
    while i < n:
        ch = flat[i]
        if ch == "[":
            end = i + 1
            while True:
                end = flat.find("]", end)
                if end < 0:
                    raise refuse("a '[' is not closed")
                if end + 1 < n and flat[end + 1] == "]":
                    end += 2
                    continue
                break
            i = end + 1
            continue
        if ch in "\"'":
            end = i + 1
            while True:
                end = flat.find(ch, end)
                if end < 0:
                    raise refuse("a text in quotes is not closed")
                if end + 1 < n and flat[end + 1] == ch:
                    end += 2
                    continue
                break
            i = end + 1
            continue
        if ch == ";":
            raise refuse("';' is not allowed")
        if flat.startswith(("--", "//", "/*"), i):
            raise refuse("comments are not allowed")
        if ch in "({":
            stack.append(ch)
        elif ch in ")}":
            if not stack or stack.pop() != _CLOSING[ch]:
                raise refuse(f"a '{ch}' has no matching opening bracket")
        elif ch.isalpha() or ch == "_":
            end = i
            while end < n and (flat[end].isalnum() or flat[end] == "_"):
                end += 1
            if flat[i:end].lower() in _STATEMENT_WORDS:
                raise refuse(f"'{flat[i:end]}' starts a statement, not a formula")
            i = end
            continue
        i += 1
    if stack:
        raise refuse(f"a '{stack[-1]}' is not closed")
    return flat
