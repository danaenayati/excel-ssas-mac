"""Convert an SSAS FORMAT_STRING into an Excel number format.

SSAS format strings follow Visual Basic's Format(): a set of named formats plus
custom strings that mostly share Excel's syntax. Excel rejects a whole range write if
one format in it is invalid, so anything not clearly safe becomes "General". The cell
then still holds the right number; only its display is plain.
"""

from __future__ import annotations

from functools import lru_cache

GENERAL = "General"

_NAMED = {
    "": GENERAL,
    "general": GENERAL,
    "general number": GENERAL,
    "standard": "#,##0.00",
    "fixed": "0.00",
    "currency": "$#,##0.00",
    "percent": "0.00%",
    "scientific": "0.00E+00",
    "yes/no": GENERAL,
    "true/false": GENERAL,
    "on/off": GENERAL,
    "general date": "yyyy-mm-dd hh:mm",
    "long date": "dddd, mmmm d, yyyy",
    "medium date": "d-mmm-yy",
    "short date": "yyyy-mm-dd",
    "long time": "hh:mm:ss",
    "medium time": "hh:mm AM/PM",
    "short time": "hh:mm",
}

# Outside quoted literals, only digit placeholders, separators, signs and common
# currency symbols are allowed, plus an exponent (E+00). Letters outside quotes mean
# date tokens or literal text, whose rules differ between VB and Excel.
_PLAIN = frozenset("#0?.,%;()+- $€£¥﷼")
_MAX = 120


def _is_safe(text: str) -> bool:
    """One pass, no backtracking: a cube can put anything in a format string."""
    i, n = 0, len(text)
    while i < n:
        ch = text[i]
        if ch == '"':
            end = text.find('"', i + 1)
            if end < 0:
                return False
            i = end + 1
        elif ch == "E":
            if i + 2 >= n or text[i + 1] not in "+-" or text[i + 2] != "0":
                return False
            i += 3
            while i < n and text[i] == "0":
                i += 1
        elif ch in _PLAIN:
            i += 1
        else:
            return False
    return True


@lru_cache(maxsize=2048)
def to_excel(format_string: str | None) -> str:
    text = (format_string or "").strip()
    named = _NAMED.get(text.lower())
    if named is not None:
        return named
    if len(text) > _MAX or any(ord(ch) < 0x20 for ch in text):
        return GENERAL
    # VB allows a fourth section (nulls); Excel's fourth section is for text.
    sections = _split_sections(text)
    if sections is None:
        return GENERAL
    kept = ";".join(sections[:3])
    if not kept or not _is_safe(kept):
        return GENERAL
    return kept


def _split_sections(text: str) -> list[str] | None:
    sections, current, quoted = [], [], False
    for ch in text:
        if ch == '"':
            quoted = not quoted
        if ch == ";" and not quoted:
            sections.append("".join(current))
            current = []
        else:
            current.append(ch)
    if quoted:
        return None
    sections.append("".join(current))
    return sections
