"""The sample cube's answers to cube formulas (formulas.py), so that CUBE.VALUE and its
siblings can be tried, and tested, without a server.

Values are worked out from the sample's fact rows, as demo_backend.py works out a
report. No MDX is parsed, except for the few set shapes `_set` understands:
`{…, …}`, `<member>.Children`, `<hierarchy>.Members`, `<hierarchy>.[Level].Members`,
a single member, and `TopCount` / `BottomCount(<set>, n, <measure>)`.

Answers have the same shape and the same per-item failures as a real cube's.
"""

from __future__ import annotations

import re

from .. import formulas
from . import demo_backend as demo

_MEASURE_NAMES = {name.lower(): name for name in demo._MEASURES}
_TOP = re.compile(r"^(TopCount|BottomCount)\s*\((.*),\s*(\d+)\s*,\s*(\[[^\]]+\]\.\[[^\]]+\])\s*\)$", re.I)
_SUFFIX = re.compile(r"^(.*)\.(Children|Members|AllMembers)$", re.I)
_SETS = (
    "The sample cube understands only these sets: {…}, <member>.Children, <field>.Members, "
    "<field>.[Level].Members, a single member, and TopCount or BottomCount(<set>, n, <measure>)."
)


def cells(backend: demo.DemoBackend, tuples: list[list[str]]) -> list[dict]:
    out = []
    for raw in tuples:
        prepared = formulas.prepare_tuple(raw)
        out.append(prepared if isinstance(prepared, dict) else _value(backend, prepared))
    return out


def lookup(backend: demo.DemoBackend, items: list[dict]) -> list[dict]:
    out = []
    for item in items:
        prepared = formulas.prepare_lookup(item)
        if isinstance(prepared, dict):
            out.append(prepared)
        elif prepared.kind == "member":
            out.append(_member(backend, prepared.members))
        elif prepared.kind == "property":
            out.append(_property(backend, prepared.members[0], prepared.prop))
        else:
            found = _set(backend, prepared.set)
            if isinstance(found, dict):
                out.append(found)
            elif prepared.kind == "count":
                out.append(formulas.ok(len(found)))
            elif prepared.rank > len(found):
                out.append(formulas.fail("not-found", f"The set has {len(found):,} items, fewer than {prepared.rank:,}."))
            else:
                out.append(formulas.ok(_caption(backend, found[prepared.rank - 1])))
    return out


# -- names -------------------------------------------------------------------------
def _keys(backend: demo.DemoBackend) -> dict[str, tuple]:
    """Lower-cased unique name -> (hierarchy, path). Names are not case-sensitive in a cube."""
    cached = getattr(backend, "_formula_keys", None)
    if cached is None:
        cached = {name.lower(): key for name, key in backend._by_name.items()}
        backend._formula_keys = cached
    return cached


def _missing(name: str) -> dict:
    return formulas.fail("not-found", f"The sample cube has no member {name}.")


def _caption(backend: demo.DemoBackend, name: str) -> str | None:
    if formulas.hierarchy_key(name) == formulas.MEASURES:
        measure = _MEASURE_NAMES.get(name.lower())
        return demo._MEASURES[measure][0] if measure else None
    key = _keys(backend).get(name.lower())
    return backend._caption(key) if key else None


# -- values ------------------------------------------------------------------------
def _aggregate(backend: demo.DemoBackend, keys: list[tuple], measure: str) -> object:
    slot, found = [0, 0.0, 0], False
    for fact in backend._facts:
        if all(backend._path_of(h, fact)[: len(path)] == path for h, path in keys):
            demo._add(slot, fact)
            found = True
    return demo._measure_value(measure, slot) if found else None


def _value(backend: demo.DemoBackend, names: tuple[str, ...]) -> dict:
    measure = demo._DEFAULT_MEASURE
    keys = []
    for name in names:
        if formulas.hierarchy_key(name) == formulas.MEASURES:
            measure = _MEASURE_NAMES.get(name.lower())
            if measure is None:
                return _missing(name)
            continue
        key = _keys(backend).get(name.lower())
        if key is None:
            return _missing(name)
        keys.append(key)
    return formulas.ok(_aggregate(backend, keys, measure))


# -- members -----------------------------------------------------------------------
def _member(backend: demo.DemoBackend, names: tuple[str, ...]) -> dict:
    captions = []
    for name in names:
        caption = _caption(backend, name)
        if caption is None:
            return _missing(name)
        captions.append(caption)
    return formulas.ok(captions[-1])


def _property(backend: demo.DemoBackend, name: str, prop: str) -> dict:
    key = _keys(backend).get(name.lower())
    if key is None:
        return _missing(name)
    which = prop.lower()
    if which in ("caption", "member_caption", "name", "member_name"):
        value = backend._caption(key)
    elif which in ("key", "key0", "member_key"):
        value = str(key[1][-1]) if key[1] else ""
    elif which in ("uniquename", "unique_name", "member_unique_name"):
        value = backend._name_of[key]
    elif which == "level_number":
        value = str(len(key[1]))
    else:
        return formulas.fail("not-found", f"The member has no property called {prop}.")
    return formulas.ok(value)


# -- sets --------------------------------------------------------------------------
def _split(text: str) -> list[str]:
    """Split on the commas that are not inside brackets, braces or parentheses."""
    parts, depth, start, i = [], 0, 0, 0
    while i < len(text):
        ch = text[i]
        if ch == "[":
            end = i + 1
            while True:
                end = text.find("]", end)
                if end < 0 or not text.startswith("]]", end):
                    break
                end += 2
            i = len(text) if end < 0 else end + 1
            continue
        if ch in "({":
            depth += 1
        elif ch in ")}":
            depth -= 1
        elif ch == "," and depth == 0:
            parts.append(text[start:i].strip())
            start = i + 1
        i += 1
    parts.append(text[start:].strip())
    return [p for p in parts if p]


def _set(backend: demo.DemoBackend, text: str, depth: int = 0) -> list[str] | dict:
    """The unique names of a set's members, in order, or the answer saying why not."""
    text = text.strip()
    if depth > 8:
        return formulas.fail("bad-request", _SETS)
    if text.startswith("{") and text.endswith("}"):
        out: list[str] = []
        for part in _split(text[1:-1]):
            found = _set(backend, part, depth + 1)
            if isinstance(found, dict):
                return found
            out.extend(found)
        return out
    top = _TOP.match(text)
    if top:
        found = _set(backend, top.group(2), depth + 1)
        if isinstance(found, dict):
            return found
        measure = _MEASURE_NAMES.get(top.group(4).lower())
        if measure is None or any(formulas.hierarchy_key(n) == formulas.MEASURES for n in found):
            return formulas.fail("bad-request", _SETS)
        scored = [(name, _aggregate(backend, [_keys(backend)[name.lower()]], measure)) for name in found]
        if top.group(1).lower() == "topcount":
            scored.sort(key=lambda pair: float("-inf") if pair[1] is None else pair[1], reverse=True)
        else:
            scored = sorted((p for p in scored if p[1] is not None), key=lambda pair: pair[1])
        return [name for name, _ in scored[: int(top.group(3))]]
    suffix = _SUFFIX.match(text)
    if suffix:
        return _expand(backend, suffix.group(1).strip(), suffix.group(2).lower())
    if formulas.hierarchy_key(text) == formulas.MEASURES:
        measure = _MEASURE_NAMES.get(text.lower())
        return [measure] if measure else _missing(text)
    key = _keys(backend).get(text.lower())
    return [backend._name_of[key]] if key else formulas.fail("bad-request", _SETS)


def _expand(backend: demo.DemoBackend, base: str, what: str) -> list[str] | dict:
    if what == "children":
        key = _keys(backend).get(base.lower())
        if key is None:
            return formulas.fail("bad-request", _SETS)
        return [backend._name_of[(key[0], child)] for child in backend._children.get(key, [])]
    hierarchies = {name.lower(): h for name, h in backend._hierarchies.items()}
    h = hierarchies.get(base.lower())
    depth = None
    if h is None:
        # "<hierarchy>.[Level]": the members of one level.
        head = formulas.hierarchy_part(base)
        h = hierarchies.get(head.lower())
        level = base[len(head) :].lstrip(".").strip("[]").lower()
        names = [name.lower() for name, _attr in h.levels] if h else []
        if level not in names:
            return formulas.fail("bad-request", _SETS)
        depth = names.index(level) + 1
    walked = [key for key in backend._walk(h, ()) if depth is None or len(key[1]) == depth]
    # A hierarchy's Members include its All member, first, as in a cube.
    keys = ([(h.unique_name, ())] if depth is None else []) + walked
    return [backend._name_of[key] for key in keys]
