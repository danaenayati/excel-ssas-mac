"""Cube formulas in the sheet: CUBE.VALUE, CUBE.MEMBER, CUBE.MEMBERPROPERTY,
CUBE.RANKEDMEMBER and CUBE.SETCOUNT, answered in batches.

Excel calls a custom function once per cell. The functions runtime collects the calls
of a few milliseconds and sends them here in one request per cube, and every item of
the request gets its own answer, `{"value": …}` or `{"error": {category, message}}`,
so that one bad member costs only its own cell (FormulaResult in addin/src/types.ts).

How a batch is asked:

- CUBE.VALUE tuples are grouped by the hierarchies they name, because all tuples of one
  MDX set must name the same hierarchies in the same order. Each group is asked as
  `SELECT {(…), (…)} ON 0 FROM [cube] CELL PROPERTIES VALUE`. A tuple without a measure
  gets the cube's default measure, as in Excel.
- Member captions come from `SELECT {} ON 0, {members} ON 1`, one query per hierarchy;
  member properties from a calculated measure over the same rows; set counts from
  calculated measures `Count({set})`; ranked members from `Head({set}, n)`.
- When the cube rejects a query, or answers it with fewer positions than it was asked
  for (a cube set to ignore missing members drops them silently), the group is split in
  halves and asked again, down to single items, so that only the bad items fail. Every
  query of a request counts against MAX_QUERIES, which bounds what a batch full of bad
  items can cost.
- Only "the cube rejected the request" is split. Anything else (the server is gone, the
  sign-in failed, no access to the cube) is the same for every item, so the request
  fails as a whole and every cell says why.

Nothing is cached: a recalculation in Excel asks the cube again.

Everything typed by the user is checked by pivot/names.py first: member names must be
unique names, set expressions pass `validate_expression` and are wrapped in braces
where they are placed, and property names become MDX string literals.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Hashable, Sequence
from dataclasses import dataclass

from .errors import CubeError
from .pivot.grid import clean_value
from .pivot.model import CellSet
from .pivot.names import quote_identifier, quote_string, validate_expression, validate_unique_name

MAX_ITEMS = 5000  # tuples, or lookups, in one request
MAX_TUPLE_MEMBERS = 64
MAX_PROPERTY = 200
MAX_RANK = 10_000
MAX_QUERIES = 200  # queries one request may send, splitting included
MAX_PER_QUERY = 500  # items placed in one query
MAX_QUERY_TEXT = 200_000  # characters of names placed in one query

DEFAULT_MEASURE = "[Measures].DefaultMember"
MEASURES = "[measures]"
# What ssas/xmla.py reads from a cell that failed to calculate.
CELL_ERROR = "#VALUE!"
# The message ssas/xmla.py gives a query the cube refused (bad syntax, unknown member).
# Access refusals and transport failures have other messages and are never split.
REJECTED = "The cube rejected the request."

LOOKUP_KINDS = frozenset({"member", "property", "count", "ranked"})

_PART = r"\[(?:[^\]]|\]\])*\]"
_HEAD = re.compile(rf"({_PART})(?:\.({_PART}))?")
_PROPERTY_MEASURE = "[Measures].[__CubeExplorer Property]"
_TOO_MANY = (
    "Too many cube formulas failed at once to check them one by one. "
    "Correct the ones that fail, then recalculate."
)


# -- answers -----------------------------------------------------------------------
def ok(value: object) -> dict:
    return {"value": clean_value(value)}


def fail(category: str, message: str, detail: str | None = None) -> dict:
    error = {"category": category, "message": message}
    if detail:
        error["detail"] = detail[:500]
    return {"error": error}


def is_failure(answer: object) -> bool:
    return isinstance(answer, dict) and "error" in answer


# -- names -------------------------------------------------------------------------
def hierarchy_key(name: str) -> str:
    """The hierarchy a unique name belongs to, lower-cased; "[measures]" for any measure.

    "[Date].[Calendar].[Year].&[2025]" and "[Date].[Calendar].&[2025]" are both in
    "[date].[calendar]"; "[Date].&[2025]" names a member of the dimension's only
    hierarchy, "[date]".
    """
    match = _HEAD.match(name)
    if not match:
        return name.lower()
    first, second = match.group(1), match.group(2)
    if first.lower() == MEASURES:
        return MEASURES
    return (first + "." + second).lower() if second else first.lower()


def hierarchy_part(name: str) -> str:
    """The hierarchy part of a unique name as written, for use in MDX."""
    match = _HEAD.match(name)
    if not match:
        return name
    first, second = match.group(1), match.group(2)
    return first + "." + second if second and first.lower() != MEASURES else first


def _member(text: object) -> str | dict:
    """A unique name, or an answer saying why the text is not one."""
    try:
        return validate_unique_name(text.strip() if isinstance(text, str) else text)
    except CubeError:
        shown = text if isinstance(text, str) else ""
        return fail(
            "bad-request",
            f'"{shown[:80]}" is not a member\'s unique name. Write names the way the cube does, '
            "such as [Date].[Year].&[2025] or [Measures].[Sales].",
        )


def _one_per_hierarchy(names: Sequence[str]) -> dict | None:
    seen: dict[str, str] = {}
    for name in names:
        key = hierarchy_key(name)
        if key in seen and seen[key].lower() != name.lower():
            return fail(
                "bad-request",
                f"{seen[key]} and {name} are items of the same field. One formula can take one item of each field.",
            )
        seen[key] = name
    return None


def prepare_tuple(raw: Sequence[object]) -> tuple[str, ...] | dict:
    """A CUBE.VALUE tuple ready for MDX: valid names, at most one per hierarchy, in a
    fixed order (by hierarchy), duplicates dropped. Or the answer saying why not."""
    if len(raw) > MAX_TUPLE_MEMBERS:
        return fail("bad-request", f"One formula can name at most {MAX_TUPLE_MEMBERS} members.")
    names: list[str] = []
    for text in raw:
        name = _member(text)
        if isinstance(name, dict):
            return name
        names.append(name)
    clash = _one_per_hierarchy(names)
    if clash:
        return clash
    unique: dict[str, str] = {}
    for name in names:
        unique.setdefault(hierarchy_key(name), name)
    return tuple(unique[key] for key in sorted(unique))


@dataclass(frozen=True)
class Lookup:
    kind: str  # "member" | "property" | "count" | "ranked"
    members: tuple[str, ...] = ()  # member: the tuple; property: the one member
    prop: str = ""
    set: str = ""
    rank: int = 0


def prepare_lookup(item: dict) -> Lookup | dict:
    """Check one lookup (already of the right shape, see `read_lookups`)."""
    kind = item["kind"]
    if kind in ("member", "property"):
        raw = item["members"] if kind == "member" else [item["member"]]
        if not raw:
            return fail("bad-request", "The formula names no member.")
        if len(raw) > MAX_TUPLE_MEMBERS:
            return fail("bad-request", f"One formula can name at most {MAX_TUPLE_MEMBERS} members.")
        names = []
        for text in raw:
            name = _member(text)
            if isinstance(name, dict):
                return name
            names.append(name)
        clash = _one_per_hierarchy(names)
        if clash:
            return clash
        if kind == "member":
            return Lookup("member", members=tuple(names))
        prop = item["property"].strip()
        if not prop or len(prop) > MAX_PROPERTY or any(ord(ch) < 0x20 for ch in prop):
            return fail("bad-request", f"The property name must be 1 to {MAX_PROPERTY} characters, such as Key or Caption.")
        if hierarchy_key(names[0]) == MEASURES:
            return fail("bad-request", "Measures have no member properties.")
        return Lookup("property", members=(names[0],), prop=prop)
    try:
        expression = validate_expression(item["set"], "The set expression")
    except CubeError as exc:
        return fail("bad-request", exc.message)
    if kind == "count":
        return Lookup("count", set=expression)
    rank = item["rank"]
    if not 1 <= rank <= MAX_RANK:
        return fail("bad-request", f"The rank must be a whole number from 1 to {MAX_RANK:,}.")
    return Lookup("ranked", set=expression, rank=rank)


# -- request shapes ----------------------------------------------------------------
def _bad(message: str) -> CubeError:
    return CubeError("bad-request", message)


def read_tuples(raw: object) -> list[list[str]]:
    """The `tuples` of a cells request: a list of lists of text. Wrong shapes are a
    mistake of the caller and fail the request; the names are checked per item."""
    if not isinstance(raw, list):
        raise _bad("'tuples' must be a list.")
    if len(raw) > MAX_ITEMS:
        raise _bad(f"One request can ask for at most {MAX_ITEMS:,} values.")
    out = []
    for item in raw:
        if not isinstance(item, list) or not all(isinstance(name, str) for name in item):
            raise _bad("Each tuple must be a list of member names.")
        out.append(item)
    return out


def read_lookups(raw: object) -> list[dict]:
    """The `items` of a lookup request, each in its kind's shape."""
    if not isinstance(raw, list):
        raise _bad("'items' must be a list.")
    if len(raw) > MAX_ITEMS:
        raise _bad(f"One request can ask for at most {MAX_ITEMS:,} items.")
    out = []
    for item in raw:
        kind = item.get("kind") if isinstance(item, dict) else None
        if kind not in LOOKUP_KINDS:
            raise _bad("Each item needs a kind: member, property, count or ranked.")
        if kind == "member":
            members = item.get("members")
            if not isinstance(members, list) or not all(isinstance(m, str) for m in members):
                raise _bad("A member item needs 'members', a list of member names.")
            out.append({"kind": kind, "members": members})
        elif kind == "property":
            if not isinstance(item.get("member"), str) or not isinstance(item.get("property"), str):
                raise _bad("A property item needs 'member' and 'property' as text.")
            out.append({"kind": kind, "member": item["member"], "property": item["property"]})
        else:
            if not isinstance(item.get("set"), str):
                raise _bad("A set item needs 'set' as text.")
            entry = {"kind": kind, "set": item["set"]}
            if kind == "ranked":
                rank = item.get("rank")
                if isinstance(rank, bool) or not isinstance(rank, int):
                    raise _bad("A ranked item needs 'rank' as a whole number.")
                entry["rank"] = rank
            out.append(entry)
    return out


# -- asking the cube ---------------------------------------------------------------
Run = Callable[[str], CellSet]


def _cell_answer(cellset: CellSet, ordinal: int, failed: str) -> dict:
    cell = cellset.cells.get(ordinal)
    if cell is None:
        return ok(None)
    if cell.value == CELL_ERROR:
        return fail("not-found", failed)
    return ok(cell.value)


def _rows(cellset: CellSet) -> list:
    return cellset.axes[1] if len(cellset.axes) > 1 else []


def _columns(cellset: CellSet) -> list:
    return cellset.axes[0] if cellset.axes else []


class _Asker:
    """Sends the queries of one request and keeps count of them."""

    def __init__(self, cube: str, run: Run) -> None:
        self.cube = quote_identifier(cube)
        self.run = run
        self.queries = 0

    def settle(
        self,
        items: list[Hashable],
        query: Callable[[list], str],
        read: Callable[[CellSet, list], list | None],
        weight: Callable[[Hashable], int],
        missing: str,
    ) -> dict:
        """An answer for every item. `read` returns one answer per item, in order, or
        None when the result does not line up with what was asked."""
        answers: dict = {}
        pending = list(reversed(_chunks(items, weight)))
        while pending:
            chunk = pending.pop()
            if self.queries >= MAX_QUERIES:
                for item in chunk:
                    answers[item] = fail("not-found", _TOO_MANY)
                continue
            self.queries += 1
            rejected: CubeError | None = None
            try:
                found = read(self.run(query(chunk)), chunk)
            except CubeError as exc:
                if exc.category != "server" or exc.message != REJECTED:
                    raise
                found, rejected = None, exc
            if found is not None:
                answers.update(zip(chunk, found))
            elif len(chunk) > 1:
                half = len(chunk) // 2
                pending.append(chunk[half:])
                pending.append(chunk[:half])
            elif rejected is not None:
                answers[chunk[0]] = fail("not-found", "The cube could not answer this formula.", rejected.detail)
            else:
                answers[chunk[0]] = fail("not-found", missing)
        return answers


def _chunks(items: list, weight: Callable[[Hashable], int]) -> list[list]:
    out: list[list] = []
    current: list = []
    size = 0
    for item in items:
        cost = weight(item)
        if current and (len(current) >= MAX_PER_QUERY or size + cost > MAX_QUERY_TEXT):
            out.append(current)
            current, size = [], 0
        current.append(item)
        size += cost
    if current:
        out.append(current)
    return out


def evaluate_cells(cube: str, tuples: list[list[str]], run: Run) -> list[dict]:
    """CUBE.VALUE: one answer per tuple."""
    answers: list[dict | None] = [None] * len(tuples)
    canonical: dict[tuple[str, ...], tuple[str, ...]] = {}
    wanted: dict[tuple[str, ...], list[int]] = {}
    for index, raw in enumerate(tuples):
        prepared = prepare_tuple(raw)
        if isinstance(prepared, dict):
            answers[index] = prepared
            continue
        key = tuple(name.lower() for name in prepared)
        canonical.setdefault(key, prepared)
        wanted.setdefault(key, []).append(index)

    groups: dict[tuple[str, ...], list[tuple[str, ...]]] = {}
    for key, names in canonical.items():
        groups.setdefault(tuple(hierarchy_key(n) for n in names), []).append(key)

    asker = _Asker(cube, run)

    def query(chunk: list) -> str:
        sets = ", ".join("(" + ", ".join(canonical[key] or (DEFAULT_MEASURE,)) + ")" for key in chunk)
        return f"SELECT {{{sets}}} ON 0 FROM {asker.cube} CELL PROPERTIES VALUE"

    def read(cellset: CellSet, chunk: list) -> list | None:
        if len(_columns(cellset)) != len(chunk):
            return None
        return [_cell_answer(cellset, i, "The cube could not calculate this value.") for i in range(len(chunk))]

    for keys in groups.values():
        found = asker.settle(
            keys,
            query,
            read,
            lambda key: sum(len(n) + 2 for n in canonical[key]) + 4,
            "The cube has no such member, or no value for this combination of members.",
        )
        for key, answer in found.items():
            for index in wanted[key]:
                answers[index] = answer
    return answers  # type: ignore[return-value]


def evaluate_lookups(cube: str, items: list[dict], run: Run) -> list[dict]:
    """CUBE.MEMBER, CUBE.MEMBERPROPERTY, CUBE.SETCOUNT, CUBE.RANKEDMEMBER: one answer per item."""
    answers: list[dict | None] = [None] * len(items)
    prepared: list[Lookup | None] = [None] * len(items)
    for index, item in enumerate(items):
        lookup = prepare_lookup(item)
        if isinstance(lookup, dict):
            answers[index] = lookup
        else:
            prepared[index] = lookup
    asker = _Asker(cube, run)

    # Captions of every member named by a member item, one query per hierarchy.
    members: dict[str, str] = {}
    for lookup in prepared:
        if lookup and lookup.kind == "member":
            for name in lookup.members:
                members.setdefault(name.lower(), name)
    captions = _captions(asker, members)
    for index, lookup in enumerate(prepared):
        if lookup and lookup.kind == "member":
            parts = [captions[name.lower()] for name in lookup.members]
            failed = next((p for p in parts if is_failure(p)), None)
            # A tuple shows as its last member, as Excel's CUBEMEMBER does.
            answers[index] = failed or parts[-1]

    properties = _properties(asker, [p for p in prepared if p and p.kind == "property"])
    counts = _counts(asker, [p for p in prepared if p and p.kind == "count"])
    ranked = _ranked(asker, [p for p in prepared if p and p.kind == "ranked"])
    for index, lookup in enumerate(prepared):
        if lookup is None or lookup.kind == "member":
            continue
        if lookup.kind == "property":
            answers[index] = properties[(lookup.members[0].lower(), lookup.prop.lower())]
        elif lookup.kind == "count":
            answers[index] = counts[lookup.set]
        else:
            found = ranked[lookup.set]
            if is_failure(found):
                answers[index] = found
            elif lookup.rank > len(found["captions"]):
                answers[index] = fail("not-found", f"The set has {len(found['captions']):,} items, fewer than {lookup.rank:,}.")
            else:
                answers[index] = ok(found["captions"][lookup.rank - 1])
    return answers  # type: ignore[return-value]


def _captions(asker: _Asker, members: dict[str, str]) -> dict[str, dict]:
    by_hierarchy: dict[str, list[str]] = {}
    for key, name in members.items():
        by_hierarchy.setdefault(hierarchy_key(name), []).append(key)
    out: dict[str, dict] = {}

    def query(chunk: list) -> str:
        return f"SELECT {{}} ON 0, {{{', '.join(members[k] for k in chunk)}}} ON 1 FROM {asker.cube}"

    def read(cellset: CellSet, chunk: list) -> list | None:
        rows = _rows(cellset)
        if len(rows) != len(chunk) or not all(rows):
            return None
        return [ok(row[-1].caption) for row in rows]

    for keys in by_hierarchy.values():
        out.update(asker.settle(keys, query, read, lambda k: len(members[k]) + 2, "The cube has no such member."))
    return out


def _properties(asker: _Asker, lookups: list[Lookup]) -> dict[tuple[str, str], dict]:
    groups: dict[tuple[str, str], dict[str, str]] = {}
    props: dict[tuple[str, str], str] = {}
    for lookup in lookups:
        name = lookup.members[0]
        group = (hierarchy_key(name), lookup.prop.lower())
        groups.setdefault(group, {}).setdefault(name.lower(), name)
        props.setdefault(group, lookup.prop)
    out: dict[tuple[str, str], dict] = {}
    for group, names in groups.items():
        prop = props[group]
        hierarchy = hierarchy_part(next(iter(names.values())))

        def query(chunk: list, names=names, prop=prop, hierarchy=hierarchy) -> str:
            return (
                f"WITH MEMBER {_PROPERTY_MEASURE} AS {hierarchy}.CurrentMember.Properties({quote_string(prop)}) "
                f"SELECT {{{_PROPERTY_MEASURE}}} ON 0, {{{', '.join(names[k] for k in chunk)}}} ON 1 "
                f"FROM {asker.cube} CELL PROPERTIES VALUE"
            )

        def read(cellset: CellSet, chunk: list, prop=prop) -> list | None:
            if len(_rows(cellset)) != len(chunk) or len(_columns(cellset)) != 1:
                return None
            return [_cell_answer(cellset, i, f"The member has no property called {prop}.") for i in range(len(chunk))]

        found = asker.settle(list(names), query, read, lambda k, names=names: len(names[k]) + 2, "The cube has no such member.")
        for key, answer in found.items():
            out[(key, prop.lower())] = answer
    return out


def _counts(asker: _Asker, lookups: list[Lookup]) -> dict[str, dict]:
    sets = list(dict.fromkeys(lookup.set for lookup in lookups))
    if not sets:
        return {}

    def query(chunk: list) -> str:
        names = [f"[Measures].[__CubeExplorer Count {i}]" for i in range(len(chunk))]
        members = " ".join(f"MEMBER {name} AS Count({{{expression}}})" for name, expression in zip(names, chunk))
        return f"WITH {members} SELECT {{{', '.join(names)}}} ON 0 FROM {asker.cube} CELL PROPERTIES VALUE"

    def read(cellset: CellSet, chunk: list) -> list | None:
        if len(_columns(cellset)) != len(chunk):
            return None
        return [_cell_answer(cellset, i, "The cube could not count this set.") for i in range(len(chunk))]

    return asker.settle(sets, query, read, lambda s: len(s) + 60, "The cube could not count this set.")


def _ranked(asker: _Asker, lookups: list[Lookup]) -> dict[str, dict]:
    deepest: dict[str, int] = {}
    for lookup in lookups:
        deepest[lookup.set] = max(deepest.get(lookup.set, 0), lookup.rank)
    out: dict[str, dict] = {}
    for expression, count in deepest.items():

        def query(_chunk: list, expression=expression, count=count) -> str:
            return f"SELECT {{}} ON 0, Head({{{expression}}}, {int(count)}) ON 1 FROM {asker.cube}"

        def read(cellset: CellSet, _chunk: list) -> list:
            # Not an answer yet: the captions in order, for each rank asked of this set.
            return [{"captions": [row[-1].caption if row else "" for row in _rows(cellset)]}]

        out.update(asker.settle([expression], query, read, lambda s: len(s), "The cube could not read this set."))
    return out
