"""What the API needs from a cube, whichever kind it is.

The JSON-shaped return values match addin/src/types.ts (Catalog, Cube, Fields,
MembersResponse). `query` returns a CellSet, which pivot/grid.py turns into a Grid.
"""

from __future__ import annotations

import re
from typing import Protocol

from ..pivot.model import CellSet, Layout

DEFAULT_MEMBER_LIMIT = 500
MAX_MEMBER_LIMIT = 5000
DEFAULT_DETAIL_ROWS = 1000
MAX_DETAIL_ROWS = 100_000

_PART = re.compile(r"\[((?:[^\]]|\]\])*)\]")


def plain_column_names(names: list[str]) -> list[str]:
    """"[$Trips].[City]" becomes "City", or "Trips City" when another table has a City too."""
    split = []
    for name in names:
        parts = [p.replace("]]", "]") for p in _PART.findall(name)] or [name]
        split.append((parts[-2].lstrip("$") if len(parts) > 1 else "", parts[-1]))
    counts: dict[str, int] = {}
    for _, last in split:
        counts[last.lower()] = counts.get(last.lower(), 0) + 1
    return [f"{table} {last}".strip() if counts[last.lower()] > 1 else last for table, last in split]


class Backend(Protocol):
    def catalogs(self) -> list[dict]: ...

    def cubes(self, catalog: str) -> list[dict]: ...

    def fields(self, catalog: str, cube: str) -> dict: ...

    def members(
        self,
        catalog: str,
        cube: str,
        hierarchy: str,
        parent: str | None,
        search: str | None,
        limit: int,
        level: str | None = None,
    ) -> dict: ...

    def query(self, catalog: str, cube: str, layout: Layout) -> tuple[CellSet, str]:
        """Return the result and the query text that produced it."""
        ...

    def details(
        self, catalog: str, cube: str, layout: Layout, members: list[str], measure: str, max_rows: int
    ) -> dict:
        """The rows behind one value cell: {"columns", "rows", "truncated"} (DetailsResponse)."""
        ...

    def cancel(self) -> None: ...

    def close(self) -> None: ...

    # Cube formulas in the sheet (formulas.py). Both answer one item per item asked,
    # {"value": …} or {"error": {category, message}} (FormulaResult in types.ts).
    def cells(self, catalog: str, cube: str, tuples: list[list[str]]) -> list[dict]:
        """CUBE.VALUE: the value of each tuple of member unique names."""
        ...

    def lookup(self, catalog: str, cube: str, items: list[dict]) -> list[dict]:
        """CUBE.MEMBER, CUBE.MEMBERPROPERTY, CUBE.SETCOUNT, CUBE.RANKEDMEMBER (`formulas.read_lookups` shapes)."""
        ...
