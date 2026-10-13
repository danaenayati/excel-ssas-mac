"""What the API needs from a cube, whichever kind it is.

The JSON-shaped return values match addin/src/types.ts (Catalog, Cube, Fields,
MembersResponse). `query` returns a CellSet, which pivot/grid.py turns into a Grid.
"""

from __future__ import annotations

from typing import Protocol

from ..pivot.model import CellSet, Layout

DEFAULT_MEMBER_LIMIT = 500
MAX_MEMBER_LIMIT = 5000


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
    ) -> dict: ...

    def query(self, catalog: str, cube: str, layout: Layout) -> tuple[CellSet, str]:
        """Return the result and the query text that produced it."""
        ...

    def cancel(self) -> None: ...

    def close(self) -> None: ...
