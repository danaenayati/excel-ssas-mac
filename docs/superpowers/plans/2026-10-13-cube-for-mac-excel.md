# Cube for Mac Excel Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A Mac Excel task pane plus a local helper that lets the user explore an on-prem SSAS cube PivotTable-style, with refreshable results in the sheet.

**Architecture:** An Office.js task pane (React/TypeScript) calls a local HTTPS helper on `https://localhost:47613`. The helper (Python) turns a pivot layout into MDX, talks to SSAS over the native TCP binding from [MS-SSAS] with NTLM sign-in, and returns a ready-to-write grid. A built-in Sample cube backend implements the same interface with no server.

**Tech Stack:** Python 3.12 (uv-managed), pyspnego, keyring, cryptography, pytest. TypeScript, React 18, Vite, Vitest, Office.js (ExcelApi ≤ 1.9). launchd, macOS Keychain.

Spec: `docs/superpowers/specs/2026-10-13-excel-ssas-mac-design.md`

---

## File map

```
helper/
  pyproject.toml
  src/cubebridge/
    __init__.py            version
    errors.py              CubeError(category, message, detail) + categories
    ssas/
      dime.py              DIME record encode/decode, negotiation check
      sealing.py           sealed-frame layout (uint16 dataSize, uint16 tokenSize, ciphertext, token)
      transport.py         socket channel, linear-time DIME message reader
      envelopes.py         Authenticate / Discover / Execute SOAP builders (read-only)
      xmla.py              fault, rowset and cellset parsers -> CellSet
      session.py           connect, authenticate, discover(), execute_mdx()
      resolve.py           "host", "host:port", "host\\instance" -> (host, port)
    pivot/
      names.py             unique-name validation, MDX identifier/string quoting
      model.py             Layout dataclasses + from_json validation; CellSet/AxisMember types
      mdx.py               build_query(cube, layout) -> str; members_query(...)
      formats.py           SSAS FORMAT_STRING -> Excel number format
      grid.py              build_grid(cellset, layout, stats) -> dict (the `Grid` JSON)
    backends/
      base.py              Backend protocol
      ssas_backend.py      metadata via MDSCHEMA_*, members and query via MDX
      demo_backend.py      in-memory Sample cube producing the same CellSet shape
    store.py               config.json (token, port, connections) + Keychain passwords
    api.py                 HTTPS server, auth checks, routes, static pane files
    certs.py               localhost CA + leaf generation
    probe.py               `python -m cubebridge.probe` connection diagnostics
    __main__.py            `python -m cubebridge` runs the server
  tests/                   one test file per module, plus fake_ssas.py (loopback server)
addin/
  package.json, vite.config.ts, tsconfig.json, index.html
  manifest/manifest.xml    Office add-in manifest (port 47613)
  public/assets/           icons
  src/types.ts             THE API CONTRACT (mirrors pivot/model.py and grid.py output)
  src/api.ts               typed fetch client with X-Cube-Token
  src/state/               layout reducer (pure, tested)
  src/sheet/layout.ts      Grid -> cell blocks (pure, tested)
  src/sheet/writer.ts      Office.js: write/clear/format a report
  src/sheet/reports.ts     workbook settings + named ranges
  src/components/          ConnectionForm, CubePicker, FieldList, Areas, MemberPicker, StatusBar
  src/App.tsx, src/main.tsx, src/styles.css
install.sh, uninstall.sh, cube-probe
README.md, THIRD_PARTY_NOTICES.md
```

## Contracts (source of truth)

`addin/src/types.ts` is written first and is authoritative for the JSON API. Python
produces and consumes exactly those shapes, camelCase keys included.

Routes (all under `/api`, header `X-Cube-Token`, JSON bodies):

| Method | Path | Body / query | Returns |
|---|---|---|---|
| GET | `/status` | | `Status` |
| GET | `/connections` | | `Connection[]` |
| POST | `/connections` | `NewConnection` | `Connection` (tested before saving) |
| DELETE | `/connections/{id}` | | `{ok: true}` |
| GET | `/connections/{id}/catalogs` | | `Catalog[]` |
| GET | `/connections/{id}/cubes` | `?catalog=` | `Cube[]` |
| GET | `/connections/{id}/fields` | `?catalog=&cube=` | `Fields` |
| POST | `/connections/{id}/members` | `MembersRequest` | `MembersResponse` |
| POST | `/connections/{id}/query` | `QueryRequest` | `Grid` |
| POST | `/connections/{id}/cancel` | | `{ok: true}` |

Errors: non-2xx with `{"error": ApiError}`. Status codes: 400 bad-request, 401 auth /
auth-mechanism, 403 missing token or wrong Host/Origin, 404 not-found, 413 too-large,
502 unreachable / instance / server, 499 cancelled, 500 internal.

The connection with id `demo` always exists (kind `demo`, the Sample cube).

Python backend interface (`backends/base.py`):

```python
class Backend(Protocol):
    def catalogs(self) -> list[dict]: ...                      # Catalog[]
    def cubes(self, catalog: str) -> list[dict]: ...           # Cube[]
    def fields(self, catalog: str, cube: str) -> dict: ...     # Fields
    def members(self, catalog: str, cube: str, hierarchy: str,
                parent: str | None, search: str | None, limit: int) -> dict: ...  # MembersResponse
    def query(self, catalog: str, cube: str, layout: Layout) -> tuple[CellSet, str]: ...  # result + MDX text
    def cancel(self) -> None: ...
    def close(self) -> None: ...
```

`CellSet` (`pivot/model.py`): `axes: list[list[tuple[AxisMember, ...]]]` (axis 0 =
columns, axis 1 = rows, may be absent), `cells: dict[int, Cell]` keyed by ordinal
`col + row * ncols`, where `AxisMember(unique_name, caption, hierarchy, level,
child_count)` and `Cell(value, format_string, formatted)`.

## MDX rules (`pivot/mdx.py`)

For an axis with hierarchies h1..hn, with `D(h)` the detail set and `A(h)` the All member:

- `D(h)` = `{A(h)}.Children` when the hierarchy has an All member, else
  `h.Levels(0).Members`; wrapped as `DrilldownMember(D, {expanded...}, RECURSIVE)` when
  any member is expanded.
- `T(n) = {D(hn), {A(hn)}}`, `T(k) = {CrossJoin(D(hk), T(k+1)), {(A(hk),…,A(hn))}}`.
  The All tuple at depth 1 is the grand total (included when `grandTotals`); at depth
  ≥ 2 it is a subtotal (included when `subtotals`). Any All tuple is dropped when one of
  its hierarchies has no All member.
- Columns: `CrossJoin(T(1), {measures})`, or `{measures}` with no column hierarchies.
  With no values chosen the measure set is `{[Measures].DefaultMember}`.
- Rows axis omitted when there are no row hierarchies.
- Any field with `members` selected adds a nested sub-select
  `FROM (SELECT ({m1,m2}) ON COLUMNS FROM …)`, except a Filters-area field with exactly
  one member, which goes in the `WHERE` tuple.
- `NON EMPTY` on both axes when `options.nonEmpty`.
- Tail: `CELL PROPERTIES VALUE, FORMAT_STRING, FORMATTED_VALUE`.
- Every identifier inserted into MDX must pass `names.validate_unique_name`
  (bracketed parts and `&[key]` parts only). The cube name is bracket-quoted with `]`
  doubled. Search text is a double-quoted MDX string with `"` doubled.

## Tasks

Each backend task is test-first: write the listed tests, see them fail, implement, see
them pass, commit. Run with `cd helper && uv run pytest -q`.

### Task 1: Contracts and scaffolding
- [ ] `addin/src/types.ts` with every type in the route table.
- [ ] `helper/pyproject.toml`, package skeleton, `errors.py`; `uv sync` works; `pytest` collects 0 tests.
- [ ] `THIRD_PARTY_NOTICES.md` crediting `oluies/ssas-xmla-tcp` (Apache-2.0).
- [ ] Commit.

### Task 2: `ssas.dime`
Tests (`tests/test_dime.py`): header bytes of a text/xml record equal the [MS-SSAS]
worked example prefix `0E 10 00 04 00 00 00 08`; fields padded to 4 bytes; round-trip;
chunked message (MB…ME) reassembles; wrong version raises `ProtocolError`;
`check_negotiated` rejects sx/xpress option bits and non-`text/xml` types.

### Task 3: `ssas.sealing`
Tests: frame = `<HH` sizes + ciphertext + token, ciphertext before token; BOM sealed as
its own first frame; payload chunked at 2888; unseal concatenates and strips BOM;
truncated header and truncated frame raise `IncompleteMessage`; padding mechanism refused.
Uses a fake context whose `wrap_winrm` returns `(b"T"*16, xor(data), 0)`.

### Task 4: `ssas.transport`
Tests: `MessageReader` returns whole messages when bytes arrive 1 at a time, when two
messages arrive in one chunk, and for a 5 MB chunked message without re-parsing from the
start (assert `recv` count is linear); EOF mid-message raises `CubeError(unreachable)`;
oversize declared length raises `ProtocolError`.

### Task 5: `ssas.envelopes` and `ssas.xmla`
Tests: Authenticate uses the `…/2003/ext` namespace; first request carries
`BeginSession`, later ones `Session SessionId`; restriction names validated, values
escaped; Execute carries `Catalog`, `Format=Multidimensional`,
`AxisFormat=TupleFormat`, `Content=Data`, `Timeout`; no builder for any other command.
`xmla.parse_rowset` on a canned MDSCHEMA response; `xmla.find_fault` on a SOAP fault;
`xmla.parse_cellset` on canned 2-axis, 1-axis, empty-axis and error-cell responses
(typed values: double, int, long, decimal, string, boolean; sparse ordinals).

### Task 6: `ssas.session` with loopback server
`tests/fake_ssas.py`: a threaded TCP server that speaks DIME, runs a real pyspnego NTLM
acceptor (`NTLM_USER_FILE`), requires sealed frames after the handshake, and answers
Discover/Execute from canned XML.
Tests: sign-in succeeds with the right password and `discover("DBSCHEMA_CATALOGS")`
returns rows; wrong password raises `CubeError(auth)`; closed port raises
`unreachable`; a 3 MB response round-trips; `SessionId` is reused; `cancel()` from a
second thread unblocks a slow query with `cancelled`; reconnect after server drop.

### Task 7: `ssas.resolve`
Tests: `"srv"` → `("srv", 2383, None)`; `"srv:2384"`; `"srv\\INST"` → instance lookup
called; `"tcp:srv,2384"`-style and junk rejected with `bad-request`. Instance lookup
sends an unauthenticated `DISCOVER_INSTANCES` to port 2382 and reads
`INSTANCE_PORT_NUMBER` (best effort; failure raises `CubeError(instance)`).

### Task 8: `pivot.names`, `pivot.model`, `pivot.mdx`
Tests: valid and invalid unique names (injection attempts such as
`[A].[B]} ON 0 FROM [X]; DROP`, unbalanced brackets, bare words); `Layout.from_json`
rejects unknown areas and invalid names; MDX golden strings for: measures only; one row
hierarchy; rows + columns + two measures; expansion; grand totals off; subtotals with
three row hierarchies; hierarchy without All member; single-member filter → WHERE;
multi-member filter → sub-select; axis field with member selection → sub-select;
`nonEmpty` off; `members_query` for top level, children and search.

### Task 9: `pivot.formats`
Tests: named formats (`Currency`, `Percent`, `Standard`, `Fixed`, `General Number`,
`Scientific`, dates), pass-through of custom strings, empty → `General`, strings Excel
cannot take (e.g. containing `\n`) → `General`.

### Task 10: `pivot.grid`
Tests on hand-built `CellSet`s: header text and indent for a multi-level hierarchy;
repeated outer labels blanked; subtotal and grand-total labelling and `rowKinds`;
column header rows with measures innermost; `filterRows` text (`All`, caption,
`(Multiple Items)`); format palette de-duplication; no rows → single `Total` row;
empty axis → `N == 0`; cell limit raises `too-large`.

### Task 11: `backends.demo_backend`
Tests: fields list has measures and four hierarchies; query totals equal the sum of
details for an additive measure and the ratio (not the sum) for `Avg Fare`; expansion
adds children directly under their parent; filter reduces totals; `members` search.

### Task 12: `backends.ssas_backend`
Tests against `fake_ssas.py` with canned MDSCHEMA and cellset XML: `cubes` hides
`$`-prefixed dimension cubes; `fields` groups hierarchies by dimension, drops hidden
ones and `[Measures]`, maps `HIERARCHY_ORIGIN` to kind and empty `ALL_MEMBER` to null;
`members` parses axis members; `query` sends the generated MDX and returns a `CellSet`.

### Task 13: `store`
Tests with a temp dir and a fake keyring: token created once; connections CRUD;
password never written to `config.json`; file mode 0600.

### Task 14: `api`
Tests with the demo backend over a real loopback HTTPS socket (test cert): 403 without
token, with wrong `Host`, with foreign `Origin`; every route's happy path; error body
shape and status code per category; static file serving with token injection; path
traversal rejected.

### Task 15: Task pane (`addin/`)
Built against `src/types.ts`. See "Task pane brief" below. Tests: Vitest for
`state/` reducer and `sheet/layout.ts`. `npm run build` produces `addin/dist`.

### Task 16: Install, probe, docs
- [ ] `certs.py` + test (leaf has SANs `localhost`, `127.0.0.1`; ≤ 825 days; CA key not kept).
- [ ] `probe.py`: prompts for server, username, password (`getpass`); prints DNS, port,
      sign-in, server version, catalogs, cubes, and a one-cell query; never prints the password.
- [ ] `install.sh` / `uninstall.sh` / `cube-probe`; `addin/manifest/manifest.xml`; icons.
- [ ] `README.md` for the user (install, first run, limits, troubleshooting).

### Task 17: End-to-end verification
- [ ] `uv run pytest -q` and `npm test` green.
- [ ] Helper running; `curl --cacert` hits `/api/status` and a Sample cube query.
- [ ] Pane loaded in a browser with an Office.js stub: connect to Sample cube, build a view, expand, filter.
- [ ] Manifest validated; sideload file in place.
- [ ] Record what could not be verified without the user (certificate trust prompt, Excel itself, the real server).

## Task pane brief (Task 15)

Look: Excel for Mac's own task panes. System font, 12–13 px, compact rows, light and
dark via `prefers-color-scheme`. No UI framework beyond React.

Screens:

1. **Connections**: list (Sample cube always present), "Add connection" form (name,
   server, username, password), delete. Errors shown inline with `ApiError.message`.
2. **Explorer** (after a connection is chosen): database and cube selects; search box;
   field list (Measures grouped by `group` then `folder`; dimensions → hierarchies) with
   checkboxes; four areas (Filters, Columns, Rows, Values) holding chips; options
   (Grand totals, Subtotals, Hide empty, Defer updates); status bar with row/column
   counts, elapsed time, "Show query", Refresh and Refresh All.

Behaviour:

- Ticking a measure adds it to Values; ticking a hierarchy adds it to Rows. Each chip
  has a menu: move to another area, move up/down, filter members, remove. HTML5
  drag-and-drop between areas is an enhancement on top of the menu, never the only way.
- A hierarchy can be in only one area. Moving it keeps `expanded` and `members`.
- Member picker: lazy tree (`parent`), search, select all/none, shows `truncated`.
- Every layout change re-queries after a 400 ms debounce unless "Defer updates" is on.
  A newer request supersedes an older one; stale responses are dropped.
- Selecting a header cell inside a report shows Expand/Collapse for its member when
  `hasChildren`. Selecting any cell inside a report makes that report the active one.
- "Insert here" creates a report anchored at the active cell.

Sheet writing (`sheet/layout.ts` pure, `sheet/writer.ts` Office.js):

- Block = filter rows, blank row (only if there are filter rows), column header rows,
  then data rows. Width = `rowFieldCaptions.length + C`. The last header row carries
  the row field captions in its left cells.
- Before writing, clear the report's previous extent. If the new extent would cover
  non-empty cells outside the previous extent, ask for confirmation in the pane.
- Values and number formats are written as 2-D arrays in chunks of ≤ 20,000 cells.
  Header and total rows bold; header fill; indent via `format.indentLevel` applied to
  batched `RangeAreas`.
- Report definition stored in `workbook.settings` under `cube.report.<id>`; location
  tracked by the workbook name `CubeReport_<id>`.
- Use only ExcelApi ≤ 1.9 and guard optional calls with
  `Office.context.requirements.isSetSupported`.
