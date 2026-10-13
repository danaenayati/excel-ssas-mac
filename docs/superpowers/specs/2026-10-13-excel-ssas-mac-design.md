# Cube for Mac Excel: design

Date: 2026-10-13. Status: approved by Dana ("do all project").

## Goal

Let a Mac Excel user explore an on-premises SQL Server Analysis Services (SSAS) cube
the way a Windows user does with a cube PivotTable: pick measures and dimensions from
a field list, filter, drill, and get refreshable numbers in the sheet.

Constraints set by the user:

- Server is reached by company server name with Windows sign-in (native TCP, port 2383).
- No help from IT: nothing may be changed or added on the company side, no Windows machine.
- The user is not a developer. One install, then everything happens inside Excel.
- Built for Dana's Mac first. Distribution to colleagues is out of scope.

## What is impossible, and what replaces it

Mac Excel has no OLAP provider and no extension point for one, so a native cube
PivotTable cannot be enabled. Office add-ins on Mac run in a web view that can only make
HTTPS requests, so an add-in alone cannot open the TCP connection SSAS needs. Microsoft's
own ADOMD.NET supports TCP on Windows only.

Replacement: an Office add-in (task pane) plus a local helper service that speaks the
SSAS TCP binding from Microsoft's open specification [MS-SSAS].

## Architecture

```
Excel for Mac
  └─ task pane (Office.js, TypeScript/React)      addin/
        │  HTTPS, same origin, https://localhost:47613
        ▼
  helper service (Python, launchd agent)          helper/cubebridge/
        ├─ api        local JSON API + serves the built task pane
        ├─ pivot      layout → MDX; cellset → grid model
        ├─ backends   SsasBackend | DemoBackend (same interface)
        ├─ ssas       DIME framing, NTLM sign-in, sealing, XMLA envelopes, parsers
        └─ store      connections (JSON file) + passwords (macOS Keychain)
        │  TCP 2383, [MS-SSAS] native binding
        ▼
  SSAS server (unchanged)
```

Each unit has one job and a narrow interface:

| Unit | Does | Depends on |
|---|---|---|
| `ssas.dime` | Encode/decode DIME records | nothing |
| `ssas.sealing` | Post-handshake frame: `uint16 dataSize, uint16 tokenSize, ciphertext, token` | a security context |
| `ssas.session` | Connect, negotiate clear-text XML, authenticate, send/receive sealed XMLA | dime, sealing, pyspnego |
| `ssas.xmla` | Build Discover/Execute envelopes; parse rowsets, cellsets, faults | nothing |
| `ssas.resolve` | Parse `host`, `host:port`, `host\instance`; look up instance port on 2382 | session transport |
| `pivot.model` | Typed layout (rows, columns, values, filters, expanded) and grid result | nothing |
| `pivot.mdx` | Pure function: layout → MDX string | model |
| `pivot.grid` | Pure function: cellset → grid result (headers, cells, formats, member map) | model, xmla |
| `backends` | `catalogs()`, `cubes()`, `fields()`, `members()`, `query(layout)` | ssas + pivot, or demo data |
| `store` | Saved connections, Keychain passwords | keyring |
| `api` | HTTPS on 127.0.0.1, token + Host/Origin check, JSON routes, static files | backends, store |
| task pane | Field list, areas, filters, drill, write/refresh grid in sheet | api, Office.js |

The wire-level knowledge (namespace of `Authenticate`, BOM frame, sealed-frame layout,
NEGO bit, `BeginSession`) comes from the Apache-2.0 project `oluies/ssas-xmla-tcp`, which
verified it against SSAS 2022 over NTLM. Our `ssas` package is a derived work with
attribution in `THIRD_PARTY_NOTICES.md`. Added here: MDX execution with multidimensional
results, MDSCHEMA discovery, linear-time reads of large responses, instance resolution.

## User flow

1. `install.sh` once: creates the Python runtime, builds the pane, creates and trusts a
   localhost-only certificate (macOS asks for the login password once), registers the
   launchd agent, sideloads the add-in manifest.
2. Excel → Insert → My Add-ins → **Cube Explorer**; a ribbon button opens the pane.
3. First use: add a connection (server, username as `DOMAIN\user`, password). Password
   goes to the Keychain; the workbook never stores it.
4. Choose database and cube.
5. Tick or drag fields into Filters / Columns / Rows / Values.
6. The grid is written at the report's anchor cell. Every number is computed by the cube.
7. Select a row or column label → Expand / Collapse in the pane.
8. The report definition is saved in the workbook. Refresh re-runs it; Refresh All
   re-runs every report in the workbook.

A built-in **Sample cube** connection (DemoBackend) works with no server, so the whole
loop can be tried and tested off VPN.

## Query semantics

- MDX is used for both multidimensional and tabular models.
- Values: measures, placed on columns after any column hierarchies.
- Rows / Columns: hierarchies. Multi-level hierarchies start at the top level and are
  expanded with `DrilldownMember`. Several hierarchies on one axis are cross-joined.
- Totals: grand total per axis from the hierarchies' All members; optional subtotals.
  Because the cube computes them, ratios and distinct counts are correct.
- Filters: one selected member → `WHERE` tuple; several → sub-select (same split Excel
  uses on Windows).
- `NON EMPTY` on both axes. Cell properties `VALUE`, `FORMAT_STRING`, `FORMATTED_VALUE`.
- Guard: the helper refuses a result above a cell limit (default 300,000) with a clear
  message instead of freezing Excel.

## Sheet model

- A report is `{id, connectionId, catalog, cube, layout, options}` stored as JSON in the
  workbook settings.
- Its location is a workbook named range (`CubeReport_<id>`) so it survives inserted
  rows and columns. Refresh clears the previous extent, then writes the new one.
- Grid: header rows for column hierarchies and measures, one sheet column per row
  hierarchy, indent by level, bold totals, Excel number formats derived from the cube's
  format strings.
- The member behind each label is kept in memory for drill. After reopening a file,
  one Refresh restores it.

## Security

- Helper listens on 127.0.0.1 only. Requests must carry the install token and a
  `Host`/`Origin` of `localhost:47613`; anything else gets 403. This blocks other
  websites and DNS rebinding.
- Certificate: a private CA name-constrained to `localhost`/`127.0.0.1`; its key is
  deleted after signing the leaf, so it cannot sign anything else.
- Passwords live only in the macOS Keychain. They are never logged, never sent to the
  pane after entry, never written to the workbook.
- The SSAS layer is read-only by construction: it can build Discover and Execute
  requests only, and the only statements it executes are SELECTs produced by the MDX
  generator. Every name placed into MDX is validated as a bracketed identifier, so a
  crafted report definition in a workbook cannot inject other statements.

## Errors

Categorised in the helper and shown in plain words in the pane:

| Category | Message |
|---|---|
| unreachable | Can't reach the server. Are you on the VPN? |
| auth | The server refused the sign-in. Check username (DOMAIN\user) and password. |
| auth-mechanism | The server does not accept password (NTLM) sign-in. See README, "Kerberos". |
| instance | Couldn't find instance X. Enter the port as server:port. |
| too-large | This view would return N cells. Add a filter or remove a field. |
| server | The cube rejected the query: … (server text) |
| helper-down | The helper is not running. (pane shows the restart command) |

## Known limits (v1)

- Not a native PivotTable: no slicers, PivotCharts, or PivotTable ribbon.
- Existing Windows-built OLAP PivotTables are not refreshed or converted.
- NTLM sign-in only. Kerberos-only servers are not supported (frame layout for
  Kerberos is unverified upstream).
- Named-instance lookup on port 2382 is undocumented by Microsoft; implemented
  best-effort, with `server:port` as the reliable form.
- No sorting/top-N UI, KPIs, named sets, actions, drill-through, or writeback.
- Single Mac. The helper depends on `uv` and a user-level launchd agent.

## Testing

- Python unit tests: DIME, sealing, envelopes, rowset/cellset parsing, MDX generation,
  grid building, API auth checks.
- Loopback integration: a fake SSAS server in the test suite (DIME + real NTLM acceptor
  + sealing + canned XMLA) exercises the full client transport over a real socket.
- Frontend unit tests: sheet layout and number-format mapping.
- End-to-end without a server: Sample cube through the real API and pane.
- Against the real server: `cube-probe` CLI (reachability, sign-in, version, catalogs,
  cubes, one tiny query). This is the only test that proves the user's server works and
  must be run by the user, who types the password locally.
