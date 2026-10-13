# Cube Explorer for Excel on Mac

Explore a SQL Server Analysis Services (SSAS) cube from Excel for Mac: pick measures
and fields, filter, drill down, and get refreshable numbers in your sheet. Windows
Excel has this built in; Mac Excel does not, so this adds it.

It has two parts that install together:

- **A panel inside Excel** (an Office add-in) where you build the view.
- **A small helper** that runs quietly on your Mac and talks to the cube. You never
  open it. It only listens to Excel on this Mac.

Nothing is installed or changed on the company server.

## Install

Open Terminal and run:

```
cd ~/Desktop/DEV/excel-ssas-mac
./install.sh
```

macOS asks for your Mac password once. That lets Excel trust the helper. The
certificate involved works for `https://localhost` only.

Then:

1. Quit Excel completely (Cmd+Q) and open it again.
2. On the **Home** tab, click **Add-ins** (far right of the ribbon) and choose
   **Cube Explorer** from the list.
3. A **Cube Explorer** button then appears at the far right of the **Home** tab. Click
   it to open the panel.

Older versions of Excel have no Add-ins button on Home. There it is **Insert** → the
small arrow beside **My Add-ins** → **Cube Explorer**.

To update later, run `./install.sh` again.

## First use

**Try the Sample cube first.** It is built in and needs no server or VPN, so you can
see how everything works.

**Then test your real server.** Connect the company VPN and run:

```
./cube-probe
```

It asks for the server name (exactly what you type in Windows Excel), your username
as `DOMAIN\user`, and your password (hidden as you type). It checks each step and
tells you where it stops if something is wrong. When everything passes it offers to
save the connection so it appears in Excel.

You can also add a connection from the panel itself with **Add connection**.

## Using it

1. Choose a connection, then a database and a cube.
2. Tick measures (they go to **Values**) and fields (they go to **Rows**). Use a
   field's menu to move it to **Columns** or **Filters**, or to choose which members
   to include.
3. The result is written into the sheet where your cursor was. The cube calculates
   every number, including totals, so they match what Windows Excel shows.
4. **Filter from the table:** click any cell that shows a small ▾ (a page filter's
   value, a row field's name or a column field's name). A checklist opens in the
   panel; tick the items you want and press **Apply**.
5. Click a row or column label, then **Expand** or **Collapse** in the panel to drill.
6. **Refresh** updates the numbers. The view is saved inside the workbook, so you can
   reopen the file later and refresh it.
7. **New report here** starts another report at the selected cell. If a report's sheet
   was deleted, select a cell and press **Place here** to put it back, fields included.

Your password is kept in the Mac's Keychain. It is never written into a workbook.

## What it does not do

- It is not a native PivotTable. Slicers, PivotCharts and the PivotTable tab do not
  attach to it. Ordinary charts and formulas that point at the result cells work.
- PivotTables that colleagues built on Windows against the cube still do not refresh
  on Mac.
- It is read-only. It cannot change anything in the cube.
- You need the VPN or office network, the same as on Windows.
- A colleague opening your workbook on their Mac needs their own install and
  connection. On Windows the numbers are plain cells.

## If something goes wrong

| What you see | What to do |
|---|---|
| "The helper is not running." | In Terminal: `launchctl kickstart -k gui/$(id -u)/com.cubebridge.helper` |
| Panel is blank, or Excel says it cannot load the add-in | Run `./install.sh` again, then quit and reopen Excel. |
| "Can't reach the server" | Connect the VPN. Run `./cube-probe` to see which step fails. |
| "The server refused the sign-in" | Check the username is `DOMAIN\user` and the password is current. |
| Sign-in fails although the password is right | Your server may accept only Kerberos sign-in. That is not supported yet. Run `./cube-probe --verbose` and send its output (it shows step names and sizes, never the password). |
| "Couldn't find the port of instance …" | Ask a Windows colleague for the port and enter the server as `name:port`. |
| "This view would return … cells" | Add a filter or remove a field. The limit protects Excel from freezing. |
| macOS asks whether Python may use your keychain or local network | Click **Always Allow** / **Allow**. It is the helper. |
| Cube Explorer is not in the **Add-ins** list | Quit Excel fully (Cmd+Q), reopen, and click **Add-ins** on the Home tab again. |

The helper's log is at
`~/Library/Application Support/CubeBridge/logs/helper.log`. It never contains
passwords.

The localhost certificate lasts about two years. The panel warns you before it
expires; running `./install.sh` renews it.

## Uninstall

```
./uninstall.sh
```

Removes the helper, the add-in, saved connections and their Keychain passwords, and
the certificate trust. Your workbooks keep the numbers already in them.

## For developers

```
helper/   Python service: SSAS wire protocol, MDX generation, local HTTPS API
addin/    Office.js task pane (React + TypeScript + Vite)
docs/     Design and implementation plan
```

- Tests: `cd helper && uv run pytest` and `cd addin && npm test`.
- The JSON contract between the two halves is `addin/src/types.ts`.
- Try the panel in a browser without Excel: `cd addin && npm run dev`, then open
  `http://localhost:5173/?mock=1`.
- The helper talks to SSAS over its native TCP binding ([MS-SSAS]) with NTLM
  sign-in. The wire details Microsoft does not document come from the
  [ssas-xmla-tcp](https://github.com/oluies/ssas-xmla-tcp) project; see
  `THIRD_PARTY_NOTICES.md`.
- Verified against a loopback stand-in server and a built-in sample cube. Named
  instance lookup on port 2382 and Kerberos are the known unverified or missing parts.
