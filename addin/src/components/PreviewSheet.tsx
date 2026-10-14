// Draws the PreviewHost's sheet next to the pane when the page is opened in an
// ordinary browser. Clicking a cell selects it and counts as a click, as in Excel; the
// arrow keys move the selection without clicking (click the grid first to give it the
// keyboard). The toolbar simulates the few other things a user does to a sheet that
// the pane has to cope with: typing into a cell, inserting rows, deleting the sheet,
// starting a new workbook.

import { memo, useState, useSyncExternalStore, type KeyboardEvent } from "react";
import { cellA1, columnLetters, containsCell } from "../sheet/geometry";
import { isFilterFormat } from "../sheet/layout";
import { formatForPreview, type PreviewCell, type PreviewHost, type PreviewSnapshot } from "../sheet/preview";
import { cx } from "../util";

const MIN_ROWS = 40;
const MIN_COLS = 12;
const MAX_ROWS = 400;
const MAX_COLS = 40;

interface RowProps {
  row: number;
  cols: number;
  /** Changes whenever any cell or name changes. */
  version: number;
  snapshot: PreviewSnapshot;
  /** Selected column in this row, or -1. */
  selected: number;
  onSelect: (row: number, col: number) => void;
  /** Rows frozen at the top of this sheet. */
  frozen: number;
}

const Row = memo(
  function Row({ row, cols, snapshot, selected, onSelect, frozen }: RowProps) {
    const cells: JSX.Element[] = [];
    for (let col = 0; col < cols; col++) {
      const cell: PreviewCell | undefined = snapshot.cells.get(row + "," + col);
      let inReport = false;
      snapshot.names.forEach((rect) => {
        if (containsCell(rect, row, col)) inReport = true;
      });
      const numeric = cell !== undefined && typeof cell.value === "number";
      cells.push(
        <td
          key={col}
          className={cx(
            "pv-cell",
            cell && cell.bold && "bold",
            cell && cell.fill && "fill",
            cell && cell.muted && "muted",
            cell && isFilterFormat(cell.format) && "filter",
            row < frozen && "frozen",
            numeric && "num",
            inReport && "in-report",
            selected === col && "selected",
          )}
          style={cell && cell.indent ? { paddingLeft: 4 + cell.indent * 10 } : undefined}
          onMouseDown={() => onSelect(row, col)}
        >
          {cell ? formatForPreview(cell.value, cell.format) : ""}
        </td>,
      );
    }
    return (
      <tr>
        <th className="pv-head">{row + 1}</th>
        {cells}
      </tr>
    );
  },
  (a, b) =>
    a.version === b.version &&
    a.selected === b.selected &&
    a.cols === b.cols &&
    a.row === b.row &&
    a.frozen === b.frozen,
);

export function PreviewSheet({ host }: { host: PreviewHost }) {
  const snapshot = useSyncExternalStore(
    (notify) => host.subscribe(notify),
    () => host.snapshot(),
  );
  const [typed, setTyped] = useState("");

  const rows = Math.min(MAX_ROWS, Math.max(MIN_ROWS, snapshot.usedRows + 10, snapshot.selection.row + 5));
  const cols = Math.min(MAX_COLS, Math.max(MIN_COLS, snapshot.usedCols + 3, snapshot.selection.col + 2));
  const clipped = snapshot.usedRows > MAX_ROWS || snapshot.usedCols > MAX_COLS;
  const { row, col } = snapshot.selection;
  const click = (r: number, c: number) => host.click(r, c);
  const frozen = host.frozenRows();
  const keys: Record<string, [number, number]> = {
    ArrowUp: [-1, 0],
    ArrowDown: [1, 0],
    ArrowLeft: [0, -1],
    ArrowRight: [0, 1],
  };
  const onKeyDown = (event: KeyboardEvent) => {
    const step = keys[event.key];
    if (!step) return;
    event.preventDefault();
    host.select(Math.max(0, row + step[0]), Math.max(0, col + step[1]));
  };

  const header: JSX.Element[] = [];
  for (let c = 0; c < cols; c++) {
    header.push(
      <th key={c} className="pv-head">
        {columnLetters(c)}
      </th>,
    );
  }
  const body: JSX.Element[] = [];
  for (let r = 0; r < rows; r++) {
    body.push(
      <Row
        key={r}
        row={r}
        cols={cols}
        version={snapshot.contentVersion}
        snapshot={snapshot}
        selected={r === row ? col : -1}
        onSelect={click}
        frozen={frozen}
      />,
    );
  }

  return (
    <div className="preview">
      <div className="pv-toolbar">
        <strong>Preview: {snapshot.sheetName}</strong>
        <span className="pv-name" data-testid="pv-selection">
          {cellA1(row, col)}
        </span>
        <form
          className="pv-form"
          onSubmit={(event) => {
            event.preventDefault();
            const number = typed.trim() !== "" && !isNaN(Number(typed)) ? Number(typed) : null;
            host.type(row, col, number !== null ? number : typed);
            setTyped("");
          }}
        >
          <input
            className="input"
            value={typed}
            onChange={(event) => setTyped(event.target.value)}
            placeholder="Type into the selected cell"
            aria-label="Value for the selected cell"
          />
          <button type="submit" className="button small">
            Set
          </button>
        </form>
        <button type="button" className="button small" onClick={() => host.insertRows(row, 1)}>
          Insert row above
        </button>
        <button
          type="button"
          className="button small"
          title="The reports stay in the workbook but lose their cells, as when a sheet is deleted in Excel"
          onClick={() => host.deleteSheet()}
        >
          Delete sheet
        </button>
        <button
          type="button"
          className="button small"
          onClick={() => {
            // The pane only reads the workbook's reports when it opens, as in Excel,
            // so a new workbook means opening the pane again.
            host.reset();
            window.location.reload();
          }}
        >
          New workbook
        </button>
        <span className="pv-note">
          Not Excel: a stand-in for testing. Reloading the page is like reopening the workbook.
        </span>
      </div>
      <div className="pv-scroll" tabIndex={0} onKeyDown={onKeyDown} aria-label="Preview sheet">
        <table className="pv-table">
          <thead>
            <tr>
              <th className="pv-head pv-corner" />
              {header}
            </tr>
          </thead>
          <tbody>{body}</tbody>
        </table>
        {clipped && (
          <p className="pv-note">
            The preview draws the first {MAX_ROWS} rows and {MAX_COLS} columns only. The block itself is
            complete.
          </p>
        )}
      </div>
      <div className="pv-tabs" role="tablist" aria-label="Sheets">
        {snapshot.sheets.map((sheet) => (
          <button
            key={sheet.id}
            type="button"
            role="tab"
            aria-selected={sheet.id === snapshot.activeId}
            className={cx("pv-tab", sheet.id === snapshot.activeId && "active")}
            onClick={() => host.activate(sheet.id)}
          >
            {sheet.name}
          </button>
        ))}
      </div>
    </div>
  );
}
