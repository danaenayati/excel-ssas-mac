// PreviewHost: a stand-in sheet for when the pane is opened in an ordinary browser
// (Office.onReady resolves with host null). It keeps one sheet of cells, the report
// names and the "workbook settings" in memory, mirrors them to sessionStorage so that
// reloading the page behaves like reopening the workbook, and lets the page click a
// cell, move the selection with the keyboard, type into a cell, insert rows and delete
// the sheet. `components/PreviewSheet.tsx` draws it.
//
// It imitates the parts of Excel the pane has to cope with: a deleted sheet leaves the
// report definitions behind with no cells; a click is reported after the selection it
// causes; and moving with the keyboard selects without clicking.

import { containsCell, intersect, subtract, type Rect } from "./geometry";
import {
  SETTING_PREFIX,
  settingKey,
  type CellAddress,
  type Extent,
  type ReportDef,
  type SheetClick,
  type SheetHost,
  type SheetSelection,
} from "./host";
import type { Block, SheetValue } from "./layout";
import { parseReportDef, serializeReportDef } from "../state/report";

export interface PreviewCell {
  value: SheetValue;
  format: string;
  bold?: boolean;
  fill?: boolean;
  muted?: boolean;
  indent?: number;
}

export interface PreviewSnapshot {
  /** Goes up on every change, including a change of selection. */
  version: number;
  /** Goes up only when cells or report names change. */
  contentVersion: number;
  sheetName: string;
  cells: ReadonlyMap<string, PreviewCell>;
  /** Report id -> its block. */
  names: ReadonlyMap<string, Rect>;
  selection: { row: number; col: number };
  /** Bottom-right corner of everything on the sheet (exclusive). */
  usedRows: number;
  usedCols: number;
}

const STORAGE_KEY = "cube.preview.v1";
const STORAGE_LIMIT = 3000000;

function key(row: number, col: number): string {
  return row + "," + col;
}

function storage(): Storage | null {
  try {
    return typeof sessionStorage === "undefined" ? null : sessionStorage;
  } catch {
    return null;
  }
}

interface Stored {
  cells: [string, PreviewCell][];
  names: [string, Rect][];
  settings: [string, string][];
  selection: { row: number; col: number };
  sheet: number;
}

export class PreviewHost implements SheetHost {
  readonly kind = "preview" as const;

  private cells = new Map<string, PreviewCell>();
  private names = new Map<string, Rect>();
  private settings = new Map<string, string>();
  private selection = { row: 1, col: 1 };
  /** Number of the one sheet that exists; goes up when the sheet is deleted. */
  private sheet = 1;
  private version = 0;
  private contentVersion = 0;
  private snapshotCache: PreviewSnapshot | null = null;
  private selectionListeners = new Set<(selection: SheetSelection) => void>();
  private clickListeners = new Set<(click: SheetClick) => void>();
  private viewers = new Set<() => void>();

  private readonly persist: boolean;

  constructor(options: { persist?: boolean } = {}) {
    this.persist = options.persist !== false;
    if (this.persist) this.restore();
  }

  // -- SheetHost ------------------------------------------------------------

  async getSelection(): Promise<SheetSelection> {
    return this.currentSelection();
  }

  onSelectionChanged(listener: (selection: SheetSelection) => void): () => void {
    this.selectionListeners.add(listener);
    return () => {
      this.selectionListeners.delete(listener);
    };
  }

  onCellClicked(listener: (click: SheetClick) => void): () => void {
    this.clickListeners.add(listener);
    return () => {
      this.clickListeners.delete(listener);
    };
  }

  async getReportExtent(reportId: string): Promise<Extent | null> {
    const rect = this.names.get(reportId);
    return rect ? this.extent(rect) : null;
  }

  async hasContentOutside(target: Extent, previous: Extent | null): Promise<boolean> {
    const parts = subtract(target, previous && previous.sheetId === target.sheetId ? previous : null);
    if (parts.length === 0) return false;
    for (const [at, cell] of this.cells) {
      if (cell.value === "") continue;
      const [row, col] = at.split(",").map(Number);
      if (parts.some((part) => containsCell(part, row, col))) return true;
    }
    return false;
  }

  async writeReport(reportId: string, anchor: CellAddress, block: Block): Promise<Extent> {
    const previous = this.names.get(reportId);
    if (previous) this.clearRect(previous);
    const target: Rect = { row: anchor.row, col: anchor.col, rows: block.rows, cols: block.cols };
    this.clearRect(target);

    for (let r = 0; r < block.rows; r++) {
      for (let c = 0; c < block.cols; c++) {
        const value = block.values[r][c];
        const format = block.numberFormats[r][c];
        if (value === "" && format === "General") continue;
        this.cells.set(key(anchor.row + r, anchor.col + c), { value, format });
      }
    }
    const cellAt = (row: number, col: number): PreviewCell => {
      const at = key(anchor.row + row, anchor.col + col);
      let cell = this.cells.get(at);
      if (!cell) {
        cell = { value: "", format: "General" };
        this.cells.set(at, cell);
      }
      return cell;
    };
    const each = (rects: Rect[], apply: (cell: PreviewCell) => void) => {
      for (const rect of rects) {
        for (let r = rect.row; r < rect.row + rect.rows; r++) {
          for (let c = rect.col; c < rect.col + rect.cols; c++) apply(cellAt(r, c));
        }
      }
    };
    each(block.bold, (cell) => (cell.bold = true));
    each(block.fill, (cell) => (cell.fill = true));
    each(block.muted, (cell) => (cell.muted = true));
    for (const group of block.indents) each(group.rects, (cell) => (cell.indent = group.level));

    this.names.set(reportId, target);
    this.changed();
    return this.extent(target);
  }

  async clearReport(reportId: string): Promise<void> {
    const rect = this.names.get(reportId);
    if (!rect) return;
    this.clearRect(rect);
    this.names.set(reportId, { row: rect.row, col: rect.col, rows: 1, cols: 1 });
    this.changed();
  }

  async loadReports(): Promise<ReportDef[]> {
    const out: ReportDef[] = [];
    for (const [name, value] of this.settings) {
      if (name.indexOf(SETTING_PREFIX) !== 0) continue;
      const def = parseReportDef(value, name.slice(SETTING_PREFIX.length));
      if (def) out.push(def);
    }
    return out;
  }

  async saveReport(def: ReportDef): Promise<void> {
    this.settings.set(settingKey(def.id), serializeReportDef(def));
    this.changed();
  }

  async deleteReport(reportId: string): Promise<void> {
    this.settings.delete(settingKey(reportId));
    this.names.delete(reportId);
    this.changed();
  }

  // -- Preview only ---------------------------------------------------------

  /** Simulate moving the selection to a cell without clicking it (arrow keys). */
  select(row: number, col: number): void {
    if (row === this.selection.row && col === this.selection.col) return;
    this.selection = { row, col };
    this.changed(false);
    this.announceSelection();
  }

  /**
   * Simulate a mouse click on a cell: the selection moves there (if it was elsewhere),
   * and then the click itself is reported, as Excel's single-click event is: also when
   * the cell was already selected, and only for cells inside a report.
   */
  click(row: number, col: number): void {
    this.select(row, col);
    const sheet = this.sheet;
    setTimeout(() => {
      if (sheet !== this.sheet) return;
      const report = this.reportAt(row, col);
      if (!report) return;
      const cell = this.cells.get(key(row, col));
      const click: SheetClick = { report, value: cell ? String(cell.value) : "" };
      for (const listener of Array.from(this.clickListeners)) listener(click);
    }, 0);
  }

  /** Simulate the user typing into a cell (to try the overwrite question). */
  type(row: number, col: number, value: SheetValue): void {
    if (value === "") this.cells.delete(key(row, col));
    else this.cells.set(key(row, col), { value, format: "General" });
    this.changed();
  }

  /** Simulate inserting sheet rows above `at`: cells and report names move down. */
  insertRows(at: number, count: number): void {
    const moved = new Map<string, PreviewCell>();
    for (const [where, cell] of this.cells) {
      const [row, col] = where.split(",").map(Number);
      moved.set(key(row >= at ? row + count : row, col), cell);
    }
    this.cells = moved;
    for (const [id, rect] of this.names) {
      if (rect.row >= at) this.names.set(id, { ...rect, row: rect.row + count });
      else if (rect.row + rect.rows > at) this.names.set(id, { ...rect, rows: rect.rows + count });
    }
    if (this.selection.row >= at) this.selection = { ...this.selection, row: this.selection.row + count };
    this.changed();
  }

  /**
   * Simulate deleting the sheet: its cells are gone and every report name points at
   * nothing, but the report definitions stay in the workbook. A fresh, empty sheet
   * takes its place and becomes the active one.
   */
  deleteSheet(): void {
    this.cells.clear();
    this.names.clear();
    this.sheet++;
    this.selection = { row: 0, col: 0 };
    this.changed();
    this.announceSelection();
  }

  /** Empty the sheet and forget every report: a brand new workbook. */
  reset(): void {
    this.cells.clear();
    this.names.clear();
    this.settings.clear();
    this.sheet = 1;
    this.selection = { row: 1, col: 1 };
    this.changed();
    this.announceSelection();
  }

  subscribe(viewer: () => void): () => void {
    this.viewers.add(viewer);
    return () => {
      this.viewers.delete(viewer);
    };
  }

  /** Stable between changes, so it can back `useSyncExternalStore`. */
  snapshot(): PreviewSnapshot {
    if (this.snapshotCache && this.snapshotCache.version === this.version) return this.snapshotCache;
    let usedRows = 0;
    let usedCols = 0;
    for (const at of this.cells.keys()) {
      const [row, col] = at.split(",").map(Number);
      if (row + 1 > usedRows) usedRows = row + 1;
      if (col + 1 > usedCols) usedCols = col + 1;
    }
    this.snapshotCache = {
      version: this.version,
      contentVersion: this.contentVersion,
      sheetName: this.sheetName(),
      cells: this.cells,
      names: this.names,
      selection: this.selection,
      usedRows,
      usedCols,
    };
    return this.snapshotCache;
  }

  // -- internals ------------------------------------------------------------

  private sheetId(): string {
    return "preview-sheet-" + this.sheet;
  }

  private sheetName(): string {
    return "Sheet" + this.sheet;
  }

  private extent(rect: Rect): Extent {
    return { sheetId: this.sheetId(), sheetName: this.sheetName(), ...rect };
  }

  private reportAt(row: number, col: number): SheetSelection["report"] {
    for (const [id, rect] of this.names) {
      if (containsCell(rect, row, col)) {
        return { id, row: row - rect.row, col: col - rect.col, rows: rect.rows, cols: rect.cols };
      }
    }
    return null;
  }

  private currentSelection(): SheetSelection {
    const { row, col } = this.selection;
    return {
      cell: { sheetId: this.sheetId(), sheetName: this.sheetName(), row, col },
      report: this.reportAt(row, col),
    };
  }

  private announceSelection(): void {
    const selection = this.currentSelection();
    for (const listener of Array.from(this.selectionListeners)) listener(selection);
  }

  private clearRect(rect: Rect): void {
    if (rect.rows * rect.cols < this.cells.size) {
      for (let r = rect.row; r < rect.row + rect.rows; r++) {
        for (let c = rect.col; c < rect.col + rect.cols; c++) this.cells.delete(key(r, c));
      }
      return;
    }
    for (const at of Array.from(this.cells.keys())) {
      const [row, col] = at.split(",").map(Number);
      if (intersect(rect, { row, col, rows: 1, cols: 1 })) this.cells.delete(at);
    }
  }

  private changed(content = true): void {
    this.version++;
    if (content) this.contentVersion++;
    // A selection change alone is not worth re-serialising the whole sheet.
    if (content && this.persist) this.store();
    for (const viewer of Array.from(this.viewers)) viewer();
  }

  private store(): void {
    const target = storage();
    if (!target) return;
    try {
      const full: Stored = {
        cells: Array.from(this.cells.entries()),
        names: Array.from(this.names.entries()),
        settings: Array.from(this.settings.entries()),
        selection: this.selection,
        sheet: this.sheet,
      };
      let text = JSON.stringify(full);
      if (text.length > STORAGE_LIMIT) {
        // Too big to keep: remember the definitions only. After a reload the reports
        // are listed but have no cells, like a report whose sheet was deleted.
        text = JSON.stringify({ ...full, cells: [], names: [] });
      }
      target.setItem(STORAGE_KEY, text);
    } catch {
      // Storage full or blocked: the preview simply does not survive a reload.
    }
  }

  private restore(): void {
    const source = storage();
    if (!source) return;
    try {
      const text = source.getItem(STORAGE_KEY);
      if (!text) return;
      const stored = JSON.parse(text) as Partial<Stored>;
      this.cells = new Map(stored.cells ?? []);
      this.names = new Map(stored.names ?? []);
      this.settings = new Map(stored.settings ?? []);
      if (stored.selection) this.selection = stored.selection;
      if (typeof stored.sheet === "number" && stored.sheet >= 1) this.sheet = stored.sheet;
    } catch {
      // Unreadable: start empty.
    }
  }
}

/** A rough rendering of an Excel number format, good enough to eyeball a preview. */
export function formatForPreview(value: SheetValue, format: string): string {
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "string") {
    // A text format with literal text after the placeholder, such as the filter
    // cells' arrow: Excel shows the text followed by the literal.
    const suffix = /^@"([^"]*)"$/.exec(format);
    return suffix && value !== "" ? value + suffix[1] : value;
  }
  if (!isFinite(value)) return String(value);
  if (format === "General" || format === "@" || format === "") {
    return String(Math.round(value * 1e10) / 1e10);
  }
  const section = format.split(";")[0];
  const percent = section.indexOf("%") >= 0;
  const decimals = /\.(0+)/.exec(section);
  const digits = decimals ? decimals[1].length : 0;
  const text = (percent ? value * 100 : value).toLocaleString("en-US", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
    useGrouping: section.indexOf(",") >= 0,
  });
  const currency = /^\s*(?:"([^"]*)"|([$€£¥]))/.exec(section);
  const prefix = currency ? currency[1] ?? currency[2] ?? "" : "";
  return prefix + text + (percent ? "%" : "");
}
