// PreviewHost: a stand-in workbook for when the pane is opened in an ordinary browser
// (Office.onReady resolves with host null). It keeps sheets of cells, the report names
// and the "workbook settings" in memory, mirrors them to sessionStorage so that
// reloading the page behaves like reopening the workbook, and lets the page click a
// cell, move the selection with the keyboard, type into a cell, insert rows, switch
// sheets and delete one. `components/PreviewSheet.tsx` draws it.
//
// It imitates the parts of Excel the pane has to cope with: a deleted sheet leaves the
// report definitions behind with no cells; a click is reported after the selection it
// causes; moving with the keyboard selects without clicking; Show Details adds a sheet.

import { containsCell, intersect, subtract, type Rect } from "./geometry";
import {
  SETTING_PREFIX,
  SheetError,
  detailsSheetName,
  settingKey,
  type CellAddress,
  type Extent,
  type ReportDef,
  type SheetClick,
  type SheetHost,
  type SheetSelection,
} from "./host";
import { detailsCells, type Block, type SheetValue } from "./layout";
import { parseReportDef, serializeReportDef } from "../state/report";
import type { CellValue } from "../types";

export interface PreviewCell {
  value: SheetValue;
  format: string;
  bold?: boolean;
  fill?: boolean;
  muted?: boolean;
  indent?: number;
  /** A KPI Status or Trend cell: drawn as an icon, as Excel does. */
  icon?: "status" | "trend";
}

export interface PreviewSnapshot {
  /** Goes up on every change, including a change of selection. */
  version: number;
  /** Goes up only when cells, sheets or report names change. */
  contentVersion: number;
  /** All sheets, in tab order. */
  sheets: { id: string; name: string }[];
  activeId: string;
  sheetName: string;
  /** Cells of the active sheet. */
  cells: ReadonlyMap<string, PreviewCell>;
  /** Report id -> its block, for the reports on the active sheet. */
  names: ReadonlyMap<string, Rect>;
  selection: { row: number; col: number };
  /** Bottom-right corner of everything on the active sheet (exclusive). */
  usedRows: number;
  usedCols: number;
}

interface SheetData {
  id: string;
  name: string;
  cells: Map<string, PreviewCell>;
  selection: { row: number; col: number };
  /** Frozen header rows (Show Details freezes one). */
  frozenRows?: number;
}

interface NamedBlock {
  sheetId: string;
  rect: Rect;
}

const STORAGE_KEY = "cube.preview.v2";
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
  sheets: { id: string; name: string; cells: [string, PreviewCell][]; selection: { row: number; col: number }; frozenRows?: number }[];
  names: [string, NamedBlock][];
  settings: [string, string][];
  activeId: string;
  counter: number;
}

function newSheet(n: number): SheetData {
  return { id: "preview-sheet-" + n, name: "Sheet" + n, cells: new Map(), selection: { row: 1, col: 1 } };
}

export class PreviewHost implements SheetHost {
  readonly kind = "preview" as const;

  private sheets: SheetData[] = [newSheet(1)];
  private activeId = "preview-sheet-1";
  /** Number for the next sheet that gets a "SheetN" name. */
  private counter = 2;
  private names = new Map<string, NamedBlock>();
  private settings = new Map<string, string>();
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
    const named = this.names.get(reportId);
    if (!named) return null;
    const sheet = this.sheetById(named.sheetId);
    return sheet ? this.extent(sheet, named.rect) : null;
  }

  async hasContentOutside(target: Extent, previous: Extent | null): Promise<boolean> {
    const sheet = this.sheetById(target.sheetId);
    if (!sheet) return false;
    const parts = subtract(target, previous && previous.sheetId === target.sheetId ? previous : null);
    if (parts.length === 0) return false;
    for (const [at, cell] of sheet.cells) {
      if (cell.value === "") continue;
      const [row, col] = at.split(",").map(Number);
      if (parts.some((part) => containsCell(part, row, col))) return true;
    }
    return false;
  }

  async writeReport(reportId: string, anchor: CellAddress, block: Block): Promise<Extent> {
    const sheet = this.sheetById(anchor.sheetId);
    if (!sheet) throw new SheetError("The sheet this report was on can no longer be found.");
    const previous = this.names.get(reportId);
    if (previous) {
      const old = this.sheetById(previous.sheetId);
      if (old) clearRect(old.cells, previous.rect);
    }
    const target: Rect = { row: anchor.row, col: anchor.col, rows: block.rows, cols: block.cols };
    clearRect(sheet.cells, target);

    for (let r = 0; r < block.rows; r++) {
      for (let c = 0; c < block.cols; c++) {
        const value = block.values[r][c];
        const format = block.numberFormats[r][c];
        if (value === "" && format === "General") continue;
        sheet.cells.set(key(anchor.row + r, anchor.col + c), { value, format });
      }
    }
    const each = (rects: Rect[], apply: (cell: PreviewCell) => void) => {
      for (const rect of rects) {
        for (let r = rect.row; r < rect.row + rect.rows; r++) {
          for (let c = rect.col; c < rect.col + rect.cols; c++) {
            const at = key(anchor.row + r, anchor.col + c);
            let cell = sheet.cells.get(at);
            if (!cell) {
              cell = { value: "", format: "General" };
              sheet.cells.set(at, cell);
            }
            apply(cell);
          }
        }
      }
    };
    each(block.bold, (cell) => (cell.bold = true));
    each(block.fill, (cell) => (cell.fill = true));
    each(block.muted, (cell) => (cell.muted = true));
    for (const group of block.indents) each(group.rects, (cell) => (cell.indent = group.level));
    for (const icon of block.icons ?? []) each([icon.rect], (cell) => (cell.icon = icon.kind));

    this.names.set(reportId, { sheetId: sheet.id, rect: target });
    this.changed();
    return this.extent(sheet, target);
  }

  async clearReport(reportId: string): Promise<void> {
    const named = this.names.get(reportId);
    const sheet = named ? this.sheetById(named.sheetId) : null;
    if (!named || !sheet) return;
    clearRect(sheet.cells, named.rect);
    this.names.set(reportId, { sheetId: sheet.id, rect: { row: named.rect.row, col: named.rect.col, rows: 1, cols: 1 } });
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

  async addDetailsSheet(columns: string[], rows: CellValue[][]): Promise<string> {
    const name = detailsSheetName(this.sheets.map((s) => s.name));
    const sheet: SheetData = {
      id: "preview-details-" + this.counter++,
      name,
      cells: new Map(),
      selection: { row: 1, col: 0 },
      frozenRows: 1,
    };
    const { values, numberFormats, cols } = detailsCells(columns, rows);
    values.forEach((line, r) => {
      for (let c = 0; c < cols; c++) {
        const value = line[c];
        if (value === "" && r > 0) continue;
        sheet.cells.set(key(r, c), { value, format: numberFormats[r][c], bold: r === 0 || undefined });
      }
    });
    this.sheets.push(sheet);
    this.activeId = sheet.id;
    this.changed();
    this.announceSelection();
    return name;
  }

  async getFlag(name: string): Promise<boolean> {
    return this.settings.get(name) === "true";
  }

  async setFlag(name: string, on: boolean): Promise<void> {
    if (on) this.settings.set(name, "true");
    else this.settings.delete(name);
    this.changed();
  }

  // -- Preview only ---------------------------------------------------------

  /** Simulate moving the selection to a cell without clicking it (arrow keys). */
  select(row: number, col: number): void {
    const sheet = this.active();
    if (row === sheet.selection.row && col === sheet.selection.col) return;
    sheet.selection = { row, col };
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
    const sheetId = this.activeId;
    setTimeout(() => {
      if (sheetId !== this.activeId) return;
      const report = this.reportAt(sheetId, row, col);
      if (!report) return;
      const cell = this.active().cells.get(key(row, col));
      const click: SheetClick = { report, value: cell ? String(cell.value) : "" };
      for (const listener of Array.from(this.clickListeners)) listener(click);
    }, 0);
  }

  /** Simulate the user typing into a cell (to try the overwrite question). */
  type(row: number, col: number, value: SheetValue): void {
    const cells = this.active().cells;
    if (value === "") cells.delete(key(row, col));
    else cells.set(key(row, col), { value, format: "General" });
    this.changed();
  }

  /** Simulate inserting sheet rows above `at`: cells and report names move down. */
  insertRows(at: number, count: number): void {
    const sheet = this.active();
    const moved = new Map<string, PreviewCell>();
    for (const [where, cell] of sheet.cells) {
      const [row, col] = where.split(",").map(Number);
      moved.set(key(row >= at ? row + count : row, col), cell);
    }
    sheet.cells = moved;
    for (const [id, named] of this.names) {
      if (named.sheetId !== sheet.id) continue;
      const rect = named.rect;
      if (rect.row >= at) this.names.set(id, { ...named, rect: { ...rect, row: rect.row + count } });
      else if (rect.row + rect.rows > at) this.names.set(id, { ...named, rect: { ...rect, rows: rect.rows + count } });
    }
    if (sheet.selection.row >= at) sheet.selection = { ...sheet.selection, row: sheet.selection.row + count };
    this.changed();
  }

  /** Simulate clicking a sheet tab. */
  activate(sheetId: string): void {
    if (sheetId === this.activeId || !this.sheetById(sheetId)) return;
    this.activeId = sheetId;
    // Counts as a content change: the cells on show are another sheet's.
    this.changed();
    this.announceSelection();
  }

  /**
   * Simulate deleting the active sheet: its cells are gone and the reports on it point
   * at nothing, but their definitions stay in the workbook. Another sheet becomes the
   * active one; when it was the last, a fresh empty sheet takes its place.
   */
  deleteSheet(): void {
    const gone = this.activeId;
    this.sheets = this.sheets.filter((s) => s.id !== gone);
    for (const [id, named] of Array.from(this.names)) {
      if (named.sheetId === gone) this.names.delete(id);
    }
    if (this.sheets.length === 0) {
      const fresh = newSheet(this.counter++);
      fresh.selection = { row: 0, col: 0 };
      this.sheets.push(fresh);
    }
    this.activeId = this.sheets[0].id;
    this.changed();
    this.announceSelection();
  }

  /** Empty the workbook and forget every report: a brand new workbook. */
  reset(): void {
    this.sheets = [newSheet(1)];
    this.activeId = this.sheets[0].id;
    this.counter = 2;
    this.names.clear();
    this.settings.clear();
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
    const sheet = this.active();
    let usedRows = 0;
    let usedCols = 0;
    for (const at of sheet.cells.keys()) {
      const [row, col] = at.split(",").map(Number);
      if (row + 1 > usedRows) usedRows = row + 1;
      if (col + 1 > usedCols) usedCols = col + 1;
    }
    const names = new Map<string, Rect>();
    for (const [id, named] of this.names) if (named.sheetId === sheet.id) names.set(id, named.rect);
    this.snapshotCache = {
      version: this.version,
      contentVersion: this.contentVersion,
      sheets: this.sheets.map((s) => ({ id: s.id, name: s.name })),
      activeId: sheet.id,
      sheetName: sheet.name,
      cells: sheet.cells,
      names,
      selection: sheet.selection,
      usedRows,
      usedCols,
    };
    return this.snapshotCache;
  }

  /** Frozen header rows of the active sheet (for drawing). */
  frozenRows(): number {
    return this.active().frozenRows ?? 0;
  }

  // -- internals ------------------------------------------------------------

  private active(): SheetData {
    return this.sheetById(this.activeId) ?? this.sheets[0];
  }

  private sheetById(id: string): SheetData | null {
    return this.sheets.find((s) => s.id === id) ?? null;
  }

  private extent(sheet: SheetData, rect: Rect): Extent {
    return { sheetId: sheet.id, sheetName: sheet.name, ...rect };
  }

  private reportAt(sheetId: string, row: number, col: number): SheetSelection["report"] {
    for (const [id, named] of this.names) {
      const rect = named.rect;
      if (named.sheetId === sheetId && containsCell(rect, row, col)) {
        return { id, row: row - rect.row, col: col - rect.col, rows: rect.rows, cols: rect.cols };
      }
    }
    return null;
  }

  private currentSelection(): SheetSelection {
    const sheet = this.active();
    const { row, col } = sheet.selection;
    return {
      cell: { sheetId: sheet.id, sheetName: sheet.name, row, col },
      report: this.reportAt(sheet.id, row, col),
    };
  }

  private announceSelection(): void {
    const selection = this.currentSelection();
    for (const listener of Array.from(this.selectionListeners)) listener(selection);
  }

  private changed(content = true): void {
    this.version++;
    if (content) this.contentVersion++;
    // A selection change alone is not worth re-serialising the whole workbook.
    if (content && this.persist) this.store();
    for (const viewer of Array.from(this.viewers)) viewer();
  }

  private store(): void {
    const target = storage();
    if (!target) return;
    try {
      const full: Stored = {
        sheets: this.sheets.map((s) => ({
          id: s.id,
          name: s.name,
          cells: Array.from(s.cells.entries()),
          selection: s.selection,
          frozenRows: s.frozenRows,
        })),
        names: Array.from(this.names.entries()),
        settings: Array.from(this.settings.entries()),
        activeId: this.activeId,
        counter: this.counter,
      };
      let text = JSON.stringify(full);
      if (text.length > STORAGE_LIMIT) {
        // Too big to keep: remember the definitions only. After a reload the reports
        // are listed but have no cells, like a report whose sheet was deleted.
        text = JSON.stringify({
          ...full,
          sheets: [{ ...full.sheets[0], cells: [] }],
          names: [],
          activeId: full.sheets[0].id,
        });
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
      if (!stored.sheets || stored.sheets.length === 0) return;
      this.sheets = stored.sheets.map((s) => ({
        id: s.id,
        name: s.name,
        cells: new Map(s.cells),
        selection: s.selection,
        frozenRows: s.frozenRows,
      }));
      this.names = new Map(stored.names ?? []);
      this.settings = new Map(stored.settings ?? []);
      this.activeId = stored.activeId && this.sheetById(stored.activeId) ? stored.activeId : this.sheets[0].id;
      if (typeof stored.counter === "number") this.counter = stored.counter;
    } catch {
      // Unreadable: start empty.
    }
  }
}

function clearRect(cells: Map<string, PreviewCell>, rect: Rect): void {
  if (rect.rows * rect.cols < cells.size) {
    for (let r = rect.row; r < rect.row + rect.rows; r++) {
      for (let c = rect.col; c < rect.col + rect.cols; c++) cells.delete(key(r, c));
    }
    return;
  }
  for (const at of Array.from(cells.keys())) {
    const [row, col] = at.split(",").map(Number);
    if (intersect(rect, { row, col, rows: 1, cols: 1 })) cells.delete(at);
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
