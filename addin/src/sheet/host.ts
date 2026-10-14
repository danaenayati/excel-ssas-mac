// The one door between the pane and a sheet. `ExcelHost` (reports.ts + writer.ts)
// talks to Excel through Office.js; `PreviewHost` (preview.ts) is an in-page stand-in
// used when the pane is opened in an ordinary browser. The host is chosen once at
// startup from the result of `Office.onReady()`.

import type { CellValue, Layout } from "../types";
import { cellA1 } from "./geometry";
import type { Block } from "./layout";

/** A saved report. Stored as JSON in the workbook under `cube.report.<id>`. */
export interface ReportDef {
  id: string;
  connectionId: string;
  catalog: string;
  cube: string;
  layout: Layout;
}

/** One cell. `row` and `col` are zero-based. */
export interface CellAddress {
  /** Stable worksheet id (survives a rename). */
  sheetId: string;
  /** Worksheet name, for display. */
  sheetName: string;
  row: number;
  col: number;
}

/** A rectangle of cells; `row`/`col` is its top-left cell. */
export interface Extent extends CellAddress {
  rows: number;
  cols: number;
}

export interface SheetSelection {
  /** The active cell. */
  cell: CellAddress;
  /**
   * The report whose block contains the active cell, or null. `row`/`col` are the
   * cell's offsets from the block's top-left; `rows`/`cols` are the block's current
   * size (which differs from the size last written if the user inserted rows or columns).
   */
  report: { id: string; row: number; col: number; rows: number; cols: number } | null;
}

/** The user clicked a cell inside a report's block. */
export interface SheetClick {
  /** As in `SheetSelection`: the report, the cell's offsets in its block, the block's size. */
  report: { id: string; row: number; col: number; rows: number; cols: number };
  /** The cell's value as text, without anything its number format adds. */
  value: string;
}

export interface SheetHost {
  readonly kind: "excel" | "preview";

  /** The active cell, and the report it is in, if any. */
  getSelection(): Promise<SheetSelection>;

  /** Calls `listener` whenever the active cell changes. Returns the unsubscribe function. */
  onSelectionChanged(listener: (selection: SheetSelection) => void): () => void;

  /**
   * Calls `listener` when a cell inside a report is clicked, including a click on the
   * cell that is already selected (which changes no selection). Never called for
   * keyboard movement. A host that cannot hear clicks simply never calls it; the pane
   * then relies on the selection. Returns the unsubscribe function.
   */
  onCellClicked(listener: (click: SheetClick) => void): () => void;

  /**
   * Where a report's block is now, resolved from the workbook name `CubeReport_<id>`;
   * null when the report has never been written or its name no longer points at cells.
   */
  getReportExtent(reportId: string): Promise<Extent | null>;

  /**
   * True when `target` covers a non-empty cell that is not inside `previous` (the
   * extent about to be cleared anyway). Used to ask before overwriting.
   */
  hasContentOutside(target: Extent, previous: Extent | null): Promise<boolean>;

  /**
   * Clears the report's previous extent (wherever it is), writes `block` with its
   * top-left at `anchor`, formats it, and points `CubeReport_<id>` at the new extent,
   * which it returns.
   */
  writeReport(reportId: string, anchor: CellAddress, block: Block): Promise<Extent>;

  /**
   * Empties the report's cells. The name stays on the top-left cell so the report
   * keeps its place. Does nothing when the report has no extent.
   */
  clearReport(reportId: string): Promise<void>;

  /** Every readable report definition in the workbook. Unreadable ones are skipped. */
  loadReports(): Promise<ReportDef[]>;

  saveReport(def: ReportDef): Promise<void>;

  /** Forgets the definition and the name. Cells are left as they are. */
  deleteReport(reportId: string): Promise<void>;

  /**
   * Show Details: add a sheet named "Details" (or "Details 2", "Details 3", …) holding a
   * bold header row and the rows, with the top row frozen where Excel allows it, and
   * make it the active sheet. Resolves with the new sheet's name.
   */
  addDetailsSheet(columns: string[], rows: CellValue[][]): Promise<string>;

  /** A workbook setting that is on or off, such as `cube.refreshOnOpen`. Off when absent. */
  getFlag(key: string): Promise<boolean>;
  setFlag(key: string, on: boolean): Promise<void>;
}

/** Workbook setting: refresh every report when the workbook is opened. */
export const REFRESH_ON_OPEN = "cube.refreshOnOpen";

/** "Details", then "Details 2", "Details 3", … : the first name no sheet has (case aside). */
export function detailsSheetName(existing: string[]): string {
  const taken = new Set(existing.map((name) => name.toLowerCase()));
  if (!taken.has("details")) return "Details";
  let n = 2;
  while (taken.has(`details ${n}`)) n++;
  return `Details ${n}`;
}

/** A sheet operation failed. `message` is plain language; `detail` is the raw error. */
export class SheetError extends Error {
  readonly detail?: string;
  constructor(message: string, detail?: string) {
    super(message);
    this.name = "SheetError";
    this.detail = detail;
  }
}

export const SETTING_PREFIX = "cube.report.";
export const NAME_PREFIX = "CubeReport_";

export function settingKey(reportId: string): string {
  return SETTING_PREFIX + reportId;
}

export function reportName(reportId: string): string {
  return NAME_PREFIX + reportId;
}

/** "Sheet1!B2" */
export function describeCell(cell: CellAddress): string {
  return cell.sheetName + "!" + cellA1(cell.row, cell.col);
}
