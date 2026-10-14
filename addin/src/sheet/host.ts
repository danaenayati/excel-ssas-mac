// The one door between the pane and a sheet. `ExcelHost` (reports.ts + writer.ts)
// talks to Excel through Office.js; `PreviewHost` (preview.ts) is an in-page stand-in
// used when the pane is opened in an ordinary browser. The host is chosen once at
// startup from the result of `Office.onReady()`.

import type { CellValue, Layout } from "../types";
import type { ChartKind, ChartTable, ReportChart } from "./charts";
import { cellA1 } from "./geometry";
import type { Block } from "./layout";

/** A saved report. Stored as JSON in the workbook under `cube.report.<id>`. */
export interface ReportDef {
  id: string;
  connectionId: string;
  catalog: string;
  cube: string;
  layout: Layout;
  /** Charts that follow the report (sheet/charts.ts). Absent when it has none. */
  charts?: ReportChart[];
}

/** What `SheetHost.updateCharts` did. */
export interface ChartsUpdate {
  /** The charts still in the workbook; one the user deleted is left out. */
  charts: ReportChart[];
  /** Why a chart could not be updated; that chart is kept as it was. */
  error?: unknown;
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

  /** A workbook setting that holds a number, such as `cube.refreshMinutes`. Null when absent. */
  getNumber(key: string): Promise<number | null>;
  /** Null removes the setting. */
  setNumber(key: string, value: number | null): Promise<void>;

  /**
   * Calls `listener` (some time after) the workbook settings change: report definitions
   * saved by the right-click commands, which run apart from the pane, as well as the
   * pane's own saves. Optional: a host that cannot tell leaves it out. Returns the
   * unsubscribe function.
   */
  onSettingsChanged?(listener: () => void): () => void;

  /**
   * Convert to Formulas: write formulas into the report's block. Both matrices are the
   * block's size as last written; null leaves a cell's formula or number format as it
   * is. Number formats go first, so that a cell changed from text to General takes its
   * formula as a formula. Rejects with a SheetError when the block is gone or has
   * another size (rows or columns were inserted). The report's name is left in place.
   */
  writeFormulas(reportId: string, formulas: (string | null)[][], numberFormats: (string | null)[][]): Promise<void>;

  /** Recalculate every formula in the open workbooks, as Cmd+Option+F9 does; cube formulas ask the cube again. */
  recalculate(): Promise<void>;

  /**
   * Insert Chart: write `table` as a new data block on the hidden chart data sheet and
   * add a chart of `type` over it beside the report's block, with a title and a legend.
   * `index` is how many charts the report already has (each new one steps aside a
   * little). Rejects with a SheetError when the report is not on a sheet.
   */
  addChart(reportId: string, chartId: string, type: ChartKind, table: ChartTable, index: number): Promise<ReportChart>;

  /**
   * After a write of the report: write each chart's data block again from `table` and
   * point the chart at it. Type, size, place and formatting stay as the user left them;
   * the title follows the report unless the user has changed it. A chart the user
   * deleted is dropped and its data block no longer kept up to date.
   */
  updateCharts(charts: ReportChart[], table: ChartTable): Promise<ChartsUpdate>;

  /** Delete Chart: the chart and its data block. Does nothing for what is already gone. */
  deleteChart(chart: ReportChart): Promise<void>;
}

/** Workbook setting: refresh every report when the workbook is opened. */
export const REFRESH_ON_OPEN = "cube.refreshOnOpen";

/** Workbook setting: refresh every report this many minutes apart while the pane is open. */
export const REFRESH_MINUTES = "cube.refreshMinutes";

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
