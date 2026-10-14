// Runs a report: query the helper, turn the grid into a block, write it to the sheet,
// save the definition. Owns the rules about which request wins:
//
//   - A newer run supersedes an older one. The older query is told to stop and its
//     response, if it still arrives, is dropped.
//   - Cancel stops the query on the helper and drops the response.
//   - Sheet writes never overlap; each one finishes before the next starts.
//
// It also writes the placeholder a report with no fields leaves at its place on the
// sheet, and remembers, for each report, which member and which field each cell of the
// last result stands for.
//
// A report's charts (sheet/charts.ts) follow it as a PivotChart follows its PivotTable:
// every write of the report writes their data again. Insert Chart and Delete Chart go
// through the same queue as the writes, and the runner's list of a report's charts is
// the one saved, so that a run that started before a chart was added cannot drop it.
//
// No React in here, so the whole loop can be tested against the mock API and the
// preview host.

import type { Api } from "../api";
import { ApiFailure, HelperDown, isCancelled } from "../apiErrors";
import {
  chartRefusal,
  chartTable,
  withCharts,
  type ChartKind,
  type ChartTable,
  type ReportChart,
} from "../sheet/charts";
import { fitsSheet } from "../sheet/geometry";
import {
  SheetError,
  type CellAddress,
  type ChartsUpdate,
  type Extent,
  type ReportDef,
  type SheetHost,
} from "../sheet/host";
import {
  blockFields,
  buildBlock,
  columnHeaderAt,
  expandableMembers,
  expandedDescendants,
  memberKey,
  placeholderBlock,
  valueCellMembers,
  valueFieldIndex,
  type Axis,
  type Block,
  type BlockShape,
  type DataLine,
  type FieldCell,
  type GridHeaders,
  type MemberHit,
} from "../sheet/layout";
import type { ErrorCategory, Grid, Layout, MemberRef, PlacedField } from "../types";
import { isEmptyLayout } from "./layout";
import { newReportId } from "./report";

/** An error ready to show: a plain sentence, optional technical detail. */
export interface UiError {
  message: string;
  detail?: string;
  /** The helper could not be reached at all; the pane shows the restart command. */
  helperDown?: boolean;
  category?: ErrorCategory;
}

export function toUiError(error: unknown): UiError {
  if (error instanceof HelperDown) {
    return { message: "The helper is not running.", detail: error.message, helperDown: true };
  }
  if (error instanceof ApiFailure) {
    return {
      message: error.error.message,
      detail: error.error.detail,
      category: error.error.category,
    };
  }
  if (error instanceof SheetError) return { message: error.message, detail: error.detail };
  return {
    message: "Something went wrong in the pane.",
    detail: error instanceof Error ? error.stack || error.message : String(error),
  };
}

/** A chart that could not follow its report: said as such, with Excel's reason. */
function chartUpdateError(error: unknown): UiError {
  const ui = toUiError(error);
  return { ...ui, message: `The report was updated, but its chart could not be. ${ui.message}` };
}

/** What the pane remembers about a report's last result. Lost when the pane closes. */
export interface ReportRuntime {
  /** Size of the block as written. */
  rows: number;
  cols: number;
  members: Map<string, MemberHit>;
  headers: GridHeaders;
  stats: Grid["stats"];
  /** N: zero means "No data for this selection." */
  dataRows: number;
  shape: BlockShape;
  /** Field behind each filter cell of the block, keyed like `members`. */
  fieldCells: Map<string, FieldCell>;
  /** The layout that produced the block: the fields those cells point into. */
  layout: Layout;
  /** `Grid.colValueIndex` of the result, in sheet order: the value field behind each value column. */
  colValueIndex: number[];
  /**
   * Per data row of the block, the grid line it shows (none for an outer item's own line
   * in outline and compact form) and the row members behind it. Sheet rows map to grid
   * lines through this, never by position.
   */
  lines: DataLine[];
  /** Per value column of the block, the column members behind it. */
  columnMembers: string[][];
  /** What the block wrote into its cells, for Convert to Formulas (sheet/convert.ts). */
  sheet?: Pick<Block, "values" | "numberFormats">;
  /** What a chart of the report shows, from this result (Insert Chart). */
  chart?: ChartTable;
}

export type RunOutcome =
  | {
      kind: "written";
      extent: Extent;
      runtime: ReportRuntime;
      /** The definition as saved: the one run, with the charts still in the workbook. */
      def: ReportDef;
      /** The report was written but a chart of it could not be updated. */
      chartError?: UiError;
    }
  /**
   * The layout has no field in Rows, Columns or Values. `extent` is where the
   * placeholder was written, or null when the report has no place on a sheet yet.
   */
  | { kind: "empty"; extent: Extent | null }
  | { kind: "superseded" }
  | { kind: "cancelled" }
  /** The user said no to overwriting other cells. */
  | { kind: "declined" }
  /** The report has no place on a sheet and none was offered. */
  | { kind: "needs-anchor" }
  | { kind: "error"; error: UiError };

export interface RunOptions {
  /** Where to put a report that has no cells yet. */
  anchor: CellAddress | null;
  /** When there is neither an extent nor an anchor, use the active cell (if it is
   *  not inside another report). Never used for a placeholder. */
  useSelection: boolean;
  /** Put the report at `anchor` even though it has cells elsewhere, which are cleared. */
  relocate?: boolean;
}

export interface RunnerDeps {
  api: Api;
  host: SheetHost;
  /**
   * Ask in the pane whether `target` may be overwritten. Resolves true to go ahead.
   * `signal` aborts when the run is cancelled or superseded while the question is
   * open; the pane should then take the question down.
   */
  confirmOverwrite(def: ReportDef, target: Extent, signal: AbortSignal): Promise<boolean>;
  /** Called when a run moves from waiting on the cube to writing the sheet. */
  onStage?(stage: "query" | "sheet"): void;
}

interface Flight {
  seq: number;
  connectionId: string;
  abort: AbortController;
  stage: "query" | "sheet";
}

export class ReportRunner {
  readonly runtimes = new Map<string, ReportRuntime>();
  /** Each report's charts as last saved by this runner; the definitions the pane holds may lag. */
  private readonly charts = new Map<string, ReportChart[]>();
  private seq = 0;
  private flight: Flight | null = null;
  private cancelledSeqs = new Set<number>();
  private sheetQueue: Promise<unknown> = Promise.resolve();

  constructor(private readonly deps: RunnerDeps) {}

  get busy(): boolean {
    return this.flight !== null;
  }

  /** Stop the run in progress, if any. Its `run()` promise resolves with "cancelled". */
  cancel(): void {
    const flight = this.flight;
    if (!flight) return;
    this.cancelledSeqs.add(flight.seq);
    this.seq++;
    flight.abort.abort();
    if (flight.stage === "query") {
      this.deps.api.cancel(flight.connectionId).catch(() => undefined);
    }
  }

  /** The member behind a cell of a report, from its last result. */
  memberAt(reportId: string, row: number, col: number): MemberHit | null {
    const runtime = this.runtimes.get(reportId);
    return runtime ? runtime.members.get(memberKey(row, col)) ?? null : null;
  }

  /** Expanded members under `member` in the report's last result (for Collapse). */
  descendants(reportId: string, axis: Axis, member: MemberRef): string[] {
    const runtime = this.runtimes.get(reportId);
    return runtime ? expandedDescendants(runtime.headers, axis, member) : [];
  }

  /**
   * The last result of a report, if it still describes the block: null when this
   * session has not written the report, or when the block has changed size since (rows
   * or columns were inserted), so that positions can no longer be trusted.
   */
  private current(reportId: string, size: { rows: number; cols: number }): ReportRuntime | null {
    const runtime = this.runtimes.get(reportId);
    return runtime && runtime.rows === size.rows && runtime.cols === size.cols ? runtime : null;
  }

  /**
   * The field behind a filter cell (a page filter's caption or value cell, a row
   * field's caption, a column field's caption): the cells that open a checklist when
   * clicked. Null for any other cell; undefined when there is no result to tell by.
   */
  buttonAt(
    reportId: string,
    at: { row: number; col: number; rows: number; cols: number },
  ): PlacedField | null | undefined {
    const runtime = this.current(reportId, at);
    if (!runtime) return undefined;
    const cell = runtime.fieldCells.get(memberKey(at.row, at.col));
    return cell ? runtime.layout[cell.area][cell.index] ?? null : null;
  }

  /**
   * For compact form's "Row Labels" or "Column Labels" cell, its area: the checklist it
   * opens offers each field of that area (Windows' "Select field:"). Null for any other
   * cell, and when there is no result to tell by.
   */
  labelsAt(
    reportId: string,
    at: { row: number; col: number; rows: number; cols: number },
  ): "rows" | "columns" | null {
    const runtime = this.current(reportId, at);
    const cell = runtime ? runtime.fieldCells.get(memberKey(at.row, at.col)) : undefined;
    return cell && cell.labels && cell.area !== "filters" ? cell.area : null;
  }

  /**
   * The field the pane offers to filter for a selected cell: a filter cell's field, or
   * the column field whose header row the cell is in. Same null / undefined as above.
   */
  fieldAt(
    reportId: string,
    at: { row: number; col: number; rows: number; cols: number },
  ): PlacedField | null | undefined {
    const button = this.buttonAt(reportId, at);
    if (button !== null) return button;
    const runtime = this.current(reportId, at) as ReportRuntime;
    const header = columnHeaderAt(runtime.shape, at.row, at.col, runtime.layout.columns.length);
    return header === null ? null : runtime.layout.columns[header] ?? null;
  }

  /**
   * What Show Details needs for a cell: the members behind it, the measure of its value
   * field, and the layout that produced the result. Null when the cell is not a value
   * cell, or its value field cannot be told; undefined when there is no result to tell by.
   */
  detailsFor(
    reportId: string,
    at: { row: number; col: number; rows: number; cols: number },
  ): { members: string[]; measure: string; layout: Layout } | null | undefined {
    const runtime = this.current(reportId, at);
    if (!runtime) return undefined;
    const shape = runtime.shape;
    const line = at.row - shape.dataTop;
    const column = at.col - shape.headerCols;
    if (line < 0 || line >= shape.dataRows || column < 0 || column >= shape.dataCols) return null;
    const index = valueFieldIndex(runtime.colValueIndex, column, runtime.layout.values.length);
    if (index === null) return null;
    const members = valueCellMembers(runtime, line, column);
    if (members === null) return null;
    return {
      members,
      measure: runtime.layout.values[index].measure,
      layout: runtime.layout,
    };
  }

  /**
   * For "Expand Entire Field": the field's members in the report's last result that
   * have children and are not expanded, and whether any shown member has children.
   * Null when there is no result in memory.
   */
  expandable(reportId: string, hierarchy: string): { members: string[]; anyChildren: boolean } | null {
    const runtime = this.runtimes.get(reportId);
    return runtime ? expandableMembers(runtime.headers, hierarchy) : null;
  }

  /** Also stops following the report's charts: deleted or converted, it leaves them as they are. */
  forget(reportId: string): void {
    this.runtimes.delete(reportId);
    this.charts.delete(reportId);
  }

  /** The report's charts: as this runner last saved them, else as its definition says. */
  chartsOf(def: ReportDef): ReportChart[] {
    return this.charts.get(def.id) ?? def.charts ?? [];
  }

  /**
   * Insert Chart: a chart of `type` beside the report, showing its last result. The
   * report must have been written in this session. Resolves with the definition as
   * saved, or with why there is no chart.
   */
  async insertChart(def: ReportDef, type: ChartKind): Promise<{ def: ReportDef; chart: ReportChart } | { error: UiError }> {
    const unwritten = { error: { message: "Refresh the report first, then insert the chart." } };
    if (!this.chartSource(def.id)) return unwritten;
    try {
      return await this.inSheetQueue(async () => {
        // A write queued before this one may have brought a newer result, or none.
        const table = this.chartSource(def.id);
        if (!table) return unwritten;
        const refused = chartRefusal(table);
        if (refused) return { error: { message: refused } };
        const before = this.chartsOf(def);
        const chart = await this.deps.host.addChart(def.id, newReportId(), type, table, before.length);
        const charts = [...before, chart];
        this.charts.set(def.id, charts);
        const saved = withCharts(def, charts);
        await this.deps.host.saveReport(saved);
        return { def: saved, chart };
      });
    } catch (error) {
      return { error: toUiError(error) };
    }
  }

  /** Delete Chart: the chart and its data. Resolves with the definition as saved. */
  async deleteChart(def: ReportDef, chartId: string): Promise<{ def: ReportDef } | { error: UiError }> {
    try {
      return await this.inSheetQueue(async () => {
        const before = this.chartsOf(def);
        const chart = before.find((c) => c.id === chartId);
        if (chart) await this.deps.host.deleteChart(chart);
        const charts = before.filter((c) => c.id !== chartId);
        this.charts.set(def.id, charts);
        const saved = withCharts(def, charts);
        await this.deps.host.saveReport(saved);
        return { def: saved };
      });
    } catch (error) {
      return { error: toUiError(error) };
    }
  }

  private chartSource(reportId: string): ChartTable | null {
    const runtime = this.runtimes.get(reportId);
    return runtime && runtime.chart ? runtime.chart : null;
  }

  /**
   * After a write: the report's charts show the new result. A chart deleted in Excel is
   * dropped from the definition without a word. Never fails the write: a chart that
   * cannot be updated keeps its last picture and is tried again next time.
   */
  private async followCharts(def: ReportDef, table: ChartTable): Promise<{ def: ReportDef; error?: UiError }> {
    const charts = this.chartsOf(def);
    this.charts.set(def.id, charts);
    if (charts.length === 0) return { def: withCharts(def, charts) };
    const refused = chartRefusal(table);
    if (refused) return { def: withCharts(def, charts), error: { message: `The chart was not updated. ${refused}` } };
    let update: ChartsUpdate;
    try {
      update = await this.deps.host.updateCharts(charts, table);
    } catch (error) {
      return { def: withCharts(def, charts), error: chartUpdateError(error) };
    }
    this.charts.set(def.id, update.charts);
    const saved = withCharts(def, update.charts);
    return update.error === undefined ? { def: saved } : { def: saved, error: chartUpdateError(update.error) };
  }

  async run(def: ReportDef, options: RunOptions): Promise<RunOutcome> {
    const mine = ++this.seq;
    const stale = (): RunOutcome | null => {
      if (mine === this.seq) return null;
      const cancelled = this.cancelledSeqs.delete(mine);
      return { kind: cancelled ? "cancelled" : "superseded" };
    };

    // Tell the helper to drop the query this run replaces before sending a new one,
    // so the two cannot be confused on the server.
    const previous = this.flight;
    if (previous) previous.abort.abort();
    if (previous && previous.stage === "query") {
      try {
        await this.deps.api.cancel(previous.connectionId);
      } catch {
        // Nothing to stop, or the helper is gone; the query below will say so.
      }
      const gone = stale();
      if (gone) return gone;
    }

    const flight: Flight = {
      seq: mine,
      connectionId: def.connectionId,
      abort: new AbortController(),
      stage: "query",
    };
    this.flight = flight;

    try {
      if (isEmptyLayout(def.layout)) {
        flight.stage = "sheet";
        return await this.inSheetQueue(() => this.writePlaceholder(def, options, flight, stale));
      }

      this.deps.onStage?.("query");
      let grid: Grid;
      try {
        grid = await this.query(def, flight, mine);
      } catch (error) {
        const gone = stale();
        if (gone) return gone;
        if (isCancelled(error)) return { kind: "cancelled" };
        return { kind: "error", error: toUiError(error) };
      }
      const gone = stale();
      if (gone) return gone;

      flight.stage = "sheet";
      this.deps.onStage?.("sheet");
      return await this.inSheetQueue(() => this.write(def, grid, options, flight, stale));
    } catch (error) {
      const gone = stale();
      if (gone) return gone;
      return { kind: "error", error: toUiError(error) };
    } finally {
      if (this.flight === flight) this.flight = null;
      this.cancelledSeqs.delete(mine);
    }
  }

  private async query(def: ReportDef, flight: Flight, mine: number): Promise<Grid> {
    const request = { catalog: def.catalog, cube: def.cube, layout: def.layout };
    try {
      return await this.deps.api.query(def.connectionId, request, flight.abort.signal);
    } catch (error) {
      // A "cancelled" nobody asked for: the stop sent for the previous query landed on
      // this one. Ask once more.
      if (isCancelled(error) && mine === this.seq && !flight.abort.signal.aborted) {
        return await this.deps.api.query(def.connectionId, request, flight.abort.signal);
      }
      throw error;
    }
  }

  private async write(
    def: ReportDef,
    grid: Grid,
    options: RunOptions,
    flight: Flight,
    stale: () => RunOutcome | null,
  ): Promise<RunOutcome> {
    const block = buildBlock(grid, blockFields(def.layout));
    const placed = await this.place(def, block, options, flight, stale);
    if ("kind" in placed) return placed;
    // The block's grid has the columns in sheet order.
    const shown = block.grid;
    const table = chartTable(shown, def.layout);
    const runtime: ReportRuntime = {
      rows: block.rows,
      cols: block.cols,
      members: block.members,
      headers: { rowHeaders: shown.rowHeaders, colHeaderRows: shown.colHeaderRows },
      stats: grid.stats,
      dataRows: block.shape.dataRows,
      shape: block.shape,
      fieldCells: block.fieldCells,
      layout: def.layout,
      colValueIndex: shown.colValueIndex ?? [],
      lines: block.lines,
      columnMembers: block.columnMembers,
      sheet: { values: block.values, numberFormats: block.numberFormats },
      chart: table,
    };
    this.runtimes.set(def.id, runtime);
    const followed = await this.followCharts(def, table);
    await this.deps.host.saveReport(followed.def);
    const outcome: RunOutcome = { kind: "written", extent: placed, runtime, def: followed.def };
    if (followed.error) outcome.chartError = followed.error;
    return outcome;
  }

  /**
   * A report with no fields still shows where it is: one quiet cell, tracked by the
   * report's name like any block. Written only where the report already has a place
   * (its old block, or the cell it was put at), never at whatever cell is selected.
   */
  private async writePlaceholder(
    def: ReportDef,
    options: RunOptions,
    flight: Flight,
    stale: () => RunOutcome | null,
  ): Promise<RunOutcome> {
    const placed = await this.place(def, placeholderBlock(), { ...options, useSelection: false }, flight, stale);
    if (!("kind" in placed)) {
      // Charts keep their last picture until the report has fields again.
      const charts = this.chartsOf(def);
      this.runtimes.delete(def.id);
      await this.deps.host.saveReport(withCharts(def, charts));
      return { kind: "empty", extent: placed };
    }
    if (placed.kind === "needs-anchor") {
      this.runtimes.delete(def.id);
      return { kind: "empty", extent: null };
    }
    return placed;
  }

  /**
   * Put a block on the sheet for a report: find its place, check it fits, ask before
   * covering other cells, write. Resolves with the new extent, or with the outcome
   * that stopped it.
   */
  private async place(
    def: ReportDef,
    block: Block,
    options: RunOptions,
    flight: Flight,
    stale: () => RunOutcome | null,
  ): Promise<Extent | RunOutcome> {
    const { host } = this.deps;
    let gone = stale();
    if (gone) return gone;

    const previous = await host.getReportExtent(def.id);
    let anchor: CellAddress | null =
      options.relocate && options.anchor ? options.anchor : previous ?? options.anchor;
    if (!anchor && options.useSelection) {
      const selection = await host.getSelection();
      if (!selection.report) anchor = selection.cell;
    }
    if (!anchor) return { kind: "needs-anchor" };

    const target: Extent = {
      sheetId: anchor.sheetId,
      sheetName: anchor.sheetName,
      row: anchor.row,
      col: anchor.col,
      rows: block.rows,
      cols: block.cols,
    };
    if (!fitsSheet(target)) {
      return {
        kind: "error",
        error: {
          message:
            `This view needs ${block.rows.toLocaleString("en-US")} rows and ` +
            `${block.cols.toLocaleString("en-US")} columns, which does not fit on the sheet ` +
            "from where the report starts. Add a filter or remove a field.",
        },
      };
    }

    if (await host.hasContentOutside(target, previous)) {
      const signal = flight.abort.signal;
      const ok = await new Promise<boolean>((resolve) => {
        if (signal.aborted) {
          resolve(false);
          return;
        }
        signal.addEventListener("abort", () => resolve(false));
        this.deps.confirmOverwrite(def, target, signal).then(resolve, () => resolve(false));
      });
      gone = stale();
      if (gone) return gone;
      if (!ok) return { kind: "declined" };
    }
    gone = stale();
    if (gone) return gone;

    return await host.writeReport(def.id, anchor, block);
  }

  private inSheetQueue<T>(task: () => Promise<T>): Promise<T> {
    const next = this.sheetQueue.then(task, task);
    this.sheetQueue = next.catch(() => undefined);
    return next;
  }
}
