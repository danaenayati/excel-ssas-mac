// Runs a report: query the helper, turn the grid into a block, write it to the sheet,
// save the definition. Owns the rules about which request wins:
//
//   - A newer run supersedes an older one. The older query is told to stop and its
//     response, if it still arrives, is dropped.
//   - Cancel stops the query on the helper and drops the response.
//   - Sheet writes never overlap; each one finishes before the next starts.
//
// No React in here, so the whole loop can be tested against the mock API and the
// preview host.

import type { Api } from "../api";
import { ApiFailure, HelperDown, isCancelled } from "../apiErrors";
import { fitsSheet } from "../sheet/geometry";
import { SheetError, type CellAddress, type Extent, type ReportDef, type SheetHost } from "../sheet/host";
import {
  buildBlock,
  expandedDescendants,
  memberKey,
  type Axis,
  type GridHeaders,
  type MemberHit,
} from "../sheet/layout";
import type { ErrorCategory, Grid, MemberRef } from "../types";
import { isEmptyLayout } from "./layout";

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
}

export type RunOutcome =
  | { kind: "written"; extent: Extent; runtime: ReportRuntime }
  /** The layout has no field in Rows, Columns or Values; any old block was emptied. */
  | { kind: "empty"; cleared: boolean }
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
   *  not inside another report). */
  useSelection: boolean;
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

  forget(reportId: string): void {
    this.runtimes.delete(reportId);
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
        return await this.inSheetQueue(async () => {
          const gone = stale();
          if (gone) return gone;
          const extent = await this.deps.host.getReportExtent(def.id);
          if (extent) {
            await this.deps.host.clearReport(def.id);
            await this.deps.host.saveReport(def);
          }
          this.runtimes.delete(def.id);
          return { kind: "empty", cleared: extent !== null } as RunOutcome;
        });
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
    const { host } = this.deps;
    let gone = stale();
    if (gone) return gone;

    const block = buildBlock(grid);
    const previous = await host.getReportExtent(def.id);
    let anchor: CellAddress | null = previous ?? options.anchor;
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

    const extent = await host.writeReport(def.id, anchor, block);
    const runtime: ReportRuntime = {
      rows: block.rows,
      cols: block.cols,
      members: block.members,
      headers: { rowHeaders: grid.rowHeaders, colHeaderRows: grid.colHeaderRows },
      stats: grid.stats,
      dataRows: block.shape.dataRows,
    };
    this.runtimes.set(def.id, runtime);
    await host.saveReport(def);
    return { kind: "written", extent, runtime };
  }

  private inSheetQueue<T>(task: () => Promise<T>): Promise<T> {
    const next = this.sheetQueue.then(task, task);
    this.sheetQueue = next.catch(() => undefined);
    return next;
  }
}
