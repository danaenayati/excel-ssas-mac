// Office.js: report definitions in workbook settings, report locations as workbook
// names, the selection, and `ExcelHost`, which ties these and writer.ts together
// behind the `SheetHost` interface.
//
// Requirement sets: ExcelApi 1.4 for settings and names; 1.2 for
// workbook.onSelectionChanged; guarded 1.7 for getActiveCell, worksheets.onActivated,
// worksheets.onDeleted and worksheets.onAdded; guarded 1.10 for onSingleClicked.
// Without click events nothing is lost but a shortcut: the pane offers the same
// "Filter <field>…" action for the selected cell.

import { parseReportDef, serializeReportDef } from "../state/report";
import { cellA1, containsCell, parseA1, type Rect } from "./geometry";
import {
  NAME_PREFIX,
  SETTING_PREFIX,
  reportName,
  settingKey,
  type CellAddress,
  type Extent,
  type ReportDef,
  type SheetClick,
  type SheetHost,
  type SheetSelection,
} from "./host";
import type { Block } from "./layout";
import {
  clearBlock,
  hasContentOutside,
  readExtent,
  runExcel,
  supportsExcelApi,
  writeBlock,
} from "./writer";

/** Document setting that makes Excel reopen the pane together with the workbook. */
const AUTO_SHOW = "Office.AutoShowTaskpaneWithDocument";

/** Wait this long after a selection event before reading it: arrow-key travel sends many. */
const SELECTION_SETTLE_MS = 80;

/**
 * Turn on "open this pane with the workbook". Done through the document settings API,
 * as Office requires, and before the first report definition is written through
 * `workbook.settings`, so that saving this bag can never overwrite a newer definition.
 * Never fails: a workbook that will not reopen the pane is still a working workbook.
 */
function ensureAutoShow(): Promise<void> {
  return new Promise<void>((resolve) => {
    try {
      const settings = Office.context.document.settings;
      if (settings.get(AUTO_SHOW) === true) {
        resolve();
        return;
      }
      settings.set(AUTO_SHOW, true);
      settings.saveAsync(() => resolve());
    } catch {
      resolve();
    }
  });
}

function activeCell(context: Excel.RequestContext): Excel.Range {
  // getSelectedRange fails when several separate areas are selected; getActiveCell
  // does not, but needs 1.7.
  return supportsExcelApi("1.7")
    ? context.workbook.getActiveCell()
    : context.workbook.getSelectedRange().getCell(0, 0);
}

async function readSelection(context: Excel.RequestContext): Promise<SheetSelection> {
  const cell = activeCell(context);
  cell.load("rowIndex,columnIndex,worksheet/id,worksheet/name");
  const names = context.workbook.names;
  names.load("items/name");
  await context.sync();

  const address: CellAddress = {
    sheetId: cell.worksheet.id,
    sheetName: cell.worksheet.name,
    row: cell.rowIndex,
    col: cell.columnIndex,
  };

  const prefix = NAME_PREFIX.toLowerCase();
  const mine = names.items.filter((item) => item.name.toLowerCase().indexOf(prefix) === 0);
  if (mine.length === 0) return { cell: address, report: null };

  const ranges = mine.map((item) => {
    const range = item.getRangeOrNullObject();
    range.load("rowIndex,columnIndex,rowCount,columnCount,worksheet/id");
    return range;
  });
  await context.sync();

  for (let i = 0; i < mine.length; i++) {
    const range = ranges[i];
    if (range.isNullObject || range.worksheet.id !== address.sheetId) continue;
    const rect = {
      row: range.rowIndex,
      col: range.columnIndex,
      rows: range.rowCount,
      cols: range.columnCount,
    };
    if (containsCell(rect, address.row, address.col)) {
      return {
        cell: address,
        report: {
          id: mine[i].name.slice(NAME_PREFIX.length),
          row: address.row - rect.row,
          col: address.col - rect.col,
          rows: rect.rows,
          cols: rect.cols,
        },
      };
    }
  }
  return { cell: address, report: null };
}

/** Remove event handlers registered through `Excel.run`. */
function removeHandlers(handlers: OfficeExtension.EventHandlerResult<unknown>[]): void {
  for (const handler of handlers) {
    Excel.run(handler.context, async (context) => {
      handler.remove();
      await context.sync();
    }).catch(() => undefined);
  }
}

/** Subscribe to "the selection may have changed". Resolves with the unsubscribe function. */
async function watchSelection(notify: () => void): Promise<() => void> {
  try {
    const handlers: OfficeExtension.EventHandlerResult<unknown>[] = [];
    await Excel.run(async (context) => {
      handlers.push(context.workbook.onSelectionChanged.add(async () => notify()));
      if (supportsExcelApi("1.7")) {
        handlers.push(context.workbook.worksheets.onActivated.add(async () => notify()));
      }
      await context.sync();
    });
    // A deleted sheet takes its reports' cells with it. Registered on its own, so
    // that a refusal here cannot cost the two subscriptions above.
    if (supportsExcelApi("1.7")) {
      try {
        await Excel.run(async (context) => {
          handlers.push(context.workbook.worksheets.onDeleted.add(async () => notify()));
          await context.sync();
        });
      } catch {
        // Sheet activation covers the usual case: deleting the sheet one is on.
      }
    }
    return () => removeHandlers(handlers);
  } catch {
    // Fall back to the older document-level event, which exists in every Excel.
    const handler = () => notify();
    return new Promise<() => void>((resolve) => {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        handler,
        () =>
          resolve(() => {
            Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, {
              handler,
            });
          }),
      );
    });
  }
}

/**
 * Subscribe to single clicks on every sheet (ExcelApi 1.10): one handler on the
 * worksheet collection, or, if that is refused, one per sheet and per sheet added
 * later. Resolves with the unsubscribe function, or null when clicks cannot be heard.
 */
async function watchClicks(
  notify: (worksheetId: string, address: string) => void,
): Promise<(() => void) | null> {
  if (!supportsExcelApi("1.10")) return null;
  const handlers: OfficeExtension.EventHandlerResult<unknown>[] = [];
  const onClicked = async (args: Excel.WorksheetSingleClickedEventArgs) =>
    notify(args.worksheetId, args.address);
  try {
    await Excel.run(async (context) => {
      handlers.push(context.workbook.worksheets.onSingleClicked.add(onClicked));
      await context.sync();
    });
    return () => removeHandlers(handlers);
  } catch {
    handlers.length = 0;
  }
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/id");
      await context.sync();
      for (const sheet of sheets.items) handlers.push(sheet.onSingleClicked.add(onClicked));
      handlers.push(
        sheets.onAdded.add(async (args) => {
          await Excel.run(async (inner) => {
            const added = inner.workbook.worksheets.getItem(args.worksheetId);
            handlers.push(added.onSingleClicked.add(onClicked));
            await inner.sync();
          }).catch(() => undefined);
        }),
      );
      await context.sync();
    });
    return () => removeHandlers(handlers);
  } catch {
    removeHandlers(handlers);
    return null;
  }
}

/** The report a clicked cell is in, with the cell's value; null when it is in none. */
async function readClick(
  context: Excel.RequestContext,
  worksheetId: string,
  cell: Rect,
): Promise<SheetClick | null> {
  const names = context.workbook.names;
  names.load("items/name");
  await context.sync();
  const prefix = NAME_PREFIX.toLowerCase();
  const mine = names.items.filter((item) => item.name.toLowerCase().indexOf(prefix) === 0);
  if (mine.length === 0) return null;

  const blocks = mine.map((item) => {
    const range = item.getRangeOrNullObject();
    range.load("rowIndex,columnIndex,rowCount,columnCount,worksheet/id");
    return range;
  });
  const clicked = context.workbook.worksheets.getItem(worksheetId).getRange(cellA1(cell.row, cell.col));
  clicked.load("values");
  await context.sync();

  for (let i = 0; i < mine.length; i++) {
    const range = blocks[i];
    if (range.isNullObject || range.worksheet.id !== worksheetId) continue;
    const block: Rect = {
      row: range.rowIndex,
      col: range.columnIndex,
      rows: range.rowCount,
      cols: range.columnCount,
    };
    if (!containsCell(block, cell.row, cell.col)) continue;
    const value = (clicked.values[0] || [])[0];
    return {
      report: {
        id: mine[i].name.slice(NAME_PREFIX.length),
        row: cell.row - block.row,
        col: cell.col - block.col,
        rows: block.rows,
        cols: block.cols,
      },
      value: value === null || value === undefined ? "" : String(value),
    };
  }
  return null;
}

export class ExcelHost implements SheetHost {
  readonly kind = "excel" as const;

  /** Reports this workbook is known to hold; with none, clicks are not looked at. */
  private readonly known = new Set<string>();

  getSelection(): Promise<SheetSelection> {
    return runExcel(readSelection);
  }

  onSelectionChanged(listener: (selection: SheetSelection) => void): () => void {
    let disposed = false;
    let reading = false;
    let again = false;
    let timer: ReturnType<typeof setTimeout> | undefined;
    let unsubscribe: (() => void) | null = null;

    // Only the latest selection matters: while one read is in flight, further events
    // just ask for one more read afterwards.
    const read = async () => {
      if (disposed) return;
      if (reading) {
        again = true;
        return;
      }
      reading = true;
      try {
        const selection = await this.getSelection();
        if (!disposed) listener(selection);
      } catch {
        // No cell is selected (a chart or shape is), or Excel is busy. Keep the last one.
      } finally {
        reading = false;
        if (again) {
          again = false;
          schedule();
        }
      }
    };
    const schedule = () => {
      if (disposed) return;
      if (timer !== undefined) clearTimeout(timer);
      timer = setTimeout(read, SELECTION_SETTLE_MS);
    };

    watchSelection(schedule).then((stop) => {
      if (disposed) stop();
      else unsubscribe = stop;
    });

    return () => {
      disposed = true;
      if (timer !== undefined) clearTimeout(timer);
      if (unsubscribe) unsubscribe();
    };
  }

  onCellClicked(listener: (click: SheetClick) => void): () => void {
    let disposed = false;
    let unsubscribe: (() => void) | null = null;
    // Only the latest click matters, and only one is looked up at a time.
    let reading = false;
    let waiting: { worksheetId: string; cell: Rect } | null = null;

    const read = async (worksheetId: string, cell: Rect) => {
      if (reading) {
        waiting = { worksheetId, cell };
        return;
      }
      reading = true;
      try {
        const click = await runExcel((context) => readClick(context, worksheetId, cell));
        if (click && !disposed) listener(click);
      } catch {
        // Excel is busy or the sheet is gone: the click does nothing.
      } finally {
        reading = false;
        const next = waiting;
        waiting = null;
        if (next && !disposed) void read(next.worksheetId, next.cell);
      }
    };

    watchClicks((worksheetId, address) => {
      if (disposed || this.known.size === 0) return;
      // The top-left cell, should the address be a merged range.
      const cell = parseA1(address);
      if (cell) void read(worksheetId, { row: cell.row, col: cell.col, rows: 1, cols: 1 });
    }).then(
      (stop) => {
        if (!stop) return;
        if (disposed) stop();
        else unsubscribe = stop;
      },
      () => undefined,
    );

    return () => {
      disposed = true;
      if (unsubscribe) unsubscribe();
    };
  }

  getReportExtent(reportId: string): Promise<Extent | null> {
    return runExcel((context) => readExtent(context, reportId));
  }

  hasContentOutside(target: Extent, previous: Extent | null): Promise<boolean> {
    return runExcel((context) => hasContentOutside(context, target, previous));
  }

  writeReport(reportId: string, anchor: CellAddress, block: Block): Promise<Extent> {
    this.known.add(reportId);
    return runExcel((context) => writeBlock(context, reportId, anchor, block));
  }

  clearReport(reportId: string): Promise<void> {
    return runExcel((context) => clearBlock(context, reportId));
  }

  loadReports(): Promise<ReportDef[]> {
    return runExcel(async (context) => {
      const settings = context.workbook.settings;
      settings.load("items/key,items/value");
      await context.sync();
      const out: ReportDef[] = [];
      for (const setting of settings.items) {
        if (setting.key.indexOf(SETTING_PREFIX) !== 0) continue;
        const def = parseReportDef(setting.value, setting.key.slice(SETTING_PREFIX.length));
        if (def) out.push(def);
      }
      this.known.clear();
      for (const def of out) this.known.add(def.id);
      return out;
    });
  }

  async saveReport(def: ReportDef): Promise<void> {
    await ensureAutoShow();
    await runExcel(async (context) => {
      // `add` replaces the value when the key already exists.
      context.workbook.settings.add(settingKey(def.id), serializeReportDef(def));
      await context.sync();
    });
    this.known.add(def.id);
  }

  deleteReport(reportId: string): Promise<void> {
    this.known.delete(reportId);
    return runExcel(async (context) => {
      // Calls on a null object are no-ops, so these delete only what exists.
      context.workbook.settings.getItemOrNullObject(settingKey(reportId)).delete();
      context.workbook.names.getItemOrNullObject(reportName(reportId)).delete();
      await context.sync();
    });
  }
}
