// Office.js: report definitions in workbook settings, report locations as workbook
// names, the selection, and `ExcelHost`, which ties these and writer.ts together
// behind the `SheetHost` interface.
//
// Requirement sets: ExcelApi 1.4 for settings and names; 1.2 for
// workbook.onSelectionChanged; guarded 1.7 for getActiveCell and worksheets.onActivated.

import { parseReportDef, serializeReportDef } from "../state/report";
import { containsCell } from "./geometry";
import {
  NAME_PREFIX,
  SETTING_PREFIX,
  reportName,
  settingKey,
  type CellAddress,
  type Extent,
  type ReportDef,
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
    return () => {
      for (const handler of handlers) {
        Excel.run(handler.context, async (context) => {
          handler.remove();
          await context.sync();
        }).catch(() => undefined);
      }
    };
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

export class ExcelHost implements SheetHost {
  readonly kind = "excel" as const;

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

  getReportExtent(reportId: string): Promise<Extent | null> {
    return runExcel((context) => readExtent(context, reportId));
  }

  hasContentOutside(target: Extent, previous: Extent | null): Promise<boolean> {
    return runExcel((context) => hasContentOutside(context, target, previous));
  }

  writeReport(reportId: string, anchor: CellAddress, block: Block): Promise<Extent> {
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
  }

  deleteReport(reportId: string): Promise<void> {
    return runExcel(async (context) => {
      // Calls on a null object are no-ops, so these delete only what exists.
      context.workbook.settings.getItemOrNullObject(settingKey(reportId)).delete();
      context.workbook.names.getItemOrNullObject(reportName(reportId)).delete();
      await context.sync();
    });
  }
}
