// Office.js: writing, clearing and formatting a report block.
//
// Requirement sets. Everything here is ExcelApi 1.4 or lower unless guarded:
//   1.1  getRange, values, numberFormat, formulas, clear, format.font/fill
//   1.2  format.autofitColumns
//   1.4  names.add / getItemOrNullObject / NamedItem.delete / getRangeOrNullObject,
//        worksheet.getUsedRangeOrNullObject
//   1.9  (guarded) suspendScreenUpdatingUntilNextSync, worksheet.getRanges (RangeAreas),
//        format.indentLevel. Without 1.9, formatting is applied range by range and
//        indent levels become leading spaces.
// Ranges are always addressed in A1 notation, because getRangeByIndexes is 1.7.

import {
  batchAddresses,
  cellA1,
  chunkRect,
  intersect,
  offsetRect,
  rectA1,
  subtract,
  type Rect,
} from "./geometry";
import { SheetError, detailsSheetName, reportName, type CellAddress, type Extent } from "./host";
import type { CellValue } from "../types";
import {
  GENERAL_FORMAT,
  MAX_CHUNK_CELLS,
  TEXT_FORMAT,
  chunkBlock,
  detailsCells,
  excelText,
  excelValues,
  isTextFormat,
  sliceMatrix,
  type Block,
} from "./layout";
import { EDGES, type Edge, type Paint } from "./styles";

/** Font colour of the placeholder of an empty report. */
export const MUTED_COLOR = "#7F7F7F";

/** Longest address string passed to `worksheet.getRanges()`. Kept short on purpose:
 *  Excel's own limit for a multi-area reference string is 255 characters. */
const ADDRESS_BATCH_LENGTH = 250;

export function supportsExcelApi(version: string): boolean {
  try {
    return Office.context.requirements.isSetSupported("ExcelApi", version);
  } catch {
    return false;
  }
}

interface OfficeErrorLike {
  code?: string;
  message?: string;
  debugInfo?: { errorLocation?: string; statement?: string };
}

/** Turn whatever Office.js threw into a sentence for the user plus the raw detail. */
export function toSheetError(error: unknown): SheetError {
  if (error instanceof SheetError) return error;
  const e = (error ?? {}) as OfficeErrorLike;
  const code = typeof e.code === "string" ? e.code : "";
  const parts = [code, e.message, e.debugInfo && e.debugInfo.errorLocation].filter(
    (part): part is string => typeof part === "string" && part !== "",
  );
  const detail = parts.length > 0 ? parts.join(" · ") : String(error);

  let message: string;
  switch (code) {
    case "InvalidOperationInCellEditMode":
      message = "Excel is waiting for you to finish editing a cell. Press Return, then try again.";
      break;
    case "AccessDenied":
    case "UnsupportedOperation":
      message = "Excel did not allow that change. The sheet may be protected.";
      break;
    case "ItemNotFound":
      message = "The sheet this report was on can no longer be found.";
      break;
    case "RequestPayloadSizeLimitExceeded":
    case "ResponsePayloadSizeLimitExceeded":
      message = "This view is too large for Excel to take in one go. Add a filter or remove a field.";
      break;
    default:
      message = "Excel could not complete this. Try again.";
  }
  return new SheetError(message, detail);
}

/** `Excel.run` with errors mapped to `SheetError`. */
export async function runExcel<T>(batch: (context: Excel.RequestContext) => Promise<T>): Promise<T> {
  try {
    return await Excel.run(batch);
  } catch (error) {
    throw toSheetError(error);
  }
}

/**
 * Where the report's name points now. Null when the name does not exist or no longer
 * refers to cells (its sheet was deleted).
 */
export async function readExtent(
  context: Excel.RequestContext,
  reportId: string,
): Promise<Extent | null> {
  const item = context.workbook.names.getItemOrNullObject(reportName(reportId));
  await context.sync();
  if (item.isNullObject) return null;
  const range = item.getRangeOrNullObject();
  range.load("rowIndex,columnIndex,rowCount,columnCount,worksheet/id,worksheet/name");
  await context.sync();
  if (range.isNullObject) return null;
  return {
    sheetId: range.worksheet.id,
    sheetName: range.worksheet.name,
    row: range.rowIndex,
    col: range.columnIndex,
    rows: range.rowCount,
    cols: range.columnCount,
  };
}

function sameSheet(a: CellAddress, b: CellAddress): boolean {
  return a.sheetId === b.sheetId;
}

/** Point `CubeReport_<id>` at a range, replacing whatever it pointed at before. */
function setName(context: Excel.RequestContext, reportId: string, range: Excel.Range): void {
  const name = reportName(reportId);
  // A call on a null object is a no-op, so this deletes the name only if it exists.
  context.workbook.names.getItemOrNullObject(name).delete();
  context.workbook.names.add(name, range);
}

function applyFormat(
  sheet: Excel.Worksheet,
  origin: { row: number; col: number },
  rects: Rect[],
  useAreas: boolean,
  apply: (format: Excel.RangeFormat) => void,
): void {
  if (rects.length === 0) return;
  const addresses = rects.map((rect) => rectA1(offsetRect(rect, origin.row, origin.col)));
  if (useAreas) {
    for (const batch of batchAddresses(addresses, ADDRESS_BATCH_LENGTH)) {
      apply(sheet.getRanges(batch).format);
    }
  } else {
    for (const address of addresses) apply(sheet.getRange(address).format);
  }
}

// Plain strings rather than Excel.BorderIndex: this module is loaded outside Excel too,
// where the Excel global does not exist.
const EXCEL_EDGES: Record<Edge, "EdgeTop" | "EdgeBottom" | "EdgeLeft" | "EdgeRight" | "InsideHorizontal" | "InsideVertical"> = {
  top: "EdgeTop",
  bottom: "EdgeBottom",
  left: "EdgeLeft",
  right: "EdgeRight",
  insideH: "InsideHorizontal",
  insideV: "InsideVertical",
};

/** One layer of the report's style: fill, font colour and lines (ExcelApi 1.1). */
function applyPaint(format: Excel.RangeFormat, paint: Paint): void {
  if (paint.fill) format.fill.color = paint.fill;
  if (paint.color) format.font.color = paint.color;
  if (!paint.borders) return;
  for (const edge of EDGES) {
    const line = paint.borders[edge];
    if (!line) continue;
    const border = format.borders.getItem(EXCEL_EDGES[edge]);
    border.color = line.color;
    if (line.style === "double") {
      border.style = "Double";
    } else {
      border.style = "Continuous";
      border.weight = "Thin";
    }
  }
}

/** Excel's KPI icons: -1 red / down, 0 yellow / level, 1 green / up; the icon only. */
function addKpiIcons(range: Excel.Range, kind: "status" | "trend"): void {
  const format = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
  const icons = format.iconSet;
  icons.style = kind === "status" ? Excel.IconSet.threeTrafficLights1 : Excel.IconSet.threeArrows;
  icons.showIconOnly = true;
  icons.criteria = [
    {} as Excel.ConditionalIconCriterion,
    { type: Excel.ConditionalFormatIconRuleType.number, operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual, formula: "=-0.5" },
    { type: Excel.ConditionalFormatIconRuleType.number, operator: Excel.ConditionalIconCriterionOperator.greaterThanOrEqual, formula: "=0.5" },
  ];
}

export async function writeBlock(
  context: Excel.RequestContext,
  reportId: string,
  anchor: CellAddress,
  block: Block,
): Promise<Extent> {
  const modern = supportsExcelApi("1.9");
  const suspend = () => {
    if (modern) context.application.suspendScreenUpdatingUntilNextSync();
  };

  const previous = await readExtent(context, reportId);
  const sheet = context.workbook.worksheets.getItem(anchor.sheetId);
  const target: Rect = { row: anchor.row, col: anchor.col, rows: block.rows, cols: block.cols };
  const targetRange = sheet.getRange(rectA1(target));

  // 1. Clear the old extent (wherever it is now) and the new one: values and formats.
  suspend();
  if (previous) {
    context.workbook.worksheets.getItem(previous.sheetId).getRange(rectA1(previous)).clear();
  }
  targetRange.clear();
  await context.sync();
  if (supportsExcelApi("1.6")) {
    // KPI icons from the last write are conditional formats, which clear() may leave.
    try {
      if (previous) context.workbook.worksheets.getItem(previous.sheetId).getRange(rectA1(previous)).conditionalFormats.clearAll();
      targetRange.conditionalFormats.clearAll();
      await context.sync();
    } catch {
      // Nothing to clear, or not allowed: the new write goes on regardless.
    }
  }

  // 2. Number formats, then values, one chunk per request. Formats go first so that
  //    text cells are already "Text" when their value arrives.
  const values = excelValues(block, !modern);
  for (const chunk of chunkBlock(block, MAX_CHUNK_CELLS)) {
    const address = rectA1(offsetRect(chunk, anchor.row, anchor.col));
    const formats = sliceMatrix(block.numberFormats, chunk);
    const chunkValues = sliceMatrix(values, chunk);
    try {
      suspend();
      const range = sheet.getRange(address);
      range.numberFormat = formats;
      range.values = chunkValues;
      await context.sync();
    } catch {
      // Excel refuses the whole write when it dislikes one number format (cube format
      // strings do not always translate). The numbers matter more than their display:
      // write this chunk again with plain formats, keeping text cells as text. Filter
      // cells are text too; they lose their arrow here but keep their value.
      suspend();
      const range = sheet.getRange(address);
      range.numberFormat = formats.map((line) =>
        line.map((format) => (isTextFormat(format) ? TEXT_FORMAT : GENERAL_FORMAT)),
      );
      range.values = chunkValues;
      await context.sync();
    }
  }

  // 3. Formatting, column widths and the name, in one request. The style's layers go
  //    first, in order, each one batched over all its cells.
  const format = (withStyle: boolean) => {
    suspend();
    if (withStyle) {
      for (const layer of block.paint) {
        applyFormat(sheet, anchor, layer.rects, modern, (f) => applyPaint(f, layer.paint));
      }
    }
    applyFormat(sheet, anchor, block.bold, modern, (f) => {
      f.font.bold = true;
    });
    applyFormat(sheet, anchor, block.muted, modern, (f) => {
      f.font.color = MUTED_COLOR;
      f.font.italic = true;
    });
    if (modern) {
      for (const group of block.indents) {
        applyFormat(sheet, anchor, group.rects, true, (f) => {
          f.indentLevel = group.level;
        });
      }
    }
    if (block.autofit) targetRange.format.autofitColumns();
    setName(context, reportId, targetRange);
  };
  try {
    format(true);
    await context.sync();
  } catch (error) {
    // Only a look: should Excel refuse a colour or a line, the report keeps its numbers,
    // its bold and its name. Without a style to drop, the error is real.
    if (block.paint.length === 0) throw error;
    format(false);
    await context.sync();
  }

  // 4. KPI Status and Trend as icons, as Excel shows them. Only a look: a refusal leaves
  //    the numbers (-1, 0, 1) in place.
  if (block.icons && block.icons.length > 0 && supportsExcelApi("1.6")) {
    try {
      for (const icon of block.icons) {
        addKpiIcons(sheet.getRange(rectA1(offsetRect(icon.rect, anchor.row, anchor.col))), icon.kind);
      }
      await context.sync();
    } catch {
      // Icon sets refused (an older Excel): the numbers stay.
    }
  }

  return { sheetId: anchor.sheetId, sheetName: anchor.sheetName, ...target };
}

/** Empty the report's cells; keep its name on the top-left cell. */
export async function clearBlock(context: Excel.RequestContext, reportId: string): Promise<void> {
  const extent = await readExtent(context, reportId);
  if (!extent) return;
  const sheet = context.workbook.worksheets.getItem(extent.sheetId);
  if (supportsExcelApi("1.9")) context.application.suspendScreenUpdatingUntilNextSync();
  sheet.getRange(rectA1(extent)).clear();
  setName(context, reportId, sheet.getRange(cellA1(extent.row, extent.col)));
  await context.sync();
}

/**
 * True when `target` covers a non-empty cell outside `previous`. Reads only the part of
 * `target` that overlaps the sheet's used range, in chunks, and stops at the first hit.
 */
export async function hasContentOutside(
  context: Excel.RequestContext,
  target: Extent,
  previous: Extent | null,
): Promise<boolean> {
  const sheet = context.workbook.worksheets.getItem(target.sheetId);
  const used = sheet.getUsedRangeOrNullObject(true);
  used.load("rowIndex,columnIndex,rowCount,columnCount");
  await context.sync();
  if (used.isNullObject) return false;
  const usedRect: Rect = {
    row: used.rowIndex,
    col: used.columnIndex,
    rows: used.rowCount,
    cols: used.columnCount,
  };

  const outside = subtract(target, previous && sameSheet(previous, target) ? previous : null);
  for (const part of outside) {
    const overlap = intersect(part, usedRect);
    if (!overlap) continue;
    for (const chunk of chunkRect(overlap, MAX_CHUNK_CELLS)) {
      const range = sheet.getRange(rectA1(chunk));
      // `formulas` is "" for an empty cell and non-empty for both constants and formulas.
      range.load("formulas");
      await context.sync();
      for (const line of range.formulas as unknown[][]) {
        for (const cell of line) {
          if (cell !== "" && cell !== null && cell !== undefined) return true;
        }
      }
    }
  }
  return false;
}

/** Rows of a details sheet that column widths are fitted to; enough to judge them. */
const AUTOFIT_ROWS = 1000;

/**
 * Show Details: add a sheet with a bold header row and the rows, written in chunks,
 * columns fitted, top row frozen (ExcelApi 1.7, guarded), and make it the active
 * sheet. Resolves with the sheet's name.
 */
export async function writeDetailsSheet(
  context: Excel.RequestContext,
  columns: string[],
  rows: CellValue[][],
): Promise<string> {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();
  const name = detailsSheetName(sheets.items.map((sheet) => sheet.name));
  const sheet = sheets.add(name);

  const { values, numberFormats, cols } = detailsCells(columns, rows);
  const guarded = values.map((line, r) =>
    line.map((value, c) => (typeof value === "string" && numberFormats[r][c] === TEXT_FORMAT ? excelText(value) : value)),
  );
  for (const chunk of chunkRect({ row: 0, col: 0, rows: values.length, cols }, MAX_CHUNK_CELLS)) {
    const range = sheet.getRange(rectA1(chunk));
    range.numberFormat = sliceMatrix(numberFormats, chunk);
    range.values = sliceMatrix(guarded, chunk);
    await context.sync();
  }

  sheet.getRange(rectA1({ row: 0, col: 0, rows: 1, cols })).format.font.bold = true;
  sheet
    .getRange(rectA1({ row: 0, col: 0, rows: Math.min(values.length, AUTOFIT_ROWS), cols }))
    .format.autofitColumns();
  sheet.activate();
  await context.sync();

  if (supportsExcelApi("1.7")) {
    try {
      sheet.freezePanes.freezeRows(1);
      await context.sync();
    } catch {
      // The rows are there; only the header does not stay in view when scrolling.
    }
  }
  return name;
}
