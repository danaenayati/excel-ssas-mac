// Office.js: the charts that follow a report (sheet/charts.ts). Each one is an ordinary
// Excel chart on the report's sheet, named `CubeChart_<id>`, over a block of cells on
// the hidden sheet `CubeChartData` that the workbook name `CubeChartData_<id>` points
// at. Refreshing the report writes the block again; the chart is pointed at it anew
// only when the block has moved or changed size, so that whatever the user did to the
// chart in Excel (type, size, place, colours) stays as it is.
//
// Requirement sets: nothing above ExcelApi 1.4, which the manifest already requires.
//   1.1  worksheet.charts.add, Chart.setData / setPosition / delete, name, title, legend,
//        worksheets.add, Worksheet.visibility, Range values / numberFormat / clear
//   1.4  getItemOrNullObject (worksheets, charts, names), NamedItem.getRangeOrNullObject,
//        Worksheet.getUsedRangeOrNullObject

import {
  CHART_DATA_SHEET,
  chartAnchor,
  chartCells,
  chartDataName,
  chartName,
  dataBlockPlace,
  type ChartCells,
  type ChartKind,
  type ChartTable,
  type ReportChart,
} from "./charts";
import { cellA1, chunkRect, fitsSheet, rectA1, type Rect } from "./geometry";
import { SheetError, type ChartsUpdate, type Extent } from "./host";
import { GENERAL_FORMAT, MAX_CHUNK_CELLS, TEXT_FORMAT, isTextFormat, sliceMatrix } from "./layout";
import { hasContentOutside, readExtent, runExcel, supportsExcelApi } from "./writer";

interface NamedRect {
  sheetId: string;
  rect: Rect;
}

/** Where a workbook name points, or null when it does not exist or points at nothing. */
async function namedRect(context: Excel.RequestContext, name: string): Promise<NamedRect | null> {
  const item = context.workbook.names.getItemOrNullObject(name);
  await context.sync();
  if (item.isNullObject) return null;
  const range = item.getRangeOrNullObject();
  range.load("rowIndex,columnIndex,rowCount,columnCount,worksheet/id");
  await context.sync();
  if (range.isNullObject) return null;
  return {
    sheetId: range.worksheet.id,
    rect: { row: range.rowIndex, col: range.columnIndex, rows: range.rowCount, cols: range.columnCount },
  };
}

/** The hidden sheet with the charts' data, added (hidden) the first time it is needed. */
async function dataSheet(context: Excel.RequestContext): Promise<{ sheet: Excel.Worksheet; id: string }> {
  const found = context.workbook.worksheets.getItemOrNullObject(CHART_DATA_SHEET);
  await context.sync();
  let sheet: Excel.Worksheet = found;
  if (found.isNullObject) {
    sheet = context.workbook.worksheets.add(CHART_DATA_SHEET);
    sheet.visibility = Excel.SheetVisibility.hidden;
  }
  sheet.load("id");
  await context.sync();
  return { sheet, id: sheet.id };
}

function onSheet(sheetId: string, rect: Rect): Extent {
  return { sheetId, sheetName: CHART_DATA_SHEET, ...rect };
}

/**
 * Write a chart's cells on the hidden sheet: where its block was, unless the block has
 * grown into other cells there, else below everything on the sheet. Clears the old
 * block and points the chart's data name at the new one. `moved` is true when the
 * block is not exactly where and as big as before, so the chart must be pointed at it.
 */
async function writeDataBlock(
  context: Excel.RequestContext,
  chartId: string,
  cells: ChartCells,
  previous: NamedRect | null,
): Promise<{ range: Excel.Range; moved: boolean }> {
  const { sheet, id } = await dataSheet(context);
  const size = { rows: cells.rows, cols: cells.cols };
  const before = previous && previous.sheetId === id ? previous.rect : null;

  let blocked = false;
  if (before) {
    const there: Rect = { row: before.row, col: before.col, ...size };
    blocked = !fitsSheet(there) || (await hasContentOutside(context, onSheet(id, there), onSheet(id, before)));
  }
  let usedBottom = 0;
  if (!before || blocked) {
    const used = sheet.getUsedRangeOrNullObject();
    used.load("rowIndex,rowCount");
    await context.sync();
    usedBottom = used.isNullObject ? 0 : used.rowIndex + used.rowCount;
  }
  const rect = dataBlockPlace(before, size, blocked, usedBottom);
  if (!fitsSheet(rect)) {
    throw new SheetError(`The chart's data does not fit on the hidden sheet ${CHART_DATA_SHEET}.`);
  }

  if (previous) {
    context.workbook.worksheets.getItem(previous.sheetId).getRange(rectA1(previous.rect)).clear();
  }
  sheet.getRange(rectA1(rect)).clear();
  await context.sync();

  // Number formats first, so that names are text before they arrive.
  for (const chunk of chunkRect({ row: 0, col: 0, ...size }, MAX_CHUNK_CELLS)) {
    const address = rectA1({ row: rect.row + chunk.row, col: rect.col + chunk.col, rows: chunk.rows, cols: chunk.cols });
    const formats = sliceMatrix(cells.numberFormats, chunk);
    const values = sliceMatrix(cells.values, chunk);
    try {
      const range = sheet.getRange(address);
      range.numberFormat = formats;
      range.values = values;
      await context.sync();
    } catch {
      // Excel refuses the whole write over one number format it dislikes (cube format
      // strings do not always translate): the numbers matter more than their look.
      const range = sheet.getRange(address);
      range.numberFormat = formats.map((line) => line.map((f) => (isTextFormat(f) ? TEXT_FORMAT : GENERAL_FORMAT)));
      range.values = values;
      await context.sync();
    }
  }

  const range = sheet.getRange(rectA1(rect));
  const name = chartDataName(chartId);
  context.workbook.names.getItemOrNullObject(name).delete();
  context.workbook.names.add(name, range);
  await context.sync();

  const moved =
    !before ||
    before.row !== rect.row ||
    before.col !== rect.col ||
    before.rows !== rect.rows ||
    before.cols !== rect.cols;
  return { range, moved };
}

/**
 * The chart, on its sheet or, should the user have moved it with Move Chart, on any
 * other worksheet. Null when it is nowhere to be found: deleted, renamed, or moved to a
 * chart sheet, which Office.js cannot see.
 */
async function findChart(
  context: Excel.RequestContext,
  chart: ReportChart,
): Promise<{ item: Excel.Chart; sheetId: string } | null> {
  const name = chartName(chart.id);
  const home = context.workbook.worksheets.getItemOrNullObject(chart.sheetId);
  await context.sync();
  if (!home.isNullObject) {
    const item = home.charts.getItemOrNullObject(name);
    await context.sync();
    if (!item.isNullObject) return { item, sheetId: chart.sheetId };
  }
  const sheets = context.workbook.worksheets;
  sheets.load("items/id");
  await context.sync();
  const others = sheets.items
    .filter((sheet) => sheet.id !== chart.sheetId)
    .map((sheet) => ({ sheetId: sheet.id, item: sheet.charts.getItemOrNullObject(name) }));
  if (others.length === 0) return null;
  await context.sync();
  const hit = others.find((other) => !other.item.isNullObject);
  return hit ? { item: hit.item, sheetId: hit.sheetId } : null;
}

/** Forget a chart's data block; with `clear`, empty its cells too. */
async function dropDataBlock(context: Excel.RequestContext, chartId: string, clear: boolean): Promise<void> {
  const block = await namedRect(context, chartDataName(chartId));
  if (block && clear) context.workbook.worksheets.getItem(block.sheetId).getRange(rectA1(block.rect)).clear();
  // A call on a null object is a no-op, so this deletes the name only if it exists.
  context.workbook.names.getItemOrNullObject(chartDataName(chartId)).delete();
  await context.sync();
}

/** Insert Chart. See `SheetHost.addChart`. */
export async function addChart(
  context: Excel.RequestContext,
  reportId: string,
  chartId: string,
  type: ChartKind,
  table: ChartTable,
  index: number,
): Promise<ReportChart> {
  if (!supportsExcelApi("1.4")) {
    throw new SheetError("This version of Excel cannot add charts from an add-in. Update Excel, then try again.");
  }
  const extent = await readExtent(context, reportId);
  if (!extent) throw new SheetError("This report is no longer on a sheet. Refresh it, then try again.");
  const block = await writeDataBlock(context, chartId, chartCells(table), null);
  try {
    const chart = context.workbook.worksheets.getItem(extent.sheetId).charts.add(type, block.range, "Columns");
    chart.name = chartName(chartId);
    chart.title.text = table.title;
    chart.title.visible = true;
    chart.legend.visible = true;
    // Where Windows puts a PivotChart's legend.
    chart.legend.position = "Right";
    const at = chartAnchor(extent, index);
    chart.setPosition(cellA1(at.row, at.col));
    await context.sync();
  } catch (error) {
    try {
      await dropDataBlock(context, chartId, true);
    } catch {
      // The block stays on the hidden sheet; it does no harm.
    }
    throw error;
  }
  return { id: chartId, type, sheetId: extent.sheetId, title: table.title };
}

/** One chart after a write of its report; null when the chart is gone. */
async function updateChart(
  context: Excel.RequestContext,
  chart: ReportChart,
  cells: ChartCells,
  title: string,
): Promise<ReportChart | null> {
  const found = await findChart(context, chart);
  if (!found) {
    // Keep its cells: a chart that was only renamed or moved to a chart sheet still
    // shows them, as a picture that no longer updates.
    await dropDataBlock(context, chart.id, false);
    return null;
  }
  const { item, sheetId } = found;
  item.title.load("text");
  const previous = await namedRect(context, chartDataName(chart.id));
  const block = await writeDataBlock(context, chart.id, cells, previous);
  if (block.moved) item.setData(block.range, "Columns");
  // The title follows the report while it still reads what was last written.
  const follow = chart.title === undefined || item.title.text === chart.title;
  if (follow && item.title.text !== title) item.title.text = title;
  await context.sync();
  const out: ReportChart = { ...chart, sheetId };
  if (follow) out.title = title;
  return out;
}

/** See `SheetHost.updateCharts`. One request per chart, so that one failure stops no other. */
export async function updateChartsInExcel(charts: ReportChart[], table: ChartTable): Promise<ChartsUpdate> {
  const cells = chartCells(table);
  const kept: ReportChart[] = [];
  let error: unknown;
  for (const chart of charts) {
    try {
      const updated = await runExcel((context) => updateChart(context, chart, cells, table.title));
      if (updated) kept.push(updated);
    } catch (failure) {
      kept.push(chart);
      if (error === undefined) error = failure;
    }
  }
  return error === undefined ? { charts: kept } : { charts: kept, error };
}

/** Delete Chart. See `SheetHost.deleteChart`. */
export async function deleteChart(context: Excel.RequestContext, chart: ReportChart): Promise<void> {
  const found = await findChart(context, chart);
  if (found) found.item.delete();
  await context.sync();
  await dropDataBlock(context, chart.id, true);
}
