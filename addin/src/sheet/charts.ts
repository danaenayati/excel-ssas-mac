// Charts that follow a report, the nearest Excel for Mac gets to a PivotChart. Pure:
// what a chart of a report shows, how its data sits in cells, and how its record in
// the report's definition is read back. The hosts (excelCharts.ts, preview.ts) put it
// on a sheet; the runner (state/runner.ts) keeps it up to date.
//
// Excel for Mac cannot make a PivotChart from our reports (they are not native
// PivotTables), so each chart is an ordinary Excel chart whose data is a small table
// on a hidden sheet. After every write of the report that table is written again from
// the report's latest result, as a PivotChart follows its PivotTable.
//
// The table is what a PivotChart plots: one category per data line of the report, one
// series per data column, no subtotals and no grand totals. Where a member of a
// hierarchy is expanded, the lines under it are its categories and its own line is
// left out, as Windows treats the parent's line as the subtotal of its children; the
// children's labels name the parent ("2025 / Q1 2025").

import type { Grid, HeaderCell, Layout, LineKind } from "../types";
import { MAX_COLS, MAX_ROWS, type Rect } from "./geometry";
import type { ReportDef } from "./host";
import { GENERAL_FORMAT, TEXT_FORMAT, excelText, valueFieldIndex, type SheetValue } from "./layout";

/** Windows' PivotChart types, by Excel's own name for each (`Excel.ChartType`). */
export type ChartKind = "ColumnClustered" | "ColumnStacked" | "BarClustered" | "Line" | "Pie" | "Area";

/** In the order and words of Excel's Insert Chart. */
export const CHART_KINDS: readonly { kind: ChartKind; label: string }[] = [
  { kind: "ColumnClustered", label: "Clustered Column" },
  { kind: "ColumnStacked", label: "Stacked Column" },
  { kind: "BarClustered", label: "Clustered Bar" },
  { kind: "Line", label: "Line" },
  { kind: "Pie", label: "Pie" },
  { kind: "Area", label: "Area" },
];

export function chartKindLabel(kind: ChartKind): string {
  const found = CHART_KINDS.find((k) => k.kind === kind);
  return found ? found.label : kind;
}

/** A chart of a report, as kept in the report's definition (`ReportDef.charts`). */
export interface ReportChart {
  /** Letters and digits; names the chart (`CubeChart_<id>`) and its data (`CubeChartData_<id>`). */
  id: string;
  /** The type it was inserted as. The user may change it in Excel; it is never reset. */
  type: ChartKind;
  /** Stable id of the worksheet the chart is on. */
  sheetId: string;
  /**
   * The title last written. On refresh the title follows the report only while it still
   * reads this: a title the user has edited is left alone.
   */
  title?: string;
}

/** What a chart shows: the report's latest result without its totals. */
export interface ChartTable {
  /** The value fields' captions, as the report heads them. */
  title: string;
  /** One per series: a data column of the report. */
  series: string[];
  /** One per category: a data line of the report. */
  categories: string[];
  /** categories × series. Null where the report has no number (a gap in the chart). */
  values: (number | null)[][];
  /** categories × series: the number format the report shows each value in. */
  numberFormats: string[][];
}

/** Excel refuses charts with more series than this (Windows says so for a PivotChart too). */
export const MAX_CHART_SERIES = 255;
/** The hidden sheet holding every chart's data. */
export const CHART_DATA_SHEET = "CubeChartData";
export const CHART_NAME_PREFIX = "CubeChart_";
export const CHART_DATA_PREFIX = "CubeChartData_";
/** Joins the labels of the fields of one axis, outer to inner. */
export const LABEL_SEPARATOR = " / ";
/** What Windows calls the one series or category of an axis that has no fields. */
export const TOTAL_LABEL = "Total";
/** Windows' name for an item whose caption is empty. */
export const BLANK_LABEL = "(blank)";

/** Name of the Excel chart object. */
export function chartName(chartId: string): string {
  return CHART_NAME_PREFIX + chartId;
}

/** Workbook name of the chart's data block on the hidden sheet. */
export function chartDataName(chartId: string): string {
  return CHART_DATA_PREFIX + chartId;
}

// ---------------------------------------------------------------------------
// The table
// ---------------------------------------------------------------------------

interface AxisEntry {
  /** A data line (or column) that the chart plots. */
  keep: boolean;
  /** The labels of its fields, outer to inner, joined; "" when the axis has no fields. */
  label: string;
}

function captionOf(cell: HeaderCell | undefined): string {
  if (!cell) return "";
  // A repeated outer label is sent blank, with its member: the chart needs it all the same.
  return cell.member ? cell.member.caption : cell.text;
}

/**
 * Which lines of one axis a chart plots, and their labels. `lines` holds one entry per
 * line of the axis (a row of the report, or a column), each with one header cell per
 * field, outermost first.
 *
 * Totals are left out. So is the line of a member that is expanded and has its children
 * shown under it: a PivotChart plots the children, labelled with the parent, and treats
 * the parent's own line as their subtotal. Children always follow their parent on an
 * axis, one level deeper, so the parents of a line are the expanded members met before
 * it, at a lower level, under the same outer items.
 */
function axisEntries(lines: HeaderCell[][], kinds: readonly LineKind[]): AxisEntry[] {
  const width = lines.reduce((most, line) => Math.max(most, line.length), 0);
  /** Per field: the expanded members above the current line, outermost first. */
  const stacks: { name: string; level: number; caption: string }[][] = [];
  /** Per field: the outer items the stack belongs to. */
  const scopes: (string | null)[] = [];
  /** "field|outer items|member" of the members that have children shown. */
  const parents = new Set<string>();
  const keyOf = (field: number, scope: string, name: string) => field + "\u0002" + scope + "\u0002" + name;

  const seen: { own: string[]; path: string[] }[] = [];
  lines.forEach((line, i) => {
    const own: string[] = [];
    const path: string[] = [];
    if (kinds[i] === "data") {
      let scope = "";
      for (let r = 0; r < width; r++) {
        const cell = line[r];
        const member = cell ? cell.member : null;
        if (scopes[r] !== scope) {
          stacks[r] = [];
          scopes[r] = scope;
        }
        if (member) {
          const stack = stacks[r];
          while (stack.length > 0 && stack[stack.length - 1].level >= member.level) stack.pop();
          for (const parent of stack) {
            parents.add(keyOf(r, scope, parent.name));
            path.push(parent.caption);
          }
          if (member.expanded) stack.push({ name: member.uniqueName, level: member.level, caption: member.caption });
          own.push(keyOf(r, scope, member.uniqueName));
        }
        path.push(captionOf(cell));
        scope += "\u0001" + (member ? member.uniqueName : captionOf(cell));
      }
    }
    seen.push({ own, path });
  });

  return seen.map(({ own, path }, i) => {
    const parts = path.filter((part) => part !== "");
    const label = parts.length > 0 ? parts.join(LABEL_SEPARATOR) : path.length > 0 ? BLANK_LABEL : "";
    return { keep: kinds[i] === "data" && !own.some((key) => parents.has(key)), label };
  });
}

/**
 * The table a chart of the report plots, from the report's latest result. `grid` must
 * have its columns in sheet order (`Block.grid`), so that the series follow the report's
 * columns; `layout` is the layout that produced it.
 *
 *   - one category per data line: the row labels joined, outer to inner ("Iran / Tehran");
 *     "Total" when the report has no row fields;
 *   - one series per data column: the column labels joined, plus the value field's
 *     name when there are several value fields; "Total" when there is neither. A value
 *     field's name is its Custom Name if the user gave it one, else the report's header
 *     for it (which includes Show Values As); the title lists these names;
 *   - subtotal and grand total lines and columns are left out, and so are the lines and
 *     columns of expanded members whose children are shown;
 *   - a value that is not a number (empty, or an error text) is a gap.
 */
export function chartTable(grid: Grid, layout: Pick<Layout, "values">): ChartTable {
  const H = grid.colHeaderRows.length;
  const C = grid.colKinds.length;
  // Every header row but the last is a column field's; the last names the value fields.
  const fieldRows = Math.max(0, H - 1);
  const measureRow = H > 0 ? grid.colHeaderRows[H - 1] : [];
  const several = layout.values.length > 1;

  const columnLines: HeaderCell[][] = [];
  for (let c = 0; c < C; c++) {
    const line: HeaderCell[] = [];
    for (let h = 0; h < fieldRows; h++) line.push(grid.colHeaderRows[h][c]);
    columnLines.push(line);
  }
  const columns = axisEntries(columnLines, grid.colKinds);
  const rows = axisEntries(grid.rowHeaders, grid.rowKinds);

  // A value field's name: the user's own (Custom Name), else the report's header text.
  const customNames = layout.values.map((value) => (value.customName && value.customName.trim() !== "" ? value.customName : null));
  const measureName = (c: number): string => {
    const index = valueFieldIndex(grid.colValueIndex, c, customNames.length);
    const custom = index === null ? null : customNames[index];
    return custom ?? (measureRow[c] ? measureRow[c].text : "");
  };

  const kept: number[] = [];
  const series: string[] = [];
  const titles: string[] = [];
  columns.forEach((entry, c) => {
    if (!entry.keep) return;
    kept.push(c);
    const measure = measureName(c);
    if (measure !== "" && titles.indexOf(measure) < 0) titles.push(measure);
    const name = several && measure !== "" ? (entry.label !== "" ? entry.label + LABEL_SEPARATOR + measure : measure) : entry.label;
    series.push(name !== "" ? name : TOTAL_LABEL);
  });

  const userFormats = layout.values.map((value) => value.numberFormat || null);
  const formatOf = (line: number, c: number): string => {
    const index = valueFieldIndex(grid.colValueIndex, c, userFormats.length);
    const own = index === null ? null : userFormats[index];
    if (own) return own;
    const formats = grid.cellFormats[line];
    const palette = formats ? grid.formatPalette[formats[c]] : undefined;
    return palette ?? GENERAL_FORMAT;
  };

  const categories: string[] = [];
  const values: (number | null)[][] = [];
  const numberFormats: string[][] = [];
  rows.forEach((entry, g) => {
    if (!entry.keep) return;
    categories.push(entry.label !== "" ? entry.label : TOTAL_LABEL);
    const line = grid.values[g] ?? [];
    values.push(kept.map((c) => (typeof line[c] === "number" && isFinite(line[c] as number) ? (line[c] as number) : null)));
    numberFormats.push(kept.map((c) => formatOf(g, c)));
  });

  const captions =
    titles.length > 0
      ? titles
      : layout.values.map((value, i) => customNames[i] ?? value.caption).filter((text) => text !== "");
  return {
    title: captions.length > 0 ? captions.join(", ") : TOTAL_LABEL,
    series,
    categories,
    values,
    numberFormats,
  };
}

/** The data block of a chart: what goes in its cells. */
export interface ChartCells {
  rows: number;
  cols: number;
  /** rows × cols. Text is already guarded (`excelText`). */
  values: SheetValue[][];
  numberFormats: string[][];
}

/**
 * The table as cells: an empty corner, the series names across the top, the categories
 * down the left, the numbers in between. The corner stays empty so that Excel takes the
 * first row and column as names, even when the categories look like numbers ("2025");
 * names are written as text for the same reason. A table with no category or no series
 * still gets one empty one, so that the chart has a range to point at.
 */
export function chartCells(table: ChartTable): ChartCells {
  const series = table.series.length > 0 ? table.series : [""];
  const categories = table.categories.length > 0 ? table.categories : [""];
  const rows = 1 + categories.length;
  const cols = 1 + series.length;
  const values: SheetValue[][] = [];
  const numberFormats: string[][] = [];
  values.push(["", ...series.map(excelText)]);
  numberFormats.push([GENERAL_FORMAT, ...series.map(() => TEXT_FORMAT)]);
  categories.forEach((category, n) => {
    const line = table.values[n] ?? [];
    const formats = table.numberFormats[n] ?? [];
    const valueLine: SheetValue[] = [excelText(category)];
    const formatLine: string[] = [TEXT_FORMAT];
    for (let s = 0; s < series.length; s++) {
      const value = line[s];
      valueLine.push(value === null || value === undefined ? "" : value);
      formatLine.push(formats[s] || GENERAL_FORMAT);
    }
    values.push(valueLine);
    numberFormats.push(formatLine);
  });
  return { rows, cols, values, numberFormats };
}

/** Why a chart cannot show this table, or null when it can. */
export function chartRefusal(table: ChartTable): string | null {
  if (table.series.length <= MAX_CHART_SERIES) return null;
  return (
    `A chart can show at most ${MAX_CHART_SERIES} series, and this report has ` +
    `${table.series.length.toLocaleString("en-US")} columns. Move a field from Columns to Rows, or filter it.`
  );
}

// ---------------------------------------------------------------------------
// Placement
// ---------------------------------------------------------------------------

/**
 * The cell a new chart's top-left corner goes to: beside the report, one empty column to
 * its right, level with its first row. A report's further charts step down and right a
 * little each, so that none hides another.
 */
export function chartAnchor(report: Rect, index: number): { row: number; col: number } {
  const step = Math.max(0, index);
  return {
    row: Math.min(report.row + step * 2, MAX_ROWS - 1),
    col: Math.min(report.col + report.cols + 1 + step, MAX_COLS - 1),
  };
}

/**
 * Where a chart's data block goes on the hidden sheet: where it was, unless it has grown
 * into cells that hold something else (another chart's data), in which case below
 * everything on the sheet, after an empty row. `usedBottom` is the row after the last
 * used row (0 for an empty sheet).
 */
export function dataBlockPlace(
  previous: Rect | null,
  size: { rows: number; cols: number },
  blocked: boolean,
  usedBottom: number,
): Rect {
  if (previous && !blocked) return { row: previous.row, col: previous.col, ...size };
  return { row: usedBottom > 0 ? usedBottom + 1 : 0, col: 0, ...size };
}

// ---------------------------------------------------------------------------
// Reading the definition back. As for the rest of a stored definition (state/report.ts),
// nothing is trusted to have the right shape; what cannot be read is dropped.
// ---------------------------------------------------------------------------

const CHART_ID = /^[A-Za-z0-9]{1,40}$/;
/** More charts than this on one report is not something a person made. */
export const MAX_CHARTS = 50;
const MAX_TITLE = 255;

/** `ReportDef.charts` from stored JSON. Unreadable entries and repeated ids are dropped. */
export function parseCharts(value: unknown): ReportChart[] {
  if (!Array.isArray(value)) return [];
  const out: ReportChart[] = [];
  for (const item of value) {
    if (out.length >= MAX_CHARTS) break;
    if (typeof item !== "object" || item === null || Array.isArray(item)) continue;
    const record = item as Record<string, unknown>;
    const { id, type, sheetId, title } = record;
    if (typeof id !== "string" || !CHART_ID.test(id) || out.some((chart) => chart.id === id)) continue;
    if (!CHART_KINDS.some((k) => k.kind === type)) continue;
    if (typeof sheetId !== "string" || sheetId === "" || sheetId.length > 200) continue;
    const chart: ReportChart = { id, type: type as ChartKind, sheetId };
    if (typeof title === "string" && title.length <= MAX_TITLE) chart.title = title;
    out.push(chart);
  }
  return out;
}

/** The charts as stored: only the fields above, so nothing else leaks into the workbook. */
export function serializeCharts(charts: readonly ReportChart[]): ReportChart[] {
  return charts.map((chart) => {
    const out: ReportChart = { id: chart.id, type: chart.type, sheetId: chart.sheetId };
    if (chart.title !== undefined) out.title = chart.title;
    return out;
  });
}

/** The definition with these charts; without the `charts` field when there are none. */
export function withCharts(def: ReportDef, charts: readonly ReportChart[]): ReportDef {
  const out: ReportDef = { ...def };
  if (charts.length > 0) out.charts = charts.slice();
  else delete out.charts;
  return out;
}

/** Same charts, same order, same fields: nothing to save or show anew. */
export function sameCharts(a: readonly ReportChart[] | undefined, b: readonly ReportChart[] | undefined): boolean {
  return JSON.stringify(serializeCharts(a ?? [])) === JSON.stringify(serializeCharts(b ?? []));
}
