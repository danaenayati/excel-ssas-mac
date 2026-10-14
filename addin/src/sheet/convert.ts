// Convert to Formulas, as Windows Excel has it (PivotTable Analyze > OLAP Tools >
// Convert to Formulas): the report's cells become cube formulas that read the cube
// whenever Excel recalculates, and the report itself goes away. Pure: from what the pane
// remembers about the report's last result to two block-sized matrices, the formulas
// and the number formats to write, where null leaves a cell as it is.
//
//   - A member header cell becomes =CUBE.MEMBER(conn, "<unique name>"). The member comes
//     from the block's own member map, keyed by the cell's position, so this follows
//     however the block lays its lines out.
//   - A value cell becomes =CUBE.VALUE(conn, <row members>, <column members>, <measure>,
//     <page filter members>). Its sheet line maps to the grid line it shows through the
//     block's line map (compact and outline form add lines of their own), and its column
//     is the grid's column in sheet order. Totals use the hierarchies' All members.
//   - A page filter's value cell becomes =CUBE.MEMBER(conn, <its item, or All>).
//   - Labels ("Grand Total", captions) stay as they are. Cells that get a formula and were
//     formatted as text become General, or Excel would keep the formula as text; filter
//     arrows are dropped, since there is no report left to open.
//
// A value that one cube tuple cannot give is refused with a plain sentence: Show Values
// As, KPIs, the report's calculated measures, groups and calculated items, named sets,
// and page filters with several items. Label, value and date filters only choose which
// items are shown; they are allowed, but the totals then count every item.

import type { HeaderCell, Layout, PlacedField } from "../types";
import {
  TEXT_FORMAT,
  isFilterFormat,
  valueFieldIndex,
  type BlockShape,
  type GridHeaders,
  type MemberHit,
  type SheetValue,
} from "./layout";

/** What the conversion reads: a report's last result (`ReportRuntime` in state/runner.ts). */
export interface ConversionSource {
  layout: Layout;
  headers: GridHeaders;
  shape: BlockShape;
  rows: number;
  cols: number;
  members: Map<string, MemberHit>;
  colValueIndex: number[];
  /** Per data line of the block: the grid line whose values it shows, or null for none. */
  lines: { grid: number | null }[];
  /** The values and number formats the block was written with. */
  sheet?: { values: SheetValue[][]; numberFormats: string[][] };
}

export interface Conversion {
  /** rows x cols: a formula, or null to leave the cell as it is. */
  formulas: (string | null)[][];
  /** rows x cols: a number format, or null to leave the cell's own. */
  numberFormats: (string | null)[][];
  /** Some row or column field was filtered: the totals now count all of its items. */
  filtered: boolean;
}

export const GENERAL = "General";
/** Excel's limit on the length of a formula. */
export const MAX_FORMULA = 8192;
/** Excel's limit on a text constant in a formula is 255 characters. */
const MAX_LITERAL = 250;

/** Text as an Excel string constant; a long one is joined from pieces with &. */
export function excelString(text: string): string {
  const pieces: string[] = [];
  for (let i = 0; i === 0 || i < text.length; i += MAX_LITERAL) {
    pieces.push('"' + text.slice(i, i + MAX_LITERAL).replace(/"/g, '""') + '"');
  }
  return pieces.join("&");
}

/** The connection argument: "Name/Database/Cube", which always names one cube. Null when that cannot be written. */
export function formulaConnection(name: string, catalog: string, cube: string): string | null {
  if (catalog === "" || cube === "" || catalog.indexOf("/") >= 0 || cube.indexOf("/") >= 0) return null;
  return `${name.trim()}/${catalog}/${cube}`;
}

function quoted(text: string): string {
  return '"' + text + '"';
}

/** Why this layout cannot be written as cube formulas, or null when it can. */
export function conversionBlocker(layout: Layout): string | null {
  const calculated = new Set(
    (layout.calculatedMeasures ?? []).map((m) => `[measures].[${m.name.replace(/\]/g, "]]")}]`.toLowerCase()),
  );
  for (const value of layout.values) {
    const name = quoted(value.caption);
    if (value.kpi) {
      return `${name} is part of a KPI, which a cube formula cannot ask for. Remove it from Values, then convert.`;
    }
    if ((value.showAs ?? "normal") !== "normal") {
      return `${name} is shown as a calculation (Show Values As), which a cube formula cannot do. Set it back to No Calculation, then convert.`;
    }
    if (calculated.has(value.measure.toLowerCase())) {
      return `${name} is a calculated measure of this report, which the cube does not know. Remove it from Values, then convert.`;
    }
  }
  for (const field of [...layout.rows, ...layout.columns]) {
    if (field.set) {
      return `${quoted(field.caption)} is a named set. Convert to Formulas needs the field's own items: replace the set with its field, then convert.`;
    }
    if (field.customItems && field.customItems.length > 0) {
      return `${quoted(field.caption)} has groups or calculated items, which only this report knows. Remove them, then convert.`;
    }
  }
  for (const field of layout.filters) {
    if (field.members.length > 1) {
      return `The filter ${quoted(field.caption)} has several items selected, and one cube formula takes one item of each field. Pick one item or All, then convert.`;
    }
  }
  return null;
}

/** True when an axis field shows only some of its items. */
function narrowed(field: PlacedField): boolean {
  return field.members.length > 0 || Boolean(field.labelFilter || field.valueFilter || field.dateFilter);
}

type Line = (HeaderCell | undefined)[] | undefined;

/**
 * The members of one line of an axis (a row of the grid, or a column read down its
 * header rows): its member for each field, the All member from the first total on.
 * Null when the line cannot be told.
 *
 * A subtotal line names its own field's member nowhere (its cell holds "2025 Total"),
 * so it is read from the nearest line of the same group: the one before it, or after
 * it when subtotals are at the top.
 */
function lineMembers(lineAt: (index: number) => Line, index: number, fields: PlacedField[], atTop: boolean): string[] | null {
  const line = lineAt(index);
  if (!line) return null;
  const out: string[] = [];
  for (let f = 0; f < fields.length; f++) {
    const cell = line[f];
    if (!cell) return null;
    if (cell.kind === "grandTotal" || cell.kind === "subtotal") {
      if (cell.kind === "subtotal") {
        const own = groupMember(lineAt, index, f, out, atTop);
        if (own === null) return null;
        out.push(own);
        f++;
      }
      for (let g = f; g < fields.length; g++) {
        const all = fields[g].allMember;
        if (!all) return null;
        out.push(all);
      }
      return out;
    }
    if (!cell.member) return null;
    out.push(cell.member.uniqueName);
  }
  return out;
}

function groupMember(lineAt: (index: number) => Line, index: number, field: number, outer: string[], atTop: boolean): string | null {
  const step = atTop ? 1 : -1;
  for (let j = index + step; ; j += step) {
    const line = lineAt(j);
    if (!line) return null;
    for (let g = 0; g < field; g++) {
      const cell = line[g];
      if (!cell || !cell.member || cell.member.uniqueName !== outer[g]) return null;
    }
    const cell = line[field];
    if (cell && cell.member) return cell.member.uniqueName;
    // The same subtotal again (once per measure, on columns): look further.
    if (!cell || cell.kind !== "subtotal") return null;
  }
}

function matrix<T>(rows: number, cols: number): (T | null)[][] {
  const out: (T | null)[][] = [];
  for (let r = 0; r < rows; r++) out.push(new Array<T | null>(cols).fill(null));
  return out;
}

/**
 * The formulas for a report, or the sentence saying why it cannot be converted.
 * `connection` is the formulas' first argument (see `formulaConnection`).
 */
export function convertReport(source: ConversionSource, connection: string): Conversion | { refused: string } {
  const { layout, headers, shape } = source;
  const blocker = conversionBlocker(layout);
  if (blocker) return { refused: blocker };
  const sheet = source.sheet;
  if (!sheet || sheet.values.length !== source.rows) {
    return { refused: "Refresh the report first, so that the pane knows what each of its cells shows." };
  }
  const refuse = (what: string) => ({
    refused: `Convert to Formulas could not tell which items ${what} stands for. Refresh the report and try again.`,
  });

  const conn = excelString(connection);
  const formulas = matrix<string>(source.rows, source.cols);
  const numberFormats = matrix<string>(source.rows, source.cols);
  const put = (row: number, col: number, formula: string) => {
    if (row < 0 || row >= source.rows || col < 0 || col >= source.cols) return;
    formulas[row][col] = formula;
    const format = sheet.numberFormats[row] ? sheet.numberFormats[row][col] : undefined;
    if (format === TEXT_FORMAT || (format !== undefined && isFilterFormat(format))) numberFormats[row][col] = GENERAL;
  };

  // Header cells: the member each one shows, as the block recorded it.
  source.members.forEach((hit, key) => {
    const [row, col] = key.split(",").map(Number);
    const shown = sheet.values[row] ? sheet.values[row][col] : "";
    // A repeated outer label is left blank on the sheet, and stays blank.
    if (typeof shown === "string" && shown !== "") {
      put(row, col, `=CUBE.MEMBER(${conn},${excelString(hit.member.uniqueName)})`);
    }
  });

  // Page filters: their item, or All; what every value is sliced by.
  const sliced: string[] = [];
  layout.filters.forEach((field, f) => {
    const member = field.members.length === 1 ? field.members[0] : field.allMember;
    if (!member) return; // no All member: the cube's default, as in the report
    sliced.push(member);
    if (shape.filterRows === layout.filters.length) put(f, 1, `=CUBE.MEMBER(${conn},${excelString(member)})`);
  });

  // Values. Row subtotals come before their group when the report puts them at the top
  // (compact and outline form only); column subtotals always follow their items.
  const options = layout.options;
  const rowsAtTop = options.subtotalsAt === "top" && (options.form === "compact" || options.form === "outline");
  if (source.lines.length !== shape.dataRows) return refuse("the report's lines");
  const rowTuples: (string[] | null)[] = [];
  for (let n = 0; n < shape.dataRows; n++) {
    const line = source.lines[n].grid;
    if (line === null) {
      rowTuples.push(null); // an outer item's own line that shows no values
      continue;
    }
    const members = lineMembers((i) => headers.rowHeaders[i], line, layout.rows, rowsAtTop);
    if (members === null) return refuse(`row ${n + 1} of the report`);
    rowTuples.push(members);
  }
  const columnLine = (c: number): Line =>
    c < 0 || c >= shape.dataCols ? undefined : layout.columns.map((_field, h) => (headers.colHeaderRows[h] ?? [])[c]);
  for (let c = 0; c < shape.dataCols; c++) {
    const columnMembers = lineMembers(columnLine, c, layout.columns, false);
    if (columnMembers === null) return refuse(`column ${c + 1} of the report`);
    const index = valueFieldIndex(source.colValueIndex, c, layout.values.length);
    if (index === null && layout.values.length > 0) return refuse(`column ${c + 1} of the report`);
    const measure = index === null ? [] : [layout.values[index].measure];
    for (let n = 0; n < shape.dataRows; n++) {
      const rowMembers = rowTuples[n];
      if (rowMembers === null) continue;
      const names = [...rowMembers, ...columnMembers, ...measure, ...sliced];
      const formula = `=CUBE.VALUE(${conn}${names.map((name) => "," + excelString(name)).join("")})`;
      if (formula.length > MAX_FORMULA) {
        return { refused: "Some values would need a formula longer than Excel allows. Remove a field, then convert." };
      }
      put(shape.dataTop + n, shape.headerCols + c, formula);
    }
  }

  // Filter buttons that did not get a formula keep their text, without the arrow.
  for (let r = 0; r < source.rows; r++) {
    for (let c = 0; c < source.cols; c++) {
      const format = sheet.numberFormats[r] ? sheet.numberFormats[r][c] : undefined;
      if (formulas[r][c] === null && format !== undefined && isFilterFormat(format)) numberFormats[r][c] = TEXT_FORMAT;
    }
  }

  return { formulas, numberFormats, filtered: [...layout.rows, ...layout.columns].some(narrowed) };
}
