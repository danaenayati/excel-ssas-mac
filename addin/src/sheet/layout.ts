// Grid -> sheet block. Pure: turns a query result into everything a host needs to
// write it (cell values, number formats, formatting rectangles) plus the lookup from
// a cell position to the cube member behind it.
//
// Block layout, top to bottom (see `Grid` in types.ts):
//   F rows   caption | value            page filters
//   1 row    blank                      only when F > 0
//   H rows   R corner cells | C header cells   (last row's corner = row field captions)
//   N rows   R row-header cells | C values
// All positions in a Block are relative to its top-left cell.

import { isFiltered } from "../state/layout";
import type { CellValue, Grid, HeaderCell, Layout, MemberRef } from "../types";
import { chunkRect, runs, type Rect } from "./geometry";

/** Excel's "Text" number format: stops Excel from reading "2024" or "March 2025" as a number or date. */
export const TEXT_FORMAT = "@";
export const GENERAL_FORMAT = "General";
/** Most cells written or formatted in one request. */
export const MAX_CHUNK_CELLS = 20000;
/** Spaces per indent level when the host cannot indent with formatting. */
export const INDENT_SPACES = 3;
/**
 * Number format of a filter cell: the text, then a small arrow. The arrow is drawn by
 * the format and is not part of the cell's value, so "Years" stays "Years". Only text
 * shows it, so the hosts force these cells to text (see `excelValues`).
 */
export const FILTER_FORMAT = '@" ▾"';
/**
 * The same for a row or column field that is filtered (items, label or value filter):
 * a solid arrow, so a narrowed field stands out as Excel's funnel button does. Page
 * filters keep the plain arrow; their value cell already says what is chosen.
 */
export const FILTERED_FORMAT = '@" ▼"';
/** What a report with no fields shows at its place on the sheet. */
export const PLACEHOLDER_TEXT = "Cube report: tick fields in the Cube Explorer panel to build it";

export type Axis = "rows" | "columns";
export type SheetValue = string | number | boolean;

export interface MemberHit {
  member: MemberRef;
  axis: Axis;
}

export interface BlockShape {
  /** F */
  filterRows: number;
  /** H */
  headerRows: number;
  /** R */
  headerCols: number;
  /** N */
  dataRows: number;
  /** C */
  dataCols: number;
  /** Row of the first column-header row. */
  headerTop: number;
  /** Row of the first data row. */
  dataTop: number;
}

export interface IndentGroup {
  level: number;
  rects: Rect[];
}

/**
 * A filter cell: a cell that stands for a placed field and opens its checklist when
 * clicked, as the page-filter and "Row Labels" buttons of a PivotTable do.
 */
export interface FieldCell {
  area: "filters" | "rows" | "columns";
  /** Position of the field in that area of the layout. */
  index: number;
}

/** What `buildBlock` needs to know about the layout: the grid itself names no fields. */
export interface BlockFields {
  /** Number of hierarchies in Filters; filter row i belongs to filter i. */
  filters: number;
  /** Number of hierarchies in Rows; corner caption r belongs to row field r. */
  rows: number;
  /** Captions of the hierarchies in Columns; header row i belongs to column field i. */
  columns: string[];
  /** Which row fields are filtered (their caption cell gets the solid arrow). */
  rowsFiltered?: boolean[];
  /** Which column fields are filtered. */
  columnsFiltered?: boolean[];
  /**
   * The user's number format for each value field, in Values order; null keeps the
   * cube's. Applied to every value cell of that field.
   */
  valueFormats?: (string | null)[];
}

export function blockFields(
  layout: Pick<Layout, "filters" | "rows" | "columns"> & { values?: Layout["values"] },
): BlockFields {
  return {
    filters: layout.filters.length,
    rows: layout.rows.length,
    columns: layout.columns.map((field) => field.caption),
    rowsFiltered: layout.rows.map(isFiltered),
    columnsFiltered: layout.columns.map(isFiltered),
    valueFormats: (layout.values ?? []).map((value) => value.numberFormat || null),
  };
}

/**
 * Which value field a value column belongs to, as the helper reported it in
 * `Grid.colValueIndex`. Null for the cube's default measure, an index out of range, or
 * a grid without the list.
 */
export function valueFieldIndex(
  colValueIndex: readonly number[] | undefined,
  col: number,
  valueFields: number,
): number | null {
  const index = colValueIndex ? colValueIndex[col] : undefined;
  return index !== undefined && index >= 0 && index < valueFields ? index : null;
}

export interface Block {
  rows: number;
  cols: number;
  /** rows x cols. Empty cells are "". */
  values: SheetValue[][];
  /** rows x cols Excel number formats. */
  numberFormats: string[][];
  bold: Rect[];
  fill: Rect[];
  /** Cells drawn as a quiet note (grey italic): the placeholder of an empty report. */
  muted: Rect[];
  /** Cells to indent, grouped by level (1 or more). */
  indents: IndentGroup[];
  /** Field behind a filter cell, keyed by `memberKey(row, col)`. */
  fieldCells: Map<string, FieldCell>;
  /** Whether the host should fit the column widths to the block after writing. */
  autofit: boolean;
  /** Member behind a header cell, keyed by `memberKey(row, col)`. */
  members: Map<string, MemberHit>;
  shape: BlockShape;
}

export function memberKey(row: number, col: number): string {
  return row + "," + col;
}

export function memberAt(block: Pick<Block, "members">, row: number, col: number): MemberHit | null {
  return block.members.get(memberKey(row, col)) ?? null;
}

function toSheetValue(value: CellValue | undefined): SheetValue {
  return value === null || value === undefined ? "" : value;
}

function indentOf(cell: HeaderCell | undefined): number {
  const n = cell ? Math.floor(cell.indent) : 0;
  return n > 0 ? n : 0;
}

/** A filter cell's format, with either arrow. */
export function isFilterFormat(format: string): boolean {
  return format === FILTER_FORMAT || format === FILTERED_FORMAT;
}

export function isTextFormat(format: string): boolean {
  return format === TEXT_FORMAT || isFilterFormat(format);
}

export function fieldCellAt(
  block: Pick<Block, "fieldCells">,
  row: number,
  col: number,
): FieldCell | null {
  return block.fieldCells.get(memberKey(row, col)) ?? null;
}

/**
 * `fields` describes the layout the grid was asked for. With it the block marks its
 * filter cells and gives each column field a caption cell, which the grid lacks:
 *
 *   - filter row i: caption and value cell stand for filter i; the value cell shows the arrow;
 *   - last header row: corner cell r is the caption of row field r, with the arrow;
 *   - header row i (column field i): the corner cell nearest the data gets that field's
 *     caption, with the arrow.
 *
 * Without `fields` the block is the bare grid: no filter cells, no column captions.
 */
export function buildBlock(grid: Grid, fields: BlockFields = { filters: 0, rows: 0, columns: [] }): Block {
  const F = grid.filterRows.length;
  const H = grid.colHeaderRows.length;
  const R = grid.rowFieldCaptions.length;
  const N = grid.rowHeaders.length;
  const C = grid.colKinds.length;

  const headerTop = F > 0 ? F + 1 : 0;
  const dataTop = headerTop + H;
  const rows = dataTop + N;
  // A filter row needs two cells even if the grid itself is one column wide.
  const cols = Math.max(R + C, F > 0 ? 2 : 1);

  const values: SheetValue[][] = [];
  const numberFormats: string[][] = [];
  for (let r = 0; r < rows; r++) {
    values.push(new Array<SheetValue>(cols).fill(""));
    numberFormats.push(new Array<string>(cols).fill(GENERAL_FORMAT));
  }

  const members = new Map<string, MemberHit>();
  const bold: Rect[] = [];
  const fill: Rect[] = [];
  const indentRects = new Map<number, Rect[]>();
  const addIndent = (level: number, rect: Rect) => {
    const list = indentRects.get(level);
    if (list) list.push(rect);
    else indentRects.set(level, [rect]);
  };
  const setText = (row: number, col: number, text: string) => {
    values[row][col] = text;
    numberFormats[row][col] = TEXT_FORMAT;
  };

  const fieldCells = new Map<string, FieldCell>();
  const setFilterCell = (row: number, col: number, cell: FieldCell, arrow: boolean) => {
    fieldCells.set(memberKey(row, col), cell);
    if (!arrow) return;
    const filtered =
      cell.area === "rows"
        ? Boolean(fields.rowsFiltered && fields.rowsFiltered[cell.index])
        : cell.area === "columns"
          ? Boolean(fields.columnsFiltered && fields.columnsFiltered[cell.index])
          : false;
    numberFormats[row][col] = filtered ? FILTERED_FORMAT : FILTER_FORMAT;
  };

  // Page filters. The helper answers one row per filter, in order; if it ever did not,
  // there would be no telling which row is which field, so none is marked.
  const filtersKnown = F === fields.filters;
  for (let f = 0; f < F; f++) {
    setText(f, 0, grid.filterRows[f].caption);
    setText(f, 1, grid.filterRows[f].value);
    if (filtersKnown) {
      setFilterCell(f, 0, { area: "filters", index: f }, false);
      setFilterCell(f, 1, { area: "filters", index: f }, true);
    }
  }
  if (F > 0) {
    const captions: Rect = { row: 0, col: 0, rows: F, cols: 1 };
    bold.push(captions);
    fill.push(captions);
  }

  // Column header rows: R corner cells, then C header cells.
  for (let h = 0; h < H; h++) {
    const row = headerTop + h;
    const cells = grid.colHeaderRows[h] ?? [];
    const last = h === H - 1;
    for (let r = 0; r < R; r++) {
      setText(row, r, last ? grid.rowFieldCaptions[r] ?? "" : "");
      if (last && r < fields.rows) setFilterCell(row, r, { area: "rows", index: r }, true);
    }
    // A column field's own caption, in the corner cell next to its header cells. The
    // last header row is the measures' and its corner belongs to the row fields.
    if (!last && h < fields.columns.length && R > 0) {
      setText(row, R - 1, fields.columns[h]);
      setFilterCell(row, R - 1, { area: "columns", index: h }, true);
    }
    for (let c = 0; c < C; c++) {
      const cell = cells[c];
      setText(row, R + c, cell ? cell.text : "");
      if (cell && cell.member) {
        members.set(memberKey(row, R + c), { member: cell.member, axis: "columns" });
      }
    }
    runs(
      C,
      (c) => indentOf(cells[c]) || null,
      (level, start, length) =>
        addIndent(level, { row, col: R + start, rows: 1, cols: length }),
    );
  }
  if (H > 0) {
    const header: Rect = { row: headerTop, col: 0, rows: H, cols };
    bold.push(header);
    fill.push(header);
  }

  // Data rows: R row-header cells, then C values.
  const palette = grid.formatPalette;
  const valueFormats = fields.valueFormats ?? [];
  const override = (c: number): string | null => {
    const index = valueFieldIndex(grid.colValueIndex, c, valueFormats.length);
    return index === null ? null : valueFormats[index] ?? null;
  };
  for (let n = 0; n < N; n++) {
    const row = dataTop + n;
    const headers = grid.rowHeaders[n] ?? [];
    for (let r = 0; r < R; r++) {
      const cell = headers[r];
      setText(row, r, cell ? cell.text : "");
      if (cell && cell.member) {
        members.set(memberKey(row, r), { member: cell.member, axis: "rows" });
      }
    }
    const line = grid.values[n] ?? [];
    const formats = grid.cellFormats[n] ?? [];
    for (let c = 0; c < C; c++) {
      const value = toSheetValue(line[c]);
      values[row][R + c] = value;
      numberFormats[row][R + c] =
        typeof value === "string" && value !== ""
          ? TEXT_FORMAT
          : override(c) ?? palette[formats[c]] ?? GENERAL_FORMAT;
    }
  }
  for (let r = 0; r < R; r++) {
    runs(
      N,
      (n) => indentOf((grid.rowHeaders[n] ?? [])[r]) || null,
      (level, start, length) =>
        addIndent(level, { row: dataTop + start, col: r, rows: length, cols: 1 }),
    );
  }
  // Total rows are bold across the whole width.
  runs(
    N,
    (n) => (grid.rowKinds[n] && grid.rowKinds[n] !== "data" ? true : null),
    (_key, start, length) => bold.push({ row: dataTop + start, col: 0, rows: length, cols }),
  );

  const indents: IndentGroup[] = Array.from(indentRects.entries())
    .sort((a, b) => a[0] - b[0])
    .map(([level, rects]) => ({ level, rects }));

  return {
    rows,
    cols,
    values,
    numberFormats,
    bold,
    fill,
    muted: [],
    indents,
    fieldCells,
    autofit: true,
    members,
    shape: {
      filterRows: F,
      headerRows: H,
      headerCols: R,
      dataRows: N,
      dataCols: C,
      headerTop,
      dataTop,
    },
  };
}

/**
 * What a report with no fields puts on the sheet: one quiet cell at its place, so the
 * report is visible, selectable and has somewhere to grow from. The text is left to
 * spill over the cells to its right rather than widen its column.
 */
export function placeholderBlock(): Block {
  const cell: Rect = { row: 0, col: 0, rows: 1, cols: 1 };
  return {
    rows: 1,
    cols: 1,
    values: [[PLACEHOLDER_TEXT]],
    numberFormats: [[TEXT_FORMAT]],
    bold: [],
    fill: [],
    muted: [cell],
    indents: [],
    fieldCells: new Map(),
    autofit: false,
    members: new Map(),
    shape: {
      filterRows: 0,
      headerRows: 0,
      headerCols: 0,
      dataRows: 0,
      dataCols: 0,
      headerTop: 0,
      dataTop: 0,
    },
  };
}

/**
 * The column field whose header row a cell is in (its member and total cells, right of
 * the corner), or null. These cells are not filter cells (a click on one does not
 * open anything), but the pane offers "Filter <field>…" for the selected one.
 */
export function columnHeaderAt(
  shape: BlockShape,
  row: number,
  col: number,
  columnFields: number,
): number | null {
  const h = row - shape.headerTop;
  if (h < 0 || h >= shape.headerRows - 1 || h >= columnFields) return null;
  if (col < shape.headerCols || col >= shape.headerCols + shape.dataCols) return null;
  return h;
}

/**
 * Where the filter cells of a report should be, worked out from its layout alone, for
 * when there is no result in memory (the workbook was just reopened). It assumes what
 * the helper does: one filter row per filter, one header row per column field with the
 * measures' row last, one corner column per row field (one in any case). A caller that
 * has the cell's text should check it against the field's caption before trusting a
 * row or column answer.
 */
export function guessFieldCell(
  layout: Pick<Layout, "filters" | "rows" | "columns">,
  row: number,
  col: number,
): FieldCell | null {
  if (row < 0 || col < 0) return null;
  const F = layout.filters.length;
  if (row < F) return col <= 1 ? { area: "filters", index: row } : null;
  const headerTop = F > 0 ? F + 1 : 0;
  const h = row - headerTop;
  const H = layout.columns.length + 1;
  const R = Math.max(1, layout.rows.length);
  if (h < 0 || h >= H) return null;
  if (h === H - 1) return col < layout.rows.length ? { area: "rows", index: col } : null;
  return col === R - 1 ? { area: "columns", index: h } : null;
}

/** The pieces a host writes one request at a time, each at most `maxCells` cells. */
export function chunkBlock(
  block: Pick<Block, "rows" | "cols">,
  maxCells: number = MAX_CHUNK_CELLS,
): Rect[] {
  return chunkRect({ row: 0, col: 0, rows: block.rows, cols: block.cols }, maxCells);
}

/** Copy one rectangle out of a rows x cols matrix. */
export function sliceMatrix<T>(matrix: T[][], rect: Rect): T[][] {
  const out: T[][] = [];
  for (let r = rect.row; r < rect.row + rect.rows; r++) {
    out.push(matrix[r].slice(rect.col, rect.col + rect.cols));
  }
  return out;
}

/**
 * Text as Excel must receive it so that it stays text. A leading apostrophe is Excel's
 * own "this is text" prefix: it is not shown and not part of the value. Without it a
 * caption such as "=Sales" would be read as a formula.
 */
export function excelText(text: string): string {
  if (text === "") return text;
  const first = text.charAt(0);
  return first === "=" || first === "+" || first === "-" || first === "@" || first === "'"
    ? "'" + text
    : text;
}

/**
 * The values matrix as Excel should receive it: text cells guarded with `excelText`,
 * filter cells forced to text, and, when the host cannot indent with formatting, indent levels turned into leading
 * spaces. Returns a new matrix; the block is not changed.
 */
export function excelValues(block: Block, indentWithSpaces: boolean): SheetValue[][] {
  const out = block.values.map((line, r) =>
    line.map((value, c) => {
      if (typeof value !== "string" || value === "") return value;
      const format = block.numberFormats[r][c];
      // A filter cell has a custom format, under which Excel would read "2025" as a
      // number and drop the arrow. The apostrophe makes it text whatever it looks like.
      if (isFilterFormat(format)) return "'" + value;
      return format === TEXT_FORMAT ? excelText(value) : value;
    }),
  );
  if (indentWithSpaces) {
    for (const group of block.indents) {
      const pad = " ".repeat(group.level * INDENT_SPACES);
      for (const rect of group.rects) {
        for (let r = rect.row; r < rect.row + rect.rows; r++) {
          for (let c = rect.col; c < rect.col + rect.cols; c++) {
            const value = block.values[r][c];
            if (typeof value === "string" && value !== "") out[r][c] = pad + value;
          }
        }
      }
    }
  }
  return out;
}

export interface GridHeaders {
  rowHeaders: HeaderCell[][];
  colHeaderRows: HeaderCell[][];
}

/**
 * Unique names of the expanded members shown underneath `member` in the last grid, so
 * that collapsing a member also forgets the drill state of what was below it. Children
 * always follow their parent on an axis, with a deeper level.
 */
export function expandedDescendants(headers: GridHeaders, axis: Axis, member: MemberRef): string[] {
  const sequences: (HeaderCell | undefined)[][] = [];
  if (axis === "columns") {
    for (const row of headers.colHeaderRows) sequences.push(row);
  } else {
    const width = headers.rowHeaders.reduce((w, row) => Math.max(w, row.length), 0);
    for (let r = 0; r < width; r++) sequences.push(headers.rowHeaders.map((row) => row[r]));
  }

  const found = new Set<string>();
  for (const sequence of sequences) {
    let inside = false;
    for (const cell of sequence) {
      const ref = cell ? cell.member : null;
      if (!ref || ref.hierarchy !== member.hierarchy) {
        inside = false;
      } else if (ref.uniqueName === member.uniqueName) {
        inside = true;
      } else if (inside && ref.level > member.level) {
        if (ref.expanded) found.add(ref.uniqueName);
      } else {
        inside = false;
      }
    }
  }
  return Array.from(found);
}

/**
 * The members behind a value cell, for Show Details: the member of each row header
 * cell of its line and of each column header cell above it. Totals have no member and
 * add nothing; a blanked repeat still stands for its member. `row` and `col` count
 * data lines and value columns from 0.
 */
export function detailsMembers(headers: GridHeaders, row: number, col: number): string[] {
  const out: string[] = [];
  const add = (cell: HeaderCell | undefined) => {
    if (cell && cell.member && out.indexOf(cell.member.uniqueName) < 0) out.push(cell.member.uniqueName);
  };
  for (const cell of headers.rowHeaders[row] ?? []) add(cell);
  for (const line of headers.colHeaderRows) add(line[col]);
  return out;
}

/**
 * The cells of a Show Details sheet: a header row of column names, then the rows. Text
 * is written as text (so "01" or "Jan 2025" stay as they are), numbers as General.
 * Rows shorter than the header are padded; longer ones are cut to the header's width.
 */
export function detailsCells(
  columns: string[],
  rows: CellValue[][],
): { values: SheetValue[][]; numberFormats: string[][]; cols: number } {
  const cols = Math.max(1, columns.length);
  const values: SheetValue[][] = [];
  const numberFormats: string[][] = [];
  const header: SheetValue[] = [];
  const headerFormats: string[] = [];
  for (let c = 0; c < cols; c++) {
    header.push(columns[c] ?? "");
    headerFormats.push(TEXT_FORMAT);
  }
  values.push(header);
  numberFormats.push(headerFormats);
  for (const row of rows) {
    const line: SheetValue[] = [];
    const formats: string[] = [];
    for (let c = 0; c < cols; c++) {
      const value = toSheetValue(row[c]);
      line.push(value);
      formats.push(typeof value === "string" && value !== "" ? TEXT_FORMAT : GENERAL_FORMAT);
    }
    values.push(line);
    numberFormats.push(formats);
  }
  return { values, numberFormats, cols };
}
