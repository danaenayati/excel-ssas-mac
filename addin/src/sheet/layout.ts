// Grid -> sheet block. Pure: turns a query result into everything a host needs to
// write it (cell values, number formats, formatting rectangles) plus the lookup from
// a cell position to the cube member behind it.
//
// Block layout, top to bottom (see `Grid` in types.ts):
//   F rows   caption | value            page filters
//   1 row    blank                      only when F > 0
//   1 row    "Column Labels"            compact form with column fields only
//   H rows   R corner cells | C header cells   (last row's corner = row field captions)
//   N rows   R row-header cells | C values
// All positions in a Block are relative to its top-left cell.
//
// Excel's three report layouts differ in the row headers only:
//   tabular  one column per row field; an outer item shares its first inner item's line
//            (the grid's own lines, one per tuple)
//   outline  one column per row field; each outer item on a line of its own, its inner
//            items from the next line on, one column further right
//   compact  every row field in one column ("Row Labels"); each outer item on a line of
//            its own, inner items below it, indented one step per field
// In outline and compact form a sheet line therefore does not always show a grid line:
// `Block.lines` says which grid line each one shows, if any.

import { isFiltered, reportForm, subtotalsChoice, type ReportForm } from "../state/layout";
import type { CellValue, Grid, HeaderCell, Layout, LayoutOptions, LineKind, MemberRef } from "../types";
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
/** Compact form's button for all row fields, in Windows' words. */
export const ROW_LABELS = "Row Labels";
/** Compact form's button for all column fields, above the column items. */
export const COLUMN_LABELS = "Column Labels";

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
  /** Row of the first column-header row of the grid. */
  headerTop: number;
  /** Row of the first data row. */
  dataTop: number;
  /** Compact form with column fields: the "Column Labels" row, just above `headerTop`. */
  columnLabelsRow?: number;
}

/** A data line of the block and the grid line behind it. */
export interface DataLine {
  /**
   * The grid line whose values the sheet line shows. Null for an outer item's own line
   * in outline and compact form when it shows no values (subtotals off, or at the bottom).
   */
  grid: number | null;
  /**
   * Unique names of the row members its values belong to, outermost first, as Show
   * Details needs them. A subtotal's own item is included ("Tehran Total" is Tehran's).
   */
  members: string[];
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
  /**
   * "Row Labels" or "Column Labels" of compact form: the cell stands for every field of
   * its area (Windows asks "Select field:" first) and `index` is the first of them.
   */
  labels?: boolean;
}

/** The text of a "Row Labels" or "Column Labels" cell. */
export function labelsCaption(area: FieldCell["area"]): string {
  return area === "columns" ? COLUMN_LABELS : ROW_LABELS;
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
  /** "Repeat All Item Labels": write a repeated outer row label on every line. Not in compact form. */
  repeatLabels?: boolean;
  /** For each value field in Values order: a KPI's Status or Trend, shown as icons as in Excel. */
  valueIcons?: (KpiIcon | null)[];
  /** "For empty cells show": the text for value cells that have no value. */
  emptyCellText?: string;
  /** Excel's Report Layout. Absent: tabular. */
  form?: ReportForm;
  /**
   * "top" when the grid was asked for with each subtotal before its group (compact and
   * outline form only). Absent: subtotals, if any, follow their group.
   */
  subtotalsAt?: "top";
}

export function blockFields(
  layout: Pick<Layout, "filters" | "rows" | "columns"> & {
    values?: Layout["values"];
    options?: Layout["options"];
  },
): BlockFields {
  const fields: BlockFields = {
    filters: layout.filters.length,
    rows: layout.rows.length,
    columns: layout.columns.map((field) => field.caption),
    rowsFiltered: layout.rows.map(isFiltered),
    columnsFiltered: layout.columns.map(isFiltered),
    valueFormats: (layout.values ?? []).map((value) => value.numberFormat || null),
    valueIcons: (layout.values ?? []).map((value) =>
      value.kpi && (value.kpi.part === "status" || value.kpi.part === "trend") ? value.kpi.part : null,
    ),
    repeatLabels: Boolean(layout.options && layout.options.repeatLabels),
    emptyCellText: (layout.options && layout.options.emptyCellText) || "",
  };
  if (layout.options) {
    const form = reportForm(layout.options);
    if (form !== "tabular") fields.form = form;
    if (subtotalsChoice(layout.options) === "top") fields.subtotalsAt = "top";
  }
  return fields;
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

/** How a KPI column is drawn: Status as traffic lights, Trend as arrows (values -1, 0, 1). */
export type KpiIcon = "status" | "trend";

export interface Block {
  rows: number;
  cols: number;
  /** Value columns of KPI parts that Excel shows as icons rather than numbers. */
  icons?: { kind: KpiIcon; rect: Rect }[];
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
  /** One per data row: the grid line it shows and the row members behind it. */
  lines: DataLine[];
  /** One per value column: the column members behind it, a subtotal's own item included. */
  columnMembers: string[][];
  /**
   * The grid as written: the one given, or a copy with its columns in sheet order (Excel
   * keeps column subtotals after their items when the row subtotals are at the top).
   * Value column c of the block is column c of this grid.
   */
  grid: Grid;
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
 * In compact form the row fields share one column: the corner reads "Row Labels" and a
 * "Column Labels" row above the column items replaces the column fields' captions; each
 * of the two stands for every field of its area.
 *
 * Without `fields` the block is the bare grid, in tabular form: no filter cells, no
 * column captions.
 */
export function buildBlock(input: Grid, fields: BlockFields = { filters: 0, rows: 0, columns: [] }): Block {
  const form: ReportForm = fields.form ?? "tabular";
  const atTop = form !== "tabular" && fields.subtotalsAt === "top";
  const K = fields.columns.length;
  const grid = atTop ? columnsInSheetOrder(input, K) : input;

  const F = grid.filterRows.length;
  const H = grid.colHeaderRows.length;
  const R = grid.rowFieldCaptions.length;
  const C = grid.colKinds.length;

  // The row fields are laid out by form only when the grid has one column per row field.
  const rowForm: ReportForm = fields.rows > 0 && R === fields.rows ? form : "tabular";
  const compactRows = rowForm === "compact";
  const sheet = sheetRows(grid, rowForm, atTop && rowForm !== "tabular");
  const N = sheet.headers.length;
  /** Row-header columns on the sheet. */
  const RS = compactRows ? 1 : R;
  const labelsRow = form === "compact" && K > 0;

  const top = F > 0 ? F + 1 : 0;
  const headerTop = top + (labelsRow ? 1 : 0);
  const dataTop = headerTop + H;
  const rows = dataTop + N;
  // A filter row needs two cells even if the grid itself is one column wide; the
  // "Column Labels" cell needs one right of the corner.
  const cols = Math.max(RS + C, labelsRow ? RS + 1 : 1, F > 0 ? 2 : 1);

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
    const list =
      cell.area === "rows" ? fields.rowsFiltered : cell.area === "columns" ? fields.columnsFiltered : undefined;
    // "Row Labels" and "Column Labels" stand for every field of their area.
    const filtered = cell.labels ? Boolean(list && list.some(Boolean)) : Boolean(list && list[cell.index]);
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

  // Compact form: "Column Labels" right above the first column item.
  if (labelsRow) {
    for (let c = 0; c < cols; c++) setText(top, c, "");
    setText(top, RS, COLUMN_LABELS);
    setFilterCell(top, RS, { area: "columns", index: 0, labels: true }, true);
  }

  // Column header rows: RS corner cells, then C header cells.
  for (let h = 0; h < H; h++) {
    const row = headerTop + h;
    const cells = grid.colHeaderRows[h] ?? [];
    const last = h === H - 1;
    if (compactRows) {
      setText(row, 0, last ? ROW_LABELS : "");
      if (last) setFilterCell(row, 0, { area: "rows", index: 0, labels: true }, true);
    } else {
      for (let r = 0; r < R; r++) {
        setText(row, r, last ? grid.rowFieldCaptions[r] ?? "" : "");
        // A set spanning several fields takes several columns; rowFieldOf says whose.
        const field = grid.rowFieldOf ? grid.rowFieldOf[r] : r;
        if (last && field !== undefined && field < fields.rows) {
          setFilterCell(row, r, { area: "rows", index: field }, true);
        }
      }
    }
    // A column field's own caption, in the corner cell next to its header cells. The
    // last header row is the measures' and its corner belongs to the row fields. Compact
    // form has "Column Labels" instead.
    const columnField = grid.colFieldOf ? grid.colFieldOf[h] : h;
    if (!last && columnField !== undefined && columnField < K && RS > 0 && form !== "compact") {
      setText(row, RS - 1, fields.columns[columnField]);
      setFilterCell(row, RS - 1, { area: "columns", index: columnField }, true);
    }
    for (let c = 0; c < C; c++) {
      const cell = cells[c];
      setText(row, RS + c, cell ? cell.text : "");
      if (cell && cell.member) {
        members.set(memberKey(row, RS + c), { member: cell.member, axis: "columns" });
      }
    }
    runs(
      C,
      (c) => indentOf(cells[c]) || null,
      (level, start, length) =>
        addIndent(level, { row, col: RS + start, rows: 1, cols: length }),
    );
  }
  if (H > 0) {
    const header: Rect = { row: top, col: 0, rows: dataTop - top, cols };
    bold.push(header);
    fill.push(header);
  }

  // Data rows: RS row-header cells, then C values.
  const palette = grid.formatPalette;
  const emptyText = fields.emptyCellText ?? "";
  const valueFormats = fields.valueFormats ?? [];
  const repeat = Boolean(fields.repeatLabels) && !compactRows;
  const override = (c: number): string | null => {
    const index = valueFieldIndex(grid.colValueIndex, c, valueFormats.length);
    return index === null ? null : valueFormats[index] ?? null;
  };
  for (let n = 0; n < N; n++) {
    const row = dataTop + n;
    const headers = sheet.headers[n] ?? [];
    for (let r = 0; r < RS; r++) {
      const cell = headers[r];
      // A repeat of the outer label is sent blank, with its member: shown again if asked.
      const text = !cell ? "" : cell.kind === "blank" && cell.member && repeat ? cell.member.caption : cell.text;
      setText(row, r, text);
      if (cell && cell.member) {
        members.set(memberKey(row, r), { member: cell.member, axis: "rows" });
      }
    }
    // An outer item's own line without values has no value cells, not empty ones.
    const g = sheet.lines[n].grid;
    if (g === null) continue;
    const line = grid.values[g] ?? [];
    const formats = grid.cellFormats[g] ?? [];
    for (let c = 0; c < C; c++) {
      const raw = line[c];
      const empty = raw === null || raw === undefined;
      const value = empty && emptyText !== "" ? emptyText : toSheetValue(raw);
      values[row][RS + c] = value;
      numberFormats[row][RS + c] =
        typeof value === "string" && value !== ""
          ? TEXT_FORMAT
          : override(c) ?? palette[formats[c]] ?? GENERAL_FORMAT;
    }
  }
  const valueIcons = fields.valueIcons ?? [];
  const icons: { kind: KpiIcon; rect: Rect }[] = [];
  for (let c = 0; N > 0 && c < C; c++) {
    const index = valueFieldIndex(grid.colValueIndex, c, valueIcons.length);
    const kind = index === null ? null : valueIcons[index];
    if (kind) icons.push({ kind, rect: { row: dataTop, col: RS + c, rows: N, cols: 1 } });
  }
  for (let r = 0; r < RS; r++) {
    runs(
      N,
      (n) => indentOf((sheet.headers[n] ?? [])[r]) || null,
      (level, start, length) =>
        addIndent(level, { row: dataTop + start, col: r, rows: length, cols: 1 }),
    );
  }
  // Total rows, and in outline and compact form the outer items' own lines, are bold
  // across the whole width.
  runs(
    N,
    (n) => (sheet.kinds[n] && sheet.kinds[n] !== "data" ? true : null),
    (_key, start, length) => bold.push({ row: dataTop + start, col: 0, rows: length, cols }),
  );

  const indents: IndentGroup[] = Array.from(indentRects.entries())
    .sort((a, b) => a[0] - b[0])
    .map(([level, rects]) => ({ level, rects }));

  const shape: BlockShape = {
    filterRows: F,
    headerRows: H,
    headerCols: RS,
    dataRows: N,
    dataCols: C,
    headerTop,
    dataTop,
  };
  if (labelsRow) shape.columnLabelsRow = top;

  return {
    rows,
    cols,
    values,
    numberFormats,
    bold,
    fill,
    muted: [],
    indents,
    icons,
    fieldCells,
    autofit: true,
    members,
    shape,
    lines: sheet.lines,
    // Moved into sheet order, the columns follow their groups; otherwise they are as asked.
    columnMembers: columnMembers(grid, atTop && grid === input),
    grid,
  };
}

// ---------------------------------------------------------------------------
// Axis items: what each line or column stands for.
// ---------------------------------------------------------------------------

/** An item of a field on an axis line: its member, or the text shown when it has none. */
interface PathItem {
  member: MemberRef | null;
  /** Indent of its own cell: its level in the hierarchy. */
  indent: number;
  text: string;
}

/**
 * One position on an axis: a grid line, or the run of columns that repeat one column
 * tuple for each measure.
 */
interface AxisUnit {
  start: number;
  size: number;
  kind: LineKind;
  /** The header cells of its first position, one per field. */
  cells: (HeaderCell | undefined)[];
  /**
   * Its items, outermost first: every field's for a data line, none for the grand total,
   * and for a subtotal the items outside it and then its own, the one its "X Total" cell
   * names (with no member when that cannot be told).
   */
  path: PathItem[];
}

function itemOf(cell: HeaderCell | undefined): PathItem {
  return {
    member: cell && cell.member ? cell.member : null,
    indent: indentOf(cell),
    text: cell ? (cell.member ? cell.member.caption : cell.text) : "",
  };
}

function sameItem(a: PathItem | undefined, b: PathItem | undefined): boolean {
  return Boolean(a && b && a.member && b.member && a.member.uniqueName === b.member.uniqueName);
}

function memberNames(path: PathItem[]): string[] {
  const out: string[] = [];
  for (const item of path) {
    if (item.member && out.indexOf(item.member.uniqueName) < 0) out.push(item.member.uniqueName);
  }
  return out;
}

/** The same column tuple again, for the next measure: every label blanked, nothing new. */
function repeatsTuple(before: (HeaderCell | undefined)[], cells: (HeaderCell | undefined)[]): boolean {
  for (let i = 0; i < cells.length; i++) {
    const a = before[i];
    const b = cells[i];
    if (b && b.member) {
      if (b.kind !== "blank" || !a || !a.member || a.member.uniqueName !== b.member.uniqueName) return false;
    } else if ((a && a.member) || (b ? b.text : "") !== "" || (a ? a.kind : null) !== (b ? b.kind : null)) {
      return false;
    }
  }
  return true;
}

/**
 * The units of an axis with their items. A subtotal line carries the members outside
 * it but not its own item: that is read from the first line of its group, which is the
 * next unit when subtotals head their group (`atTop`) and the previous one when they
 * follow it.
 */
function axisUnits(
  count: number,
  fieldCount: number,
  cellsAt: (pos: number) => (HeaderCell | undefined)[],
  kindAt: (pos: number) => LineKind,
  groupRepeats: boolean,
  atTop: boolean,
): AxisUnit[] {
  const units: AxisUnit[] = [];
  for (let pos = 0; pos < count; pos++) {
    const cells = cellsAt(pos);
    const kind = kindAt(pos);
    const last = units[units.length - 1];
    if (groupRepeats && last && last.kind === kind && repeatsTuple(cellsAt(pos - 1), cells)) {
      last.size++;
      continue;
    }
    units.push({ start: pos, size: 1, kind, cells, path: [] });
  }
  const depthOf = (unit: AxisUnit) =>
    Math.max(0, unit.cells.findIndex((cell) => cell !== undefined && cell.kind === "subtotal"));
  units.forEach((unit, u) => {
    if (unit.kind === "grandTotal") return;
    if (unit.kind !== "subtotal") {
      for (let i = 0; i < fieldCount; i++) unit.path.push(itemOf(unit.cells[i]));
      return;
    }
    const depth = depthOf(unit);
    for (let i = 0; i < depth; i++) unit.path.push(itemOf(unit.cells[i]));
    const label = unit.cells[depth];
    let own: PathItem = { member: null, indent: 0, text: label ? label.text : "" };
    const first = units[atTop ? u + 1 : u - 1];
    const inGroup =
      first !== undefined &&
      (first.kind === "data" || (first.kind === "subtotal" && depthOf(first) > depth)) &&
      unit.path.every((item, i) => sameItem(item, itemOf(first.cells[i])));
    if (inGroup) {
      const cell = first.cells[depth];
      if (cell && cell.member) own = itemOf(cell);
    }
    unit.path.push(own);
  });
  return units;
}

function columnUnits(grid: Grid, atTop: boolean): AxisUnit[] {
  // Every header row but the last (the measures') belongs to a column field.
  const fieldRows = Math.max(0, grid.colHeaderRows.length - 1);
  return axisUnits(
    grid.colKinds.length,
    fieldRows,
    (c) => Array.from({ length: fieldRows }, (_, r) => (grid.colHeaderRows[r] ?? [])[c]),
    (c) => grid.colKinds[c] ?? "data",
    true,
    atTop,
  );
}

/** For each value column, the column members behind it, a subtotal's own item included. */
function columnMembers(grid: Grid, atTop: boolean): string[][] {
  const out: string[][] = [];
  for (const unit of columnUnits(grid, atTop)) {
    const names = memberNames(unit.path);
    for (let i = 0; i < unit.size; i++) out.push(names.slice());
  }
  return out;
}

/**
 * The grid with each column subtotal moved after the items it adds up. The helper puts
 * subtotals at the top of their group on both axes; Excel does that on rows only and
 * keeps column subtotals on the right. Labels are blanked again for the new order, as
 * the helper does: an item is written in the first column of its run only.
 */
function columnsInSheetOrder(grid: Grid, K: number): Grid {
  if (K < 2 || grid.colHeaderRows.length !== K + 1) return grid;
  const units = columnUnits(grid, true);
  const inGroup = (unit: AxisUnit, subtotal: AxisUnit) =>
    subtotal.path.every((item, i) => sameItem(item, unit.path[i]));
  const order: AxisUnit[] = [];
  const waiting: AxisUnit[] = [];
  for (const unit of units) {
    while (waiting.length > 0 && !inGroup(unit, waiting[waiting.length - 1])) {
      order.push(waiting.pop() as AxisUnit);
    }
    if (unit.kind === "subtotal") waiting.push(unit);
    else order.push(unit);
  }
  while (waiting.length > 0) order.push(waiting.pop() as AxisUnit);
  if (order.every((unit, i) => unit === units[i])) return grid;

  const columns: number[] = [];
  const unitOf: AxisUnit[] = [];
  for (const unit of order) {
    for (let i = 0; i < unit.size; i++) {
      columns.push(unit.start + i);
      unitOf.push(unit);
    }
  }
  // What a column shows for field r: an item, or a total (past a subtotal's own item).
  const key = (unit: AxisUnit, r: number): string | null => {
    if (r < unit.path.length) {
      const member = unit.path[r].member;
      return member ? member.uniqueName : null;
    }
    return unit.kind === "data" ? null : "\u0000total";
  };
  const sameUpTo = (a: AxisUnit, b: AxisUnit, r: number) => {
    for (let i = 0; i <= r; i++) {
      const x = key(a, i);
      if (x === null || x !== key(b, i)) return false;
    }
    return true;
  };
  const pick = <T>(line: T[] | undefined): T[] => columns.map((c) => (line ?? [])[c]);
  const colHeaderRows = grid.colHeaderRows.map((line, r) => {
    const moved = pick(line);
    if (r >= K) return moved;
    return moved.map((cell, p): HeaderCell => {
      if (!cell || !cell.member || (cell.kind !== "member" && cell.kind !== "blank")) return cell;
      return p > 0 && sameUpTo(unitOf[p - 1], unitOf[p], r)
        ? { ...cell, text: "", kind: "blank" }
        : { ...cell, text: cell.member.caption, kind: "member" };
    });
  });
  return {
    ...grid,
    colHeaderRows,
    colKinds: pick(grid.colKinds),
    colValueIndex: grid.colValueIndex ? pick(grid.colValueIndex) : grid.colValueIndex,
    values: grid.values.map((line) => pick(line)),
    cellFormats: grid.cellFormats.map((line) => pick(line)),
  };
}

// ---------------------------------------------------------------------------
// Row lines by report form.
// ---------------------------------------------------------------------------

interface SheetRows {
  /** Header cells of each data line: one per row field, or one in compact form. */
  headers: HeaderCell[][];
  /** Each line's kind, for the bold of total lines. */
  kinds: LineKind[];
  lines: DataLine[];
}

/**
 * The data lines of the sheet. Tabular form writes the grid's lines as they are. Outline
 * and compact form give each outer item a line of its own before its inner items; with
 * subtotals at the top of the group (`atTop`) that line shows the subtotal's values and
 * the subtotal has no line of its own. Otherwise the subtotal ("Tehran Total") follows
 * its group, as in tabular form.
 */
function sheetRows(grid: Grid, form: ReportForm, atTop: boolean): SheetRows {
  const R = grid.rowFieldCaptions.length;
  const units = axisUnits(
    grid.rowHeaders.length,
    R,
    (g) => grid.rowHeaders[g] ?? [],
    (g) => grid.rowKinds[g] ?? "data",
    false,
    atTop,
  );
  if (form === "tabular") {
    return {
      headers: grid.rowHeaders,
      kinds: grid.rowKinds,
      lines: units.map((unit) => ({ grid: unit.start, members: memberNames(unit.path) })),
    };
  }

  const compact = form === "compact";
  const out: SheetRows = { headers: [], kinds: [], lines: [] };
  /** The outer items whose own line is written, by depth, with that line's indent. */
  const open: { item: PathItem; indent: number }[] = [];
  const baseIndent = (depth: number) => (depth === 0 ? 0 : open[depth - 1] ? open[depth - 1].indent + 1 : depth);
  const blankCell = (item?: PathItem): HeaderCell => ({
    text: "",
    indent: item ? item.indent : 0,
    kind: "blank",
    member: item ? item.member : null,
  });
  const push = (cells: HeaderCell[], kind: LineKind, line: DataLine) => {
    out.headers.push(cells);
    out.kinds.push(kind);
    out.lines.push(line);
  };

  /** The line of the item at `depth` of `path`, showing grid line `values`. Returns its indent. */
  const itemLine = (path: PathItem[], depth: number, values: number | null, kind: LineKind): number => {
    const item = path[depth];
    const indent = baseIndent(depth) + item.indent;
    const cell: HeaderCell = { text: item.text, indent: compact ? indent : item.indent, kind: "member", member: item.member };
    const cells = compact
      ? [cell]
      : Array.from({ length: R }, (_, i) => (i < depth ? blankCell(path[i]) : i === depth ? cell : blankCell()));
    push(cells, kind, { grid: values, members: memberNames(path.slice(0, depth + 1)) });
    return indent;
  };

  /**
   * Write the own lines of the outer items `path[0..upto)` that are not written yet, and
   * from depth `fresh` on in any case. The innermost of them shows grid line `values`.
   */
  const openItems = (path: PathItem[], upto: number, fresh: number, values: number | null) => {
    let j = 0;
    while (j < upto && j < fresh && j < open.length && sameItem(open[j].item, path[j])) j++;
    open.length = j;
    for (; j < upto; j++) {
      const indent = itemLine(path, j, j === upto - 1 ? values : null, "subtotal");
      open.push({ item: path[j], indent });
    }
  };

  /** "Tehran Total" after its group, a subtotal whose item cannot be told, or Grand Total. */
  const totalLine = (unit: AxisUnit, depth: number) => {
    const label = (grid.rowHeaders[unit.start] ?? [])[depth];
    const text = label ? label.text : "";
    const kind: LineKind = unit.kind === "grandTotal" ? "grandTotal" : "subtotal";
    const line: DataLine = { grid: unit.start, members: memberNames(unit.path) };
    if (compact) {
      const own = open[depth] && sameItem(open[depth].item, unit.path[depth]) ? open[depth].indent : baseIndent(depth);
      push([{ text, indent: kind === "grandTotal" ? 0 : own, kind, member: null }], kind, line);
      return;
    }
    const cells = Array.from({ length: R }, (_, i): HeaderCell =>
      i < depth ? blankCell(unit.path[i]) : { text: i === depth ? text : "", indent: 0, kind, member: null },
    );
    push(cells, kind, line);
  };

  for (const unit of units) {
    if (unit.kind === "grandTotal") {
      open.length = 0;
      totalLine(unit, 0);
    } else if (unit.kind === "subtotal") {
      const depth = unit.path.length - 1;
      if (atTop && unit.path[depth].member) {
        // Its values go on its item's own line, which starts the group.
        openItems(unit.path, depth + 1, depth, unit.start);
      } else {
        totalLine(unit, depth);
        open.length = Math.min(open.length, depth);
      }
    } else {
      openItems(unit.path, R - 1, R, null);
      itemLine(unit.path, R - 1, unit.start, "data");
    }
  }
  return out;
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
    lines: [],
    columnMembers: [],
    grid: {
      filterRows: [],
      rowFieldCaptions: [""],
      colHeaderRows: [[]],
      rowHeaders: [],
      values: [],
      formatPalette: [GENERAL_FORMAT],
      cellFormats: [],
      rowKinds: [],
      colKinds: [],
      colValueIndex: [],
      stats: { rows: 0, columns: 0, elapsedMs: 0, mdx: "" },
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
 * measures' row last, one corner column per row field (one in any case, and one only
 * in compact form, which has its "Column Labels" row above the header rows). A caller
 * that has the cell's text should check it against the field's caption (or
 * `labelsCaption`) before trusting a row or column answer.
 */
export function guessFieldCell(
  layout: Pick<Layout, "filters" | "rows" | "columns"> & { options?: LayoutOptions },
  row: number,
  col: number,
): FieldCell | null {
  if (row < 0 || col < 0) return null;
  const F = layout.filters.length;
  if (row < F) return col <= 1 ? { area: "filters", index: row } : null;
  const top = F > 0 ? F + 1 : 0;
  if (layout.options && reportForm(layout.options) === "compact") {
    const labelsRow = layout.columns.length > 0;
    if (labelsRow && row === top) return col === 1 ? { area: "columns", index: 0, labels: true } : null;
    const last = top + (labelsRow ? 1 : 0) + layout.columns.length;
    return row === last && col === 0 && layout.rows.length > 0 ? { area: "rows", index: 0, labels: true } : null;
  }
  const headerTop = top;
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
 * The members behind value cell (`line`, `col`) of a block, counted from its first data
 * line and value column, for Show Details: the row members of its line and the column
 * members of its column, a subtotal's own item included. Null when the line shows no
 * values (an outer item's own line in outline or compact form) or is out of range.
 */
export function valueCellMembers(
  block: Pick<Block, "lines" | "columnMembers">,
  line: number,
  col: number,
): string[] | null {
  const shown = block.lines[line];
  if (!shown || shown.grid === null || col < 0 || col >= block.columnMembers.length) return null;
  const out = shown.members.slice();
  for (const name of block.columnMembers[col]) if (out.indexOf(name) < 0) out.push(name);
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

/**
 * What "Expand Entire Field" works on: the field's members shown in the last result
 * that have children and are not expanded yet (one level deeper each time, as in
 * Excel), and whether any of its shown members has children at all.
 */
export function expandableMembers(
  headers: GridHeaders,
  hierarchy: string,
): { members: string[]; anyChildren: boolean } {
  const members: string[] = [];
  let anyChildren = false;
  const look = (cell: HeaderCell | undefined) => {
    const ref = cell ? cell.member : null;
    if (!ref || ref.hierarchy !== hierarchy || !ref.hasChildren) return;
    anyChildren = true;
    if (!ref.expanded && members.indexOf(ref.uniqueName) < 0) members.push(ref.uniqueName);
  };
  for (const line of headers.rowHeaders) line.forEach(look);
  for (const line of headers.colHeaderRows) line.forEach(look);
  return { members, anyChildren };
}
