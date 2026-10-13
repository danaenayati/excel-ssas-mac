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

import type { CellValue, Grid, HeaderCell, MemberRef } from "../types";
import { chunkRect, runs, type Rect } from "./geometry";

/** Excel's "Text" number format: stops Excel from reading "2024" or "March 2025" as a number or date. */
export const TEXT_FORMAT = "@";
export const GENERAL_FORMAT = "General";
/** Most cells written or formatted in one request. */
export const MAX_CHUNK_CELLS = 20000;
/** Spaces per indent level when the host cannot indent with formatting. */
export const INDENT_SPACES = 3;

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

export interface Block {
  rows: number;
  cols: number;
  /** rows x cols. Empty cells are "". */
  values: SheetValue[][];
  /** rows x cols Excel number formats. */
  numberFormats: string[][];
  bold: Rect[];
  fill: Rect[];
  /** Cells to indent, grouped by level (1 or more). */
  indents: IndentGroup[];
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

export function buildBlock(grid: Grid): Block {
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

  // Page filters.
  for (let f = 0; f < F; f++) {
    setText(f, 0, grid.filterRows[f].caption);
    setText(f, 1, grid.filterRows[f].value);
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
    for (let r = 0; r < R; r++) {
      setText(row, r, h === H - 1 ? grid.rowFieldCaptions[r] ?? "" : "");
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
          : palette[formats[c]] ?? GENERAL_FORMAT;
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
    indents,
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
 * and, when the host cannot indent with formatting, indent levels turned into leading
 * spaces. Returns a new matrix; the block is not changed.
 */
export function excelValues(block: Block, indentWithSpaces: boolean): SheetValue[][] {
  const out = block.values.map((line, r) =>
    line.map((value, c) =>
      typeof value === "string" && block.numberFormats[r][c] === TEXT_FORMAT
        ? excelText(value)
        : value,
    ),
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
