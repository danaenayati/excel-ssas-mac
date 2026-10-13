// Rectangle and A1-address arithmetic. Pure; rows and columns are zero-based.

export interface Rect {
  row: number;
  col: number;
  rows: number;
  cols: number;
}

/** Excel's sheet size. */
export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

/** 0 -> "A", 25 -> "Z", 26 -> "AA", 16383 -> "XFD". */
export function columnLetters(col: number): string {
  let n = Math.floor(col);
  let out = "";
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
}

export function cellA1(row: number, col: number): string {
  return columnLetters(col) + String(row + 1);
}

/** "B3:D10", or "B3" for a single cell. Sheet-relative, no sheet name. */
export function rectA1(rect: Rect): string {
  const first = cellA1(rect.row, rect.col);
  if (rect.rows === 1 && rect.cols === 1) return first;
  return first + ":" + cellA1(rect.row + rect.rows - 1, rect.col + rect.cols - 1);
}

/** "AA" -> 26. Null when the text is not a column name on the sheet. */
function columnIndex(letters: string): number | null {
  if (!/^[A-Za-z]{1,3}$/.test(letters)) return null;
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n >= 1 && n <= MAX_COLS ? n - 1 : null;
}

function parseCell(text: string): { row: number; col: number } | null {
  const match = /^([A-Za-z]{1,3})([0-9]{1,7})$/.exec(text);
  if (!match) return null;
  const col = columnIndex(match[1]);
  const row = Number(match[2]);
  if (col === null || row < 1 || row > MAX_ROWS) return null;
  return { row: row - 1, col };
}

/**
 * The rectangle an address names: "C2", "B2:H30", "$B$2", "Sheet1!B2:C3". Null for
 * anything else (whole rows or columns, several areas, a name), which callers treat
 * as "not a small cell edit".
 */
export function parseA1(address: string): Rect | null {
  const bang = address.lastIndexOf("!");
  const local = (bang >= 0 ? address.slice(bang + 1) : address).replace(/\$/g, "").trim();
  const parts = local.split(":");
  if (parts.length > 2) return null;
  const first = parseCell(parts[0]);
  const last = parts.length === 2 ? parseCell(parts[1]) : first;
  if (!first || !last) return null;
  const row = Math.min(first.row, last.row);
  const col = Math.min(first.col, last.col);
  return {
    row,
    col,
    rows: Math.abs(first.row - last.row) + 1,
    cols: Math.abs(first.col - last.col) + 1,
  };
}

export function offsetRect(rect: Rect, dRow: number, dCol: number): Rect {
  return { row: rect.row + dRow, col: rect.col + dCol, rows: rect.rows, cols: rect.cols };
}

export function containsCell(rect: Rect, row: number, col: number): boolean {
  return (
    row >= rect.row && row < rect.row + rect.rows && col >= rect.col && col < rect.col + rect.cols
  );
}

export function intersect(a: Rect, b: Rect): Rect | null {
  const row = Math.max(a.row, b.row);
  const col = Math.max(a.col, b.col);
  const bottom = Math.min(a.row + a.rows, b.row + b.rows);
  const right = Math.min(a.col + a.cols, b.col + b.cols);
  if (bottom <= row || right <= col) return null;
  return { row, col, rows: bottom - row, cols: right - col };
}

/** The parts of `a` that are not inside `b`: at most four non-overlapping rectangles. */
export function subtract(a: Rect, b: Rect | null): Rect[] {
  if (a.rows <= 0 || a.cols <= 0) return [];
  const hit = b ? intersect(a, b) : null;
  if (!hit) return [a];
  const out: Rect[] = [];
  const aBottom = a.row + a.rows;
  const aRight = a.col + a.cols;
  const hBottom = hit.row + hit.rows;
  const hRight = hit.col + hit.cols;
  // Full-width band above, full-width band below, then left and right of the hole.
  if (hit.row > a.row) out.push({ row: a.row, col: a.col, rows: hit.row - a.row, cols: a.cols });
  if (hBottom < aBottom) out.push({ row: hBottom, col: a.col, rows: aBottom - hBottom, cols: a.cols });
  if (hit.col > a.col) out.push({ row: hit.row, col: a.col, rows: hit.rows, cols: hit.col - a.col });
  if (hRight < aRight) out.push({ row: hit.row, col: hRight, rows: hit.rows, cols: aRight - hRight });
  return out;
}

export function fitsSheet(rect: Rect): boolean {
  return (
    rect.row >= 0 &&
    rect.col >= 0 &&
    rect.row + rect.rows <= MAX_ROWS &&
    rect.col + rect.cols <= MAX_COLS
  );
}

/**
 * Split a rectangle into pieces of at most `maxCells` cells. Pieces are full-width
 * bands of rows; only when one row alone is over the limit is it cut by columns too.
 * Pieces come in reading order and cover the rectangle exactly once.
 */
export function chunkRect(rect: Rect, maxCells: number): Rect[] {
  if (rect.rows <= 0 || rect.cols <= 0) return [];
  const limit = Math.max(1, Math.floor(maxCells));
  const out: Rect[] = [];
  if (rect.cols <= limit) {
    const band = Math.max(1, Math.floor(limit / rect.cols));
    for (let r = 0; r < rect.rows; r += band) {
      out.push({
        row: rect.row + r,
        col: rect.col,
        rows: Math.min(band, rect.rows - r),
        cols: rect.cols,
      });
    }
    return out;
  }
  for (let r = 0; r < rect.rows; r++) {
    for (let c = 0; c < rect.cols; c += limit) {
      out.push({
        row: rect.row + r,
        col: rect.col + c,
        rows: 1,
        cols: Math.min(limit, rect.cols - c),
      });
    }
  }
  return out;
}

/**
 * Join addresses with commas into strings no longer than `maxLength`, for
 * `worksheet.getRanges()`. An address longer than the limit gets a string of its own.
 */
export function batchAddresses(addresses: string[], maxLength: number): string[] {
  const out: string[] = [];
  let current = "";
  for (const address of addresses) {
    if (current === "") {
      current = address;
    } else if (current.length + 1 + address.length <= maxLength) {
      current += "," + address;
    } else {
      out.push(current);
      current = address;
    }
  }
  if (current !== "") out.push(current);
  return out;
}

/**
 * Merge a list of flags into runs: calls `emit(start, length)` for every maximal run
 * of consecutive indexes whose key is the same and not null.
 */
export function runs<K>(
  count: number,
  keyAt: (index: number) => K | null,
  emit: (key: K, start: number, length: number) => void,
): void {
  let start = 0;
  let key: K | null = null;
  for (let i = 0; i <= count; i++) {
    const next = i < count ? keyAt(i) : null;
    if (i < count && next !== null && next === key) continue;
    if (key !== null) emit(key, start, i - start);
    key = next;
    start = i;
  }
}
