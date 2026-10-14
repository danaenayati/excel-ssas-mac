// Excel's PivotTable Styles and Style Options (Design tab), as the pane draws them.
//
// A style says how each part of a report looks: the page filters, the header rows, the
// row labels, the stripes of Banded Rows and Banded Columns, the subtotal and grand total
// lines and columns. `paintBlock` turns the parts of a block into what to make bold and
// what to colour, as layers in Excel's order of precedence: where two parts set the same
// thing, the later one wins (a grand total's fill over a stripe's).
//
// The colours are the Office theme's (Accent 1 blue, and the greys of "Dark 1"), with the
// tints and shades Excel's own styles use; the names are Excel's. "Light 16" is Excel's
// default and the pane's: with the default Style Options it draws exactly what reports
// drew before styles existed (light blue header and filter captions, bold totals).

import type { LayoutOptions, LineKind } from "../types";
import type { Rect } from "./geometry";

export interface BorderLine {
  color: string;
  /** Default "thin". */
  style?: "thin" | "double";
}

/** A rectangle's outer edges, and the lines between its rows ("insideH") or columns. */
export type Edge = "top" | "bottom" | "left" | "right" | "insideH" | "insideV";
export const EDGES: readonly Edge[] = ["top", "bottom", "left", "right", "insideH", "insideV"];

/** What a layer sets on its cells; whatever it leaves out stays as the layers below set it. */
export interface Paint {
  fill?: string;
  /** Font colour. */
  color?: string;
  borders?: Partial<Record<Edge, BorderLine>>;
}

/** A part of a style: its paint, and whether its text is bold. */
export interface Look extends Paint {
  bold?: boolean;
}

/** Cells to paint the same way. Layers are applied in order. */
export interface PaintLayer {
  paint: Paint;
  rects: Rect[];
}

/** One of Excel's PivotTable Styles. Each part is Excel's "table element" of that name. */
export interface PivotStyle {
  /** Excel's name: "Light 16", "Medium 2", … */
  name: string;
  group: "None" | "Light" | "Medium" | "Dark";
  /** The whole table: header rows and body (not the page filters). */
  table?: Look;
  /** "Report Filter Labels" and "Report Filter Values". */
  filterLabels?: Look;
  filterValues?: Look;
  /** The row label column(s) of the body: drawn only with Row Headers on. */
  firstColumn?: Look;
  /** The header rows: drawn only with Column Headers on. */
  header?: Look;
  /** First row and column stripes, for Banded Rows and Banded Columns. */
  rowStripe?: Look;
  columnStripe?: Look;
  subtotalColumn?: Look;
  grandTotalColumn?: Look;
  /**
   * "Tehran Total" lines, and outer items' own lines (which show the subtotal when
   * subtotals are at the top). Its bold covers the values only on an outer item's line:
   * the label's bold is `rowSubheading`'s.
   */
  subtotalRow?: Look;
  /** The label of an outer item's own line (compact and outline form): drawn only with Row Headers on. */
  rowSubheading?: Look;
  grandTotalRow?: Look;
}

const WHITE = "#FFFFFF";
const BLACK = "#000000";
// Accent 1 of the Office theme and the tints and shades Excel's styles take from it.
const BLUE = "#5B9BD5";
const BLUE_80 = "#DDEBF7";
const BLUE_60 = "#BDD7EE";
const BLUE_DARK_25 = "#2F75B5";
const BLUE_DARK_50 = "#1F4E78";
// "Dark 1" (black), lightened as Excel does for its black-and-white styles.
const GREY_95 = "#F2F2F2";
const GREY_85 = "#D9D9D9";
const GREY_65 = "#A6A6A6";
const GREY_25 = "#404040";

const thin = (color: string): BorderLine => ({ color });
const double = (color: string): BorderLine => ({ color, style: "double" });
const outline = (color: string): Look["borders"] => ({
  top: thin(color),
  bottom: thin(color),
  left: thin(color),
  right: thin(color),
});

/** Light 1 to 7: no fills, a line under the header and over the grand total. */
function lightLines(name: string, line: string, stripe: string): PivotStyle {
  return {
    name,
    group: "Light",
    filterLabels: { bold: true },
    header: { bold: true, borders: { bottom: thin(line) } },
    rowStripe: { fill: stripe },
    columnStripe: { fill: stripe },
    subtotalRow: { bold: true },
    rowSubheading: { bold: true },
    grandTotalRow: { bold: true, borders: { top: thin(line) } },
  };
}

/** Light 8 to 14: a solid header with white text, the table outlined, lines over totals. */
function lightSolidHeader(name: string, color: string, stripe: string): PivotStyle {
  return {
    name,
    group: "Light",
    table: { borders: outline(color) },
    filterLabels: { bold: true },
    header: { bold: true, fill: color, color: WHITE },
    rowStripe: { fill: stripe },
    columnStripe: { fill: stripe },
    subtotalRow: { bold: true, borders: { top: thin(color) } },
    rowSubheading: { bold: true },
    grandTotalRow: { bold: true, borders: { top: double(color) } },
  };
}

/** Light 15 to 21: a light header and light filter captions, bold totals. */
function lightTinted(name: string, tint: string): PivotStyle {
  return {
    name,
    group: "Light",
    filterLabels: { bold: true, fill: tint },
    header: { bold: true, fill: tint },
    rowStripe: { fill: tint },
    columnStripe: { fill: tint },
    subtotalRow: { bold: true },
    rowSubheading: { bold: true },
    grandTotalRow: { bold: true },
  };
}

/** The styles the pane offers, in Excel's gallery order. */
export const PIVOT_STYLES: readonly PivotStyle[] = [
  { name: "None", group: "None" },
  lightLines("Light 1", BLACK, GREY_95),
  lightSolidHeader("Light 8", BLACK, GREY_95),
  lightSolidHeader("Light 9", BLUE, BLUE_80),
  { ...lightTinted("Light 15", GREY_85), rowStripe: { fill: GREY_95 }, columnStripe: { fill: GREY_95 } },
  lightTinted("Light 16", BLUE_80),
  {
    name: "Medium 2",
    group: "Medium",
    filterLabels: { bold: true, fill: BLUE_80 },
    header: { bold: true, fill: BLUE, color: WHITE },
    rowStripe: { fill: BLUE_80 },
    columnStripe: { fill: BLUE_80 },
    grandTotalColumn: { bold: true },
    subtotalRow: { bold: true, fill: BLUE_80 },
    rowSubheading: { bold: true },
    grandTotalRow: { bold: true, fill: BLUE_60, borders: { top: double(BLUE) } },
  },
  {
    name: "Medium 9",
    group: "Medium",
    table: { fill: BLUE_80, borders: { insideH: thin(WHITE), insideV: thin(WHITE) } },
    filterLabels: { bold: true, fill: BLUE_60 },
    firstColumn: { fill: BLUE_60 },
    header: { bold: true, fill: BLUE, color: WHITE },
    rowStripe: { fill: BLUE_60 },
    columnStripe: { fill: BLUE_60 },
    grandTotalColumn: { bold: true },
    subtotalRow: { bold: true, fill: BLUE_60 },
    rowSubheading: { bold: true },
    grandTotalRow: { bold: true, fill: BLUE, color: WHITE },
  },
  {
    name: "Medium 15",
    group: "Medium",
    table: { borders: { ...outline(BLACK), insideH: thin(GREY_65) } },
    filterLabels: { bold: true, fill: GREY_85 },
    header: { bold: true, fill: GREY_25, color: WHITE },
    rowStripe: { fill: GREY_95 },
    columnStripe: { fill: GREY_95 },
    grandTotalColumn: { bold: true },
    subtotalRow: { bold: true, fill: GREY_85 },
    rowSubheading: { bold: true },
    grandTotalRow: { bold: true, fill: GREY_65, borders: { top: double(BLACK) } },
  },
  {
    name: "Dark 2",
    group: "Dark",
    table: { fill: BLUE, color: WHITE },
    filterLabels: { bold: true, fill: BLUE, color: WHITE },
    header: { bold: true, fill: BLUE_DARK_50, color: WHITE, borders: { bottom: thin(WHITE) } },
    rowStripe: { fill: BLUE_DARK_25 },
    columnStripe: { fill: BLUE_DARK_25 },
    grandTotalColumn: { bold: true },
    subtotalRow: { bold: true, fill: BLUE_DARK_25 },
    rowSubheading: { bold: true },
    grandTotalRow: { bold: true, fill: BLUE_DARK_50, borders: { top: double(WHITE) } },
  },
];

/** Excel's default PivotTable style, and what a report without a style is drawn in. */
export const DEFAULT_STYLE = "Light 16";

export function isStyleName(value: unknown): value is string {
  return typeof value === "string" && PIVOT_STYLES.some((style) => style.name === value);
}

/** The style of that name; the default for none or a name the pane does not know. */
export function findStyle(name: string | null | undefined): PivotStyle {
  const found = PIVOT_STYLES.find((style) => style.name === name);
  return found ?? (PIVOT_STYLES.find((style) => style.name === DEFAULT_STYLE) as PivotStyle);
}

/** A style with the four Style Options, as a report asks for them. */
export interface StyleChoice {
  style: PivotStyle;
  rowHeaders: boolean;
  columnHeaders: boolean;
  bandedRows: boolean;
  bandedColumns: boolean;
}

/** The style options a report's settings say. Absent: Light 16, headers on, bands off (Excel's defaults). */
export function styleChoice(
  options?: Pick<LayoutOptions, "style" | "rowHeaders" | "columnHeaders" | "bandedRows" | "bandedColumns">,
): StyleChoice {
  const o = options ?? {};
  return {
    style: findStyle(o.style),
    rowHeaders: o.rowHeaders !== false,
    columnHeaders: o.columnHeaders !== false,
    bandedRows: o.bandedRows === true,
    bandedColumns: o.bandedColumns === true,
  };
}

/**
 * The colours of a small sample of a style for the pane: its header, its body and its
 * line or accent colour.
 */
export function styleSwatch(style: PivotStyle): { header: string; body: string; line: string } {
  const body = (style.table && style.table.fill) || WHITE;
  const header = (style.header && style.header.fill) || body;
  const borders = (style.header && style.header.borders) || {};
  const line =
    (borders.bottom && borders.bottom.color) ||
    (style.grandTotalRow && style.grandTotalRow.fill) ||
    (style.header && style.header.fill) ||
    GREY_85;
  return { header, body, line };
}

// ---------------------------------------------------------------------------
// Painting a block.
// ---------------------------------------------------------------------------

/**
 * A data line of a block, as styles see it: an item's line; an outer item's own line in
 * outline and compact form ("subheading", showing its subtotal when subtotals are at the
 * top); a subtotal line ("Tehran Total"); the grand total.
 */
export type BodyLine = "data" | "subheading" | "subtotal" | "grandTotal";

/** Where the parts of a block are. Positions are relative to the block. */
export interface BlockRegions {
  cols: number;
  /** F page filter rows at the top: caption in column 0, value in column 1. */
  filterRows: number;
  /** First row of the header area (the "Column Labels" row, if any, included). */
  headerTop: number;
  dataTop: number;
  /** Row label columns at the left. */
  headerCols: number;
  /** One per data row. */
  lines: BodyLine[];
  /** One per value column, right of the row labels. */
  columns: LineKind[];
}

function paintOf(look: Look, rect: Rect): Paint | null {
  const paint: Paint = {};
  if (look.fill) paint.fill = look.fill;
  if (look.color) paint.color = look.color;
  if (look.borders) {
    const borders: Partial<Record<Edge, BorderLine>> = {};
    for (const edge of EDGES) {
      const line = look.borders[edge];
      // Lines between rows or columns only where there are several.
      if (!line || (edge === "insideH" && rect.rows < 2) || (edge === "insideV" && rect.cols < 2)) continue;
      borders[edge] = line;
    }
    if (Object.keys(borders).length > 0) paint.borders = borders;
  }
  return Object.keys(paint).length > 0 ? paint : null;
}

function samePaint(a: Paint, b: Paint): boolean {
  return JSON.stringify(a) === JSON.stringify(b);
}

/** Runs of consecutive indexes with the same non-null key. */
function keyRuns<K>(count: number, keyAt: (index: number) => K | null, same: (a: K, b: K) => boolean): { key: K; start: number; length: number }[] {
  const out: { key: K; start: number; length: number }[] = [];
  for (let i = 0; i < count; i++) {
    const key = keyAt(i);
    const last = out[out.length - 1];
    if (key === null) continue;
    if (last && last.start + last.length === i && same(last.key, key)) last.length++;
    else out.push({ key, start: i, length: 1 });
  }
  return out;
}

/**
 * What to make bold and what to paint, for a block with these parts in this style.
 * `bold` lists the header first, then the lines (neighbours merged), then columns.
 */
export function paintBlock(regions: BlockRegions, choice: StyleChoice): { bold: Rect[]; paint: PaintLayer[] } {
  const { cols, filterRows: F, headerTop: top, dataTop, headerCols: RS, lines, columns } = regions;
  const style = choice.style;
  const N = lines.length;
  const C = columns.length;
  const valueCols = Math.max(0, cols - RS);
  const bold: Rect[] = [];
  const paint: PaintLayer[] = [];

  const layer = (look: Look | undefined, rects: Rect[]) => {
    if (!look) return;
    for (const rect of rects) {
      if (rect.rows <= 0 || rect.cols <= 0) continue;
      const p = paintOf(look, rect);
      if (!p) continue;
      const last = paint[paint.length - 1];
      if (last && samePaint(last.paint, p)) last.rects.push(rect);
      else paint.push({ paint: p, rects: [rect] });
    }
  };
  const boldArea = (look: Look | undefined, rect: Rect) => {
    if (look && look.bold && rect.rows > 0 && rect.cols > 0) bold.push(rect);
  };
  const whole = (n: number, length = 1): Rect => ({ row: dataTop + n, col: 0, rows: length, cols });
  const labels = (n: number, length = 1): Rect => ({ row: dataTop + n, col: 0, rows: length, cols: RS });
  const values = (n: number, length = 1): Rect => ({ row: dataTop + n, col: RS, rows: length, cols: valueCols });
  /** Lines of one kind, neighbours merged unless the look draws a line per row. */
  const lineRects = (look: Look | undefined, pick: (line: BodyLine) => "whole" | "labels" | "values" | null): Rect[] => {
    if (!look) return [];
    const perRow = Boolean(look.borders && (look.borders.top || look.borders.bottom));
    const out: Rect[] = [];
    for (const run of keyRuns(N, (n) => pick(lines[n]), (a, b) => a === b && !perRow)) {
      out.push(run.key === "whole" ? whole(run.start, run.length) : run.key === "labels" ? labels(run.start, run.length) : values(run.start, run.length));
    }
    return out;
  };
  const columnRects = (kind: LineKind): Rect[] =>
    N === 0
      ? []
      : keyRuns(C, (c) => (columns[c] === kind ? true : null), () => true).map((run) => ({
          row: dataTop,
          col: RS + run.start,
          rows: N,
          cols: run.length,
        }));

  // In Excel's order of precedence, lowest first.
  const table: Rect = { row: top, col: 0, rows: dataTop + N - top, cols };
  boldArea(style.table, table);
  layer(style.table, [table]);

  const filterLabels: Rect = { row: 0, col: 0, rows: F, cols: 1 };
  const filterValues: Rect = { row: 0, col: 1, rows: F, cols: 1 };
  boldArea(style.filterLabels, filterLabels);
  boldArea(style.filterValues, filterValues);
  layer(style.filterLabels, [filterLabels]);
  layer(style.filterValues, [filterValues]);

  if (choice.bandedColumns && N > 0) {
    const stripes: Rect[] = [];
    for (let c = 0; c < C; c += 2) stripes.push({ row: dataTop, col: RS + c, rows: N, cols: 1 });
    layer(style.columnStripe, stripes);
  }
  const striped: boolean[] = [];
  if (choice.bandedRows) {
    let k = 0;
    for (let n = 0; n < N; n++) striped.push(lines[n] !== "grandTotal" && k++ % 2 === 0);
    layer(
      style.rowStripe,
      striped.map((on, n) => (on ? whole(n) : null)).filter((rect): rect is Rect => rect !== null),
    );
  }
  if (choice.rowHeaders && N > 0) layer(style.firstColumn, [labels(0, N)]);

  const header: Rect = { row: top, col: 0, rows: dataTop - top, cols };
  if (choice.columnHeaders) {
    boldArea(style.header, header);
    layer(style.header, [header]);
  }

  layer(style.subtotalColumn, columnRects("subtotal"));
  layer(style.grandTotalColumn, columnRects("grandTotal"));
  // The colours of a subtotal line cover the whole line, an outer item's own line too;
  // the bold of that item's label is the row subheading's (Row Headers).
  layer(
    style.subtotalRow,
    lineRects(style.subtotalRow, (line) => (line === "subtotal" || line === "subheading" ? "whole" : null)),
  );
  if (choice.rowHeaders) {
    layer(style.rowSubheading, lineRects(style.rowSubheading, (line) => (line === "subheading" ? "labels" : null)));
  }
  layer(style.grandTotalRow, lineRects(style.grandTotalRow, (line) => (line === "grandTotal" ? "whole" : null)));

  // Bold lines: the label part and the value part of each line, from every part of the
  // style that covers them; neighbours with the same bold parts merged.
  const on = (look: Look | undefined) => Boolean(look && look.bold);
  const boldKey = (n: number): "whole" | "labels" | "values" | null => {
    const line = lines[n];
    const stripe = Boolean(striped[n]) && on(style.rowStripe);
    const label =
      stripe ||
      (choice.rowHeaders && on(style.firstColumn)) ||
      (line === "subtotal" && on(style.subtotalRow)) ||
      (line === "subheading" && choice.rowHeaders && on(style.rowSubheading)) ||
      (line === "grandTotal" && on(style.grandTotalRow));
    const value =
      valueCols > 0 &&
      (stripe ||
        ((line === "subtotal" || line === "subheading") && on(style.subtotalRow)) ||
        (line === "grandTotal" && on(style.grandTotalRow)));
    return label && value ? "whole" : label ? "labels" : value ? "values" : null;
  };
  for (const run of keyRuns(N, boldKey, (a, b) => a === b)) {
    bold.push(run.key === "whole" ? whole(run.start, run.length) : run.key === "labels" ? labels(run.start, run.length) : values(run.start, run.length));
  }
  if (on(style.subtotalColumn)) bold.push(...columnRects("subtotal"));
  if (on(style.grandTotalColumn)) bold.push(...columnRects("grandTotal"));
  return { bold, paint };
}
