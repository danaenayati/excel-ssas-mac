import { describe, expect, it } from "vitest";
import type { Grid, HeaderCell, HeaderKind, LineKind } from "../types";
import type { Rect } from "./geometry";
import { blockFields, buildBlock } from "./layout";
import { PreviewHost } from "./preview";
import {
  DEFAULT_STYLE,
  PIVOT_STYLES,
  findStyle,
  isStyleName,
  paintBlock,
  styleChoice,
  styleSwatch,
  type BlockRegions,
  type BodyLine,
  type PaintLayer,
  type StyleChoice,
} from "./styles";

/**
 * A block of 1 filter row, a blank row, 1 header row, one row label column and three
 * value columns (the last a grand total column), with these lines.
 */
function regions(lines: BodyLine[], columns: LineKind[] = ["data", "data", "grandTotal"]): BlockRegions {
  return { cols: 4, filterRows: 1, headerTop: 2, dataTop: 3, headerCols: 1, lines, columns };
}

function choice(name: string, options: Partial<Omit<StyleChoice, "style">> = {}): StyleChoice {
  return { ...styleChoice(), style: findStyle(name), ...options };
}

/** The colour a cell ends up filled with: the last layer that fills it. */
function fillAt(paint: PaintLayer[], row: number, col: number): string | undefined {
  let out: string | undefined;
  for (const layer of paint) {
    if (layer.paint.fill && layer.rects.some((r) => inside(r, row, col))) out = layer.paint.fill;
  }
  return out;
}

function inside(r: Rect, row: number, col: number): boolean {
  return row >= r.row && row < r.row + r.rows && col >= r.col && col < r.col + r.cols;
}

const isBold = (bold: Rect[], row: number, col: number) => bold.some((r) => inside(r, row, col));

describe("the style list", () => {
  it("offers Excel's names, None first, with Light 16 as the default", () => {
    const names = PIVOT_STYLES.map((s) => s.name);
    expect(names[0]).toBe("None");
    expect(names).toEqual(
      expect.arrayContaining(["Light 1", "Light 8", "Light 9", "Light 15", "Light 16", "Medium 2", "Medium 9", "Medium 15", "Dark 2"]),
    );
    expect(new Set(names).size).toBe(names.length);
    expect(DEFAULT_STYLE).toBe("Light 16");
  });

  it("falls back to the default for no name or one it does not know", () => {
    expect(findStyle(undefined).name).toBe("Light 16");
    expect(findStyle("Medium 99").name).toBe("Light 16");
    expect(findStyle("Dark 2").name).toBe("Dark 2");
    expect(isStyleName("Medium 9")).toBe(true);
    expect(isStyleName("medium 9")).toBe(false);
    expect(isStyleName(3)).toBe(false);
  });

  it("reads Excel's defaults for options that are absent", () => {
    expect(styleChoice()).toMatchObject({ rowHeaders: true, columnHeaders: true, bandedRows: false, bandedColumns: false });
    expect(styleChoice().style.name).toBe("Light 16");
    const chosen = styleChoice({ style: "Medium 2", rowHeaders: false, columnHeaders: false, bandedRows: true, bandedColumns: true });
    expect(chosen).toMatchObject({ rowHeaders: false, columnHeaders: false, bandedRows: true, bandedColumns: true });
    expect(chosen.style.name).toBe("Medium 2");
  });

  it("gives every style a swatch of real colours", () => {
    for (const style of PIVOT_STYLES) {
      const swatch = styleSwatch(style);
      for (const colour of [swatch.header, swatch.body, swatch.line]) expect(colour).toMatch(/^#[0-9A-F]{6}$/);
    }
    expect(styleSwatch(findStyle("Medium 2")).header).toBe("#5B9BD5");
  });
});

describe("Light 16 with the default options (today's look)", () => {
  const out = paintBlock(regions(["data", "subheading", "data", "subtotal", "grandTotal"]), styleChoice());

  it("fills the filter captions and the header with light blue, in one layer, and nothing else", () => {
    expect(out.paint).toEqual([
      {
        paint: { fill: "#DDEBF7" },
        rects: [
          { row: 0, col: 0, rows: 1, cols: 1 },
          { row: 2, col: 0, rows: 1, cols: 4 },
        ],
      },
    ]);
  });

  it("makes captions, header and every total or outer item line bold across the width", () => {
    expect(out.bold).toEqual([
      { row: 0, col: 0, rows: 1, cols: 1 },
      { row: 2, col: 0, rows: 1, cols: 4 },
      { row: 4, col: 0, rows: 1, cols: 4 },
      { row: 6, col: 0, rows: 2, cols: 4 },
    ]);
  });
});

describe("Style Options", () => {
  const lines: BodyLine[] = ["subheading", "data", "data", "subheading", "data", "grandTotal"];

  it("Column Headers off: the header is neither filled nor bold; the filter captions still are", () => {
    const out = paintBlock(regions(lines), choice("Light 16", { columnHeaders: false }));
    expect(fillAt(out.paint, 2, 1)).toBeUndefined();
    expect(isBold(out.bold, 2, 1)).toBe(false);
    expect(fillAt(out.paint, 0, 0)).toBe("#DDEBF7");
    expect(isBold(out.bold, 0, 0)).toBe(true);
  });

  it("Row Headers off: outer items' labels are plain, their subtotal values stay bold", () => {
    const on = paintBlock(regions(lines), choice("Light 16"));
    const off = paintBlock(regions(lines), choice("Light 16", { rowHeaders: false }));
    expect(isBold(on.bold, 3, 0)).toBe(true);
    expect(isBold(off.bold, 3, 0)).toBe(false);
    expect(isBold(off.bold, 3, 1)).toBe(true);
    // The grand total keeps its own look.
    expect(isBold(off.bold, 8, 0)).toBe(true);
  });

  it("Row Headers also switches the first column's own look (Medium 9 tints it)", () => {
    expect(fillAt(paintBlock(regions(lines), choice("Medium 9")).paint, 4, 0)).toBe("#BDD7EE");
    expect(fillAt(paintBlock(regions(lines), choice("Medium 9", { rowHeaders: false })).paint, 4, 0)).toBe("#DDEBF7");
  });

  it("Banded Rows: every other line from the first, the grand total excluded", () => {
    const out = paintBlock(regions(["data", "data", "data", "data", "grandTotal"]), choice("Light 16", { bandedRows: true }));
    expect([3, 4, 5, 6, 7].map((row) => fillAt(out.paint, row, 2))).toEqual(["#DDEBF7", undefined, "#DDEBF7", undefined, undefined]);
  });

  it("Banded Columns: every other value column, data rows only", () => {
    const out = paintBlock(regions(["data", "data"]), choice("Light 1", { bandedColumns: true }));
    expect([1, 2, 3].map((col) => fillAt(out.paint, 3, col))).toEqual(["#F2F2F2", undefined, "#F2F2F2"]);
    expect(fillAt(out.paint, 2, 1)).toBeUndefined();
    expect(fillAt(out.paint, 3, 0)).toBeUndefined();
  });

  it("colours an outer item's whole line as a subtotal line, label included", () => {
    const out = paintBlock(regions(["subheading", "data"]), choice("Dark 2", { rowHeaders: false }));
    expect(fillAt(out.paint, 3, 0)).toBe("#2F75B5");
    expect(fillAt(out.paint, 3, 2)).toBe("#2F75B5");
    expect(isBold(out.bold, 3, 0)).toBe(false);
  });

  it("totals win over stripes, as in Excel", () => {
    const out = paintBlock(regions(["subtotal", "data", "grandTotal"]), choice("Medium 2", { bandedRows: true }));
    expect(fillAt(out.paint, 3, 1)).toBe("#DDEBF7"); // subtotal (and first stripe)
    expect(fillAt(out.paint, 5, 1)).toBe("#BDD7EE"); // grand total
  });
});

describe("the other styles", () => {
  const lines: BodyLine[] = ["subheading", "data", "subtotal", "grandTotal"];

  it("None: no fill, no colour, no bold", () => {
    const out = paintBlock(regions(lines), choice("None", { bandedRows: true, bandedColumns: true }));
    expect(out.paint).toEqual([]);
    expect(out.bold).toEqual([]);
  });

  it("Light 9: solid blue header with white text, a blue outline, a double line over the grand total", () => {
    const out = paintBlock(regions(lines), choice("Light 9"));
    expect(out.paint).toContainEqual({ paint: { fill: "#5B9BD5", color: "#FFFFFF" }, rects: [{ row: 2, col: 0, rows: 1, cols: 4 }] });
    expect(out.paint[0]).toEqual({
      paint: {
        borders: {
          top: { color: "#5B9BD5" },
          bottom: { color: "#5B9BD5" },
          left: { color: "#5B9BD5" },
          right: { color: "#5B9BD5" },
        },
      },
      rects: [{ row: 2, col: 0, rows: 5, cols: 4 }],
    });
    expect(out.paint[out.paint.length - 1]).toEqual({
      paint: { borders: { top: { color: "#5B9BD5", style: "double" } } },
      rects: [{ row: 6, col: 0, rows: 1, cols: 4 }],
    });
  });

  it("draws a line over each subtotal line, not over a run of them", () => {
    const out = paintBlock(regions(["subtotal", "subtotal", "data"]), choice("Light 8"));
    const lined = out.paint.find((l) => l.paint.borders && l.paint.borders.top && !l.paint.borders.left);
    expect(lined?.rects).toEqual([
      { row: 3, col: 0, rows: 1, cols: 4 },
      { row: 4, col: 0, rows: 1, cols: 4 },
    ]);
  });

  it("drops lines between rows where a rect has one row", () => {
    const out = paintBlock({ ...regions([]), headerTop: 0, dataTop: 1, filterRows: 0 }, choice("Medium 9"));
    expect(out.paint[0].paint.borders).toEqual({ insideV: { color: "#FFFFFF" } });
  });

  it("Dark 2: white text on blue over the whole table, the grand total column bold", () => {
    const out = paintBlock(regions(lines), choice("Dark 2"));
    // The filter captions look the same, so they share the table's layer.
    expect(out.paint[0]).toEqual({
      paint: { fill: "#5B9BD5", color: "#FFFFFF" },
      rects: [
        { row: 2, col: 0, rows: 5, cols: 4 },
        { row: 0, col: 0, rows: 1, cols: 1 },
      ],
    });
    expect(isBold(out.bold, 4, 3)).toBe(true);
    expect(isBold(out.bold, 4, 2)).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// Through buildBlock and the preview.
// ---------------------------------------------------------------------------

function cell(text: string, kind: HeaderKind = "member"): HeaderCell {
  return { text, indent: 0, kind, member: null };
}

/** One row field, two measures, lines 2025, 2026, Grand Total. */
function grid(): Grid {
  return {
    filterRows: [],
    rowFieldCaptions: ["Year"],
    colHeaderRows: [[cell("Rides", "measure"), cell("Fare", "measure")]],
    rowHeaders: [[cell("2025")], [cell("2026")], [cell("Grand Total", "grandTotal")]],
    values: [
      [1, 2],
      [3, 4],
      [4, 6],
    ],
    formatPalette: ["General"],
    cellFormats: [
      [0, 0],
      [0, 0],
      [0, 0],
    ],
    rowKinds: ["data", "data", "grandTotal"],
    colKinds: ["data", "data"],
    colValueIndex: [0, 1],
    stats: { rows: 3, columns: 2, elapsedMs: 1, mdx: "" },
  };
}

const layout = (options: Record<string, unknown>) => ({
  filters: [],
  rows: [],
  columns: [],
  options: { grandTotals: true, subtotals: false, nonEmpty: true, ...options },
});

describe("a report's style options reach the block", () => {
  it("draws the default when the report says nothing", () => {
    const plain = buildBlock(grid(), blockFields(layout({})));
    expect(plain.paint).toEqual([{ paint: { fill: "#DDEBF7" }, rects: [{ row: 0, col: 0, rows: 1, cols: 3 }] }]);
  });

  it("draws the style and options the report asks for", () => {
    const block = buildBlock(grid(), blockFields(layout({ style: "Medium 2", bandedRows: true })));
    expect(fillAt(block.paint, 0, 0)).toBe("#5B9BD5");
    expect(fillAt(block.paint, 1, 1)).toBe("#DDEBF7");
    expect(fillAt(block.paint, 2, 1)).toBeUndefined();
    expect(fillAt(block.paint, 3, 1)).toBe("#BDD7EE");
  });

  it("the preview shows fills, font colours and lines on the cells", async () => {
    const host = new PreviewHost({ persist: false });
    const block = buildBlock(grid(), blockFields(layout({ style: "Light 9" })));
    await host.writeReport("r1", { sheetId: "preview-sheet-1", sheetName: "Sheet1", row: 0, col: 0 }, block);
    const cells = host.snapshot().cells;
    expect(cells.get("0,1")).toMatchObject({ fill: "#5B9BD5", color: "#FFFFFF", bold: true });
    expect(cells.get("0,0")?.borders).toMatchObject({ top: { color: "#5B9BD5" }, left: { color: "#5B9BD5" } });
    expect(cells.get("3,2")?.borders).toMatchObject({
      top: { color: "#5B9BD5", style: "double" },
      bottom: { color: "#5B9BD5" },
      right: { color: "#5B9BD5" },
    });
    expect(cells.get("1,1")?.fill).toBeUndefined();
  });
});
