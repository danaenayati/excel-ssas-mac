// Excel's Report Layout (Compact, Outline, Tabular Form) and the Subtotals position, on
// grids built the way the helper builds them (helper/src/cubebridge/pivot/grid.py,
// `_axis_lines`): one line per tuple, an outer label blanked when the previous tuple
// has the same item, "X Total" on a subtotal and "Grand Total" on the grand total.

import { describe, expect, it } from "vitest";
import type { Grid, HeaderCell, LayoutOptions, LineKind, MemberRef, PlacedField } from "../types";
import {
  COLUMN_LABELS,
  FILTERED_FORMAT,
  FILTER_FORMAT,
  GENERAL_FORMAT,
  ROW_LABELS,
  TEXT_FORMAT,
  blockFields,
  buildBlock,
  columnHeaderAt,
  excelValues,
  fieldCellAt,
  guessFieldCell,
  labelsCaption,
  memberAt,
  valueCellMembers,
  type Block,
  type BlockFields,
} from "./layout";

const CITY = "[Geography].[City]";
const SERVICE = "[Service].[Service Type]";
const YEAR = "[Date].[Year]";
const CAL = "[Date].[Calendar]";

function ref(hierarchy: string, caption: string, level = 1, flags: Partial<MemberRef> = {}): MemberRef {
  return {
    uniqueName: `${hierarchy}.&[${caption}]`,
    hierarchy,
    caption,
    level,
    hasChildren: false,
    expanded: false,
    ...flags,
  };
}

const tehran = ref(CITY, "Tehran");
const shiraz = ref(CITY, "Shiraz");
const ride = ref(SERVICE, "Ride");
const bike = ref(SERVICE, "Bike");
const y2025 = ref(YEAR, "2025");
const y2026 = ref(YEAR, "2026");

/** A tuple: one member per field, null for the field's All member. */
type Tuple = (MemberRef | null)[];

/** One header line per tuple, exactly as the helper's `_axis_lines` writes them. */
function helperLines(tuples: Tuple[]): { lines: HeaderCell[][]; kinds: LineKind[] } {
  const key = (m: MemberRef | null) => (m ? m.uniqueName : "(All)");
  const same = (a: Tuple, b: Tuple, count: number) => a.slice(0, count).every((m, j) => key(m) === key(b[j]));
  const lines: HeaderCell[][] = [];
  const kinds: LineKind[] = [];
  let previous: Tuple | null = null;
  for (const tuple of tuples) {
    const firstAll = tuple.findIndex((m) => m === null);
    const kind: LineKind = firstAll < 0 ? "data" : firstAll === 0 ? "grandTotal" : "subtotal";
    const repeat = previous !== null && same(previous, tuple, tuple.length);
    lines.push(
      tuple.map((m, i): HeaderCell => {
        if (kind === "grandTotal") {
          return { text: i === 0 && !repeat ? "Grand Total" : "", indent: 0, kind, member: null };
        }
        if (kind === "subtotal" && i >= firstAll - 1) {
          const text = i === firstAll - 1 && !repeat ? `${(tuple[i] as MemberRef).caption} Total` : "";
          return { text, indent: 0, kind, member: null };
        }
        const member = m as MemberRef;
        const suppressed = previous !== null && same(previous, tuple, i + 1);
        const indent = Math.max(0, member.level - 1);
        return { text: suppressed ? "" : member.caption, indent, kind: suppressed ? "blank" : "member", member };
      }),
    );
    kinds.push(kind);
    previous = tuple;
  }
  return { lines, kinds };
}

/** City > Service Type on rows, subtotals where asked; the grand total last. */
function cityServiceTuples(subtotals: "top" | "bottom" | "none"): Tuple[] {
  const out: Tuple[] = [];
  for (const city of [tehran, shiraz]) {
    if (subtotals === "top") out.push([city, null]);
    out.push([city, ride], [city, bike]);
    if (subtotals === "bottom") out.push([city, null]);
  }
  out.push([null, null]);
  return out;
}

interface GridSpec {
  rows: Tuple[];
  rowCaptions: string[];
  /** Column tuples; each is shown once per measure. */
  columns?: Tuple[];
  measures?: string[];
  filters?: { caption: string; value: string }[];
  /** Value of row line g, column c; default (g + 1) * 100 + c. */
  value?: (g: number, c: number) => number | null;
}

/** A grid as the helper answers it: measures innermost on columns, one header row per column field, then the measures. */
function helperGrid(spec: GridSpec): Grid {
  const measures = spec.measures ?? ["Rides"];
  const colTuples = spec.columns ?? [[]];
  const expanded: Tuple[] = [];
  for (const tuple of colTuples) for (let m = 0; m < measures.length; m++) expanded.push(tuple);
  const fieldRows = colTuples[0].length;
  const cols = helperLines(expanded);
  const colHeaderRows: HeaderCell[][] = [];
  for (let r = 0; r < fieldRows; r++) colHeaderRows.push(cols.lines.map((line) => line[r]));
  colHeaderRows.push(expanded.map((_, c) => ({ text: measures[c % measures.length], indent: 0, kind: "measure", member: null })));
  const rows = helperLines(spec.rows);
  const value = spec.value ?? ((g: number, c: number) => (g + 1) * 100 + c);
  const C = expanded.length;
  return {
    filterRows: spec.filters ?? [],
    rowFieldCaptions: spec.rowCaptions,
    colHeaderRows,
    rowHeaders: rows.lines,
    values: rows.lines.map((_, g) => Array.from({ length: C }, (_, c) => value(g, c))),
    formatPalette: ["General", "#,##0"],
    cellFormats: rows.lines.map(() => new Array<number>(C).fill(1)),
    rowKinds: rows.kinds,
    colKinds: fieldRows > 0 ? cols.kinds : new Array<LineKind>(C).fill("data"),
    colValueIndex: expanded.map((_, c) => c % measures.length),
    stats: { rows: rows.lines.length, columns: C, elapsedMs: 1, mdx: "" },
  };
}

function placed(hierarchy: string, caption: string, extra: Partial<PlacedField> = {}): PlacedField {
  return { hierarchy, caption, allMember: `${hierarchy}.[All]`, expanded: [], members: [], memberCaptions: [], ...extra };
}

type Form = NonNullable<LayoutOptions["form"]>;
type Subtotals = "top" | "bottom" | "none";

function options(form: Form, subtotals: Subtotals, extra: Partial<LayoutOptions> = {}): LayoutOptions {
  return {
    grandTotals: true,
    nonEmpty: true,
    subtotals: subtotals !== "none",
    subtotalsAt: subtotals === "top" ? "top" : "bottom",
    form,
    ...extra,
  };
}

/** City > Service Type on rows, Year on columns (2025, 2026, Grand Total), Rides. */
function cityByYear(form: Form, subtotals: Subtotals, extra: Partial<LayoutOptions> = {}) {
  const layout = {
    filters: [] as PlacedField[],
    rows: [placed(CITY, "City"), placed(SERVICE, "Service Type")],
    columns: [placed(YEAR, "Year")],
    values: [{ measure: "[Measures].[Rides]", caption: "Rides" }],
    options: options(form, subtotals, extra),
  };
  // Tabular form never has subtotals at the top: the helper sends them at the bottom.
  const order = form === "tabular" && subtotals === "top" ? "bottom" : subtotals;
  const grid = helperGrid({
    rows: cityServiceTuples(order),
    rowCaptions: ["City", "Service Type"],
    columns: [[y2025], [y2026], [null]],
  });
  return { layout, grid, block: buildBlock(grid, blockFields(layout)) };
}

const column = (block: Block, col: number) => block.values.slice(block.shape.dataTop).map((line) => line[col]);
const indentAt = (block: Block, row: number, col: number) => {
  const group = block.indents.find((g) =>
    g.rects.some((r) => row >= r.row && row < r.row + r.rows && col >= r.col && col < r.col + r.cols),
  );
  return group ? group.level : 0;
};
const isBold = (block: Block, row: number, col: number) =>
  block.bold.some((r) => row >= r.row && row < r.row + r.rows && col >= r.col && col < r.col + r.cols);

describe("compact form", () => {
  it("puts every row field in one column under Row Labels, with Column Labels above the column items", () => {
    const { block } = cityByYear("compact", "top");
    expect(block.values.slice(0, 3)).toEqual([
      ["", COLUMN_LABELS, "", ""],
      ["", "2025", "2026", "Grand Total"],
      [ROW_LABELS, "Rides", "Rides", "Rides"],
    ]);
    expect(block.shape).toEqual({
      filterRows: 0,
      headerRows: 2,
      headerCols: 1,
      dataRows: 7,
      dataCols: 3,
      headerTop: 1,
      dataTop: 3,
      columnLabelsRow: 0,
    });
    expect(block.cols).toBe(4);
    expect(block.rows).toBe(10);
  });

  it("subtotals at the top: each outer item on its own line, with its subtotal's values there", () => {
    const { block } = cityByYear("compact", "top");
    // Grid lines: 0 Tehran Total, 1 Tehran/Ride, 2 Tehran/Bike, 3 Shiraz Total, 4, 5, 6 Grand Total.
    expect(block.values.slice(3)).toEqual([
      ["Tehran", 100, 101, 102],
      ["Ride", 200, 201, 202],
      ["Bike", 300, 301, 302],
      ["Shiraz", 400, 401, 402],
      ["Ride", 500, 501, 502],
      ["Bike", 600, 601, 602],
      ["Grand Total", 700, 701, 702],
    ]);
    expect(block.lines.map((line) => line.grid)).toEqual([0, 1, 2, 3, 4, 5, 6]);
    // No "Tehran Total" line: the subtotal is on Tehran's line.
    expect(column(block, 0)).not.toContain("Tehran Total");
  });

  it("indents inner items one level under their outer item", () => {
    const { block } = cityByYear("compact", "top");
    expect([3, 4, 5, 6, 7, 8, 9].map((row) => indentAt(block, row, 0))).toEqual([0, 1, 1, 0, 1, 1, 0]);
    const spaced = excelValues(block, true);
    expect(spaced[4][0]).toBe("   Ride");
    expect(spaced[3][0]).toBe("Tehran");
  });

  it("makes outer items' lines bold, like the totals, and leaves inner items plain", () => {
    const { block } = cityByYear("compact", "top");
    expect([3, 4, 5, 6, 7, 8, 9].map((row) => isBold(block, row, 2))).toEqual([true, false, false, true, false, false, true]);
    expect(block.bold).toContainEqual({ row: 0, col: 0, rows: 3, cols: 4 });
    expect(block.fill).toContainEqual({ row: 0, col: 0, rows: 3, cols: 4 });
  });

  it("subtotals at the bottom: the outer item's line is empty and '<item> Total' follows the group", () => {
    const { block } = cityByYear("compact", "bottom");
    expect(block.values.slice(3)).toEqual([
      ["Tehran", "", "", ""],
      ["Ride", 100, 101, 102],
      ["Bike", 200, 201, 202],
      ["Tehran Total", 300, 301, 302],
      ["Shiraz", "", "", ""],
      ["Ride", 400, 401, 402],
      ["Bike", 500, 501, 502],
      ["Shiraz Total", 600, 601, 602],
      ["Grand Total", 700, 701, 702],
    ]);
    expect(block.lines.map((line) => line.grid)).toEqual([null, 0, 1, 2, null, 3, 4, 5, 6]);
    // "Tehran Total" lines up with Tehran.
    expect(indentAt(block, 6, 0)).toBe(0);
    expect(isBold(block, 6, 1)).toBe(true);
    // An outer item's own line has no value cells to format.
    expect(block.numberFormats[3].slice(1)).toEqual([GENERAL_FORMAT, GENERAL_FORMAT, GENERAL_FORMAT]);
    expect(block.numberFormats[4].slice(1)).toEqual(["#,##0", "#,##0", "#,##0"]);
  });

  it("no subtotals: the outer item's line is empty and no total follows", () => {
    const { block } = cityByYear("compact", "none");
    expect(column(block, 0)).toEqual(["Tehran", "Ride", "Bike", "Shiraz", "Ride", "Bike", "Grand Total"]);
    expect(column(block, 1)).toEqual(["", 100, 200, "", 300, 400, 500]);
    expect(block.lines.map((line) => line.grid)).toEqual([null, 0, 1, null, 2, 3, 4]);
  });

  it("knows the member behind every row label, outer items included", () => {
    for (const subtotals of ["top", "bottom", "none"] as const) {
      const { block } = cityByYear("compact", subtotals);
      expect(memberAt(block, 3, 0)).toEqual({ member: tehran, axis: "rows" });
      expect(memberAt(block, 4, 0)).toEqual({ member: ride, axis: "rows" });
      expect(memberAt(block, 5, 0)).toEqual({ member: bike, axis: "rows" });
      expect(memberAt(block, 1, 1)).toEqual({ member: y2025, axis: "columns" });
      expect(memberAt(block, block.rows - 1, 0)).toBeNull(); // Grand Total
    }
    const bottom = cityByYear("compact", "bottom").block;
    expect(memberAt(bottom, 6, 0)).toBeNull(); // Tehran Total
    expect(memberAt(bottom, 7, 0)).toEqual({ member: shiraz, axis: "rows" });
  });

  it("gives Show Details the members of each line, a subtotal's own item included", () => {
    const top = cityByYear("compact", "top").block;
    expect(valueCellMembers(top, 0, 0)).toEqual([tehran.uniqueName, y2025.uniqueName]); // Tehran's own line
    expect(valueCellMembers(top, 1, 1)).toEqual([tehran.uniqueName, ride.uniqueName, y2026.uniqueName]);
    expect(valueCellMembers(top, 3, 2)).toEqual([shiraz.uniqueName]); // Shiraz, Grand Total column
    expect(valueCellMembers(top, 6, 2)).toEqual([]); // Grand Total of everything
    const bottom = cityByYear("compact", "bottom").block;
    expect(valueCellMembers(bottom, 0, 0)).toBeNull(); // Tehran's own line shows no values
    expect(valueCellMembers(bottom, 3, 0)).toEqual([tehran.uniqueName, y2025.uniqueName]); // Tehran Total
    expect(valueCellMembers(bottom, 99, 0)).toBeNull();
    expect(valueCellMembers(bottom, 1, 9)).toBeNull();
  });

  it("makes Row Labels and Column Labels filter cells for every field of their area", () => {
    const { block } = cityByYear("compact", "top");
    expect(fieldCellAt(block, 2, 0)).toEqual({ area: "rows", index: 0, labels: true });
    expect(fieldCellAt(block, 0, 1)).toEqual({ area: "columns", index: 0, labels: true });
    expect(block.numberFormats[2][0]).toBe(FILTER_FORMAT);
    expect(block.numberFormats[0][1]).toBe(FILTER_FORMAT);
    expect(block.fieldCells.size).toBe(2);
    // No per-field captions in compact form.
    expect(block.values[1][0]).toBe("");
    expect(excelValues(block, false)[2][0]).toBe("'" + ROW_LABELS);
    expect(labelsCaption("rows")).toBe(ROW_LABELS);
    expect(labelsCaption("columns")).toBe(COLUMN_LABELS);
  });

  it("shows the solid arrow when any field of the area is narrowed", () => {
    const { grid, layout } = cityByYear("compact", "top");
    layout.rows[1] = { ...layout.rows[1], members: [ride.uniqueName], memberCaptions: ["Ride"] };
    const rows = buildBlock(grid, blockFields(layout));
    expect(rows.numberFormats[2][0]).toBe(FILTERED_FORMAT);
    expect(rows.numberFormats[0][1]).toBe(FILTER_FORMAT);
    layout.columns[0] = { ...layout.columns[0], labelFilter: { operator: "equals", value: "2025" } };
    expect(buildBlock(grid, blockFields(layout)).numberFormats[0][1]).toBe(FILTERED_FORMAT);
  });

  it("offers the column field for its header cells, below the Column Labels row", () => {
    const { block } = cityByYear("compact", "top");
    expect(columnHeaderAt(block.shape, 1, 1, 1)).toBe(0);
    expect(columnHeaderAt(block.shape, 1, 3, 1)).toBe(0);
    expect(columnHeaderAt(block.shape, 0, 1, 1)).toBeNull(); // Column Labels row
    expect(columnHeaderAt(block.shape, 2, 1, 1)).toBeNull(); // measures
    expect(columnHeaderAt(block.shape, 1, 0, 1)).toBeNull(); // corner
  });

  it("ignores Repeat All Item Labels, as Excel greys it out", () => {
    const plain = cityByYear("compact", "none").block;
    const repeated = cityByYear("compact", "none", { repeatLabels: true }).block;
    expect(repeated.values).toEqual(plain.values);
  });

  it("writes 'For empty cells show' in empty value cells, not on an outer item's own line", () => {
    const { layout } = cityByYear("compact", "bottom", { emptyCellText: "n/a" });
    const grid = helperGrid({
      rows: cityServiceTuples("bottom"),
      rowCaptions: ["City", "Service Type"],
      columns: [[y2025], [y2026], [null]],
      value: (g, c) => (g === 1 && c === 0 ? null : 1),
    });
    const block = buildBlock(grid, blockFields(layout));
    expect(block.values[3].slice(1)).toEqual(["", "", ""]); // Tehran's own line
    expect(block.values[5][1]).toBe("n/a"); // Tehran / Bike in 2025
    expect(block.numberFormats[5][1]).toBe(TEXT_FORMAT);
  });

  it("keeps the filter rows and the blank row above Column Labels", () => {
    const { layout } = cityByYear("compact", "top");
    const withFilters = { ...layout, filters: [placed("[Customer].[Segment]", "Segment")] };
    const grid = helperGrid({
      rows: cityServiceTuples("top"),
      rowCaptions: ["City", "Service Type"],
      columns: [[y2025], [null]],
      filters: [{ caption: "Segment", value: "All" }],
    });
    const block = buildBlock(grid, blockFields(withFilters));
    expect(block.values[0].slice(0, 2)).toEqual(["Segment", "All"]);
    expect(block.values[1]).toEqual(["", "", ""]);
    expect(block.values[2]).toEqual(["", COLUMN_LABELS, ""]);
    expect(block.values[4][0]).toBe(ROW_LABELS);
    expect(block.shape).toMatchObject({ filterRows: 1, columnLabelsRow: 2, headerTop: 3, dataTop: 5 });
    expect(fieldCellAt(block, 0, 1)).toEqual({ area: "filters", index: 0 });
    expect(fieldCellAt(block, 2, 1)).toEqual({ area: "columns", index: 0, labels: true });
    expect(fieldCellAt(block, 4, 0)).toEqual({ area: "rows", index: 0, labels: true });
  });

  it("has no Column Labels row without column fields", () => {
    const layout = {
      filters: [],
      rows: [placed(CITY, "City"), placed(SERVICE, "Service Type")],
      columns: [],
      options: options("compact", "top"),
    };
    const grid = helperGrid({ rows: cityServiceTuples("top"), rowCaptions: ["City", "Service Type"], measures: ["Rides", "Fare"] });
    const block = buildBlock(grid, blockFields(layout));
    expect(block.values[0]).toEqual([ROW_LABELS, "Rides", "Fare"]);
    expect(block.shape.columnLabelsRow).toBeUndefined();
    expect(block.shape.dataTop).toBe(1);
    expect(block.values[1]).toEqual(["Tehran", 100, 101]);
    expect(fieldCellAt(block, 0, 0)).toEqual({ area: "rows", index: 0, labels: true });
  });

  it("has Column Labels but no Row Labels without row fields", () => {
    const layout = { filters: [], rows: [], columns: [placed(YEAR, "Year")], options: options("compact", "top") };
    const grid = helperGrid({ rows: [[]], rowCaptions: [""], columns: [[y2025], [null]] });
    grid.rowHeaders = [[{ text: "Total", indent: 0, kind: "grandTotal", member: null }]];
    const block = buildBlock(grid, blockFields(layout));
    expect(block.values).toEqual([
      ["", COLUMN_LABELS, ""],
      ["", "2025", "Grand Total"],
      ["", "Rides", "Rides"],
      ["Total", 100, 101],
    ]);
    expect(block.fieldCells.size).toBe(1);
    expect(fieldCellAt(block, 0, 1)).toEqual({ area: "columns", index: 0, labels: true });
    expect(valueCellMembers(block, 0, 0)).toEqual([y2025.uniqueName]);
  });

  it("is one column of rows for a single row field, under Row Labels", () => {
    const layout = { filters: [], rows: [placed(CITY, "City")], columns: [], options: options("compact", "top") };
    const grid = helperGrid({ rows: [[tehran], [shiraz], [null]], rowCaptions: ["City"] });
    const block = buildBlock(grid, blockFields(layout));
    expect(block.values).toEqual([
      [ROW_LABELS, "Rides"],
      ["Tehran", 100],
      ["Shiraz", 200],
      ["Grand Total", 300],
    ]);
    expect(block.lines).toEqual([
      { grid: 0, members: [tehran.uniqueName] },
      { grid: 1, members: [shiraz.uniqueName] },
      { grid: 2, members: [] },
    ]);
  });

  it("indents under a drilled-down item by its own level as well", () => {
    // Calendar (2025 open to quarters) > City, subtotals at the top.
    const cal2025 = ref(CAL, "2025", 1, { hasChildren: true, expanded: true });
    const q1 = ref(CAL, "Q1", 2, { hasChildren: true });
    const tuples: Tuple[] = [];
    for (const period of [cal2025, q1]) tuples.push([period, null], [period, tehran], [period, shiraz]);
    tuples.push([null, null]);
    const grid = helperGrid({ rows: tuples, rowCaptions: ["Calendar", "City"] });
    const layout = {
      filters: [],
      rows: [placed(CAL, "Calendar", { expanded: [cal2025.uniqueName] }), placed(CITY, "City")],
      columns: [],
      options: options("compact", "top"),
    };
    const block = buildBlock(grid, blockFields(layout));
    expect(column(block, 0)).toEqual(["2025", "Tehran", "Shiraz", "Q1", "Tehran", "Shiraz", "Grand Total"]);
    expect([1, 2, 3, 4, 5, 6, 7].map((row) => indentAt(block, row, 0))).toEqual([0, 1, 1, 1, 2, 2, 0]);
    expect(memberAt(block, 4, 0)).toEqual({ member: q1, axis: "rows" });
    expect(column(block, 1)).toEqual([100, 200, 300, 400, 500, 600, 700]);
  });

  it("gives each depth its own line with three row fields", () => {
    const a = ref("[A].[A]", "a1");
    const b1 = ref("[B].[B]", "b1");
    const b2 = ref("[B].[B]", "b2");
    const c1 = ref("[C].[C]", "c1");
    const c2 = ref("[C].[C]", "c2");
    const top: Tuple[] = [
      [a, null, null],
      [a, b1, null],
      [a, b1, c1],
      [a, b1, c2],
      [a, b2, null],
      [a, b2, c1],
      [null, null, null],
    ];
    const layout = {
      filters: [],
      rows: [placed("[A].[A]", "A"), placed("[B].[B]", "B"), placed("[C].[C]", "C")],
      columns: [],
      options: options("compact", "top"),
    };
    const block = buildBlock(helperGrid({ rows: top, rowCaptions: ["A", "B", "C"] }), blockFields(layout));
    expect(column(block, 0)).toEqual(["a1", "b1", "c1", "c2", "b2", "c1", "Grand Total"]);
    expect(column(block, 1)).toEqual([100, 200, 300, 400, 500, 600, 700]);
    expect([1, 2, 3, 4, 5, 6].map((row) => indentAt(block, row, 0))).toEqual([0, 1, 2, 2, 1, 2]);
    expect(block.lines[4].members).toEqual([a.uniqueName, b2.uniqueName]);

    const bottom: Tuple[] = [
      [a, b1, c1],
      [a, b1, c2],
      [a, b1, null],
      [a, b2, c1],
      [a, b2, null],
      [a, null, null],
      [null, null, null],
    ];
    const atBottom = buildBlock(helperGrid({ rows: bottom, rowCaptions: ["A", "B", "C"] }), {
      ...blockFields(layout),
      subtotalsAt: undefined,
    });
    expect(column(atBottom, 0)).toEqual(["a1", "b1", "c1", "c2", "b1 Total", "b2", "c1", "b2 Total", "a1 Total", "Grand Total"]);
    expect(column(atBottom, 1)).toEqual(["", "", 100, 200, 300, "", 400, 500, 600, 700]);
    expect([1, 2, 3, 4, 5, 6, 7, 8, 9].map((row) => indentAt(atBottom, row, 0))).toEqual([0, 1, 2, 2, 1, 1, 2, 1, 0]);
    expect(atBottom.lines[4].members).toEqual([a.uniqueName, b1.uniqueName]); // b1 Total
    expect(atBottom.lines[8].members).toEqual([a.uniqueName]); // a1 Total
  });

  it("writes a subtotal whose items are all gone as its own total line", () => {
    // Tehran's subtotal at the top, but none of its items came back.
    const tuples: Tuple[] = [[tehran, null], [shiraz, null], [shiraz, ride], [null, null]];
    const layout = {
      filters: [],
      rows: [placed(CITY, "City"), placed(SERVICE, "Service Type")],
      columns: [],
      options: options("compact", "top"),
    };
    const block = buildBlock(helperGrid({ rows: tuples, rowCaptions: ["City", "Service Type"] }), blockFields(layout));
    expect(column(block, 0)).toEqual(["Tehran Total", "Shiraz", "Ride", "Grand Total"]);
    expect(column(block, 1)).toEqual([100, 200, 300, 400]);
    expect(block.lines[0]).toEqual({ grid: 0, members: [] });
  });
});

describe("outline form", () => {
  it("keeps a column per row field, with the captions as headers", () => {
    const { block } = cityByYear("outline", "top");
    expect(block.values.slice(0, 2)).toEqual([
      ["", "Year", "2025", "2026", "Grand Total"],
      ["City", "Service Type", "Rides", "Rides", "Rides"],
    ]);
    expect(fieldCellAt(block, 1, 0)).toEqual({ area: "rows", index: 0 });
    expect(fieldCellAt(block, 1, 1)).toEqual({ area: "rows", index: 1 });
    expect(fieldCellAt(block, 0, 1)).toEqual({ area: "columns", index: 0 });
    expect(block.shape.columnLabelsRow).toBeUndefined();
    expect(block.shape.headerCols).toBe(2);
  });

  it("subtotals at the top: the outer item's line carries its subtotal, inner items start a line below", () => {
    const { block } = cityByYear("outline", "top");
    expect(block.values.slice(2)).toEqual([
      ["Tehran", "", 100, 101, 102],
      ["", "Ride", 200, 201, 202],
      ["", "Bike", 300, 301, 302],
      ["Shiraz", "", 400, 401, 402],
      ["", "Ride", 500, 501, 502],
      ["", "Bike", 600, 601, 602],
      ["Grand Total", "", 700, 701, 702],
    ]);
    expect([2, 3, 5, 8].map((row) => isBold(block, row, 3))).toEqual([true, false, true, true]);
  });

  it("subtotals at the bottom and off", () => {
    const bottom = cityByYear("outline", "bottom").block;
    expect(bottom.values.slice(2).map((line) => line.slice(0, 3))).toEqual([
      ["Tehran", "", ""],
      ["", "Ride", 100],
      ["", "Bike", 200],
      ["Tehran Total", "", 300],
      ["Shiraz", "", ""],
      ["", "Ride", 400],
      ["", "Bike", 500],
      ["Shiraz Total", "", 600],
      ["Grand Total", "", 700],
    ]);
    const none = cityByYear("outline", "none").block;
    expect(none.values.slice(2).map((line) => line.slice(0, 3))).toEqual([
      ["Tehran", "", ""],
      ["", "Ride", 100],
      ["", "Bike", 200],
      ["Shiraz", "", ""],
      ["", "Ride", 300],
      ["", "Bike", 400],
      ["Grand Total", "", 500],
    ]);
  });

  it("keeps the outer member behind the blank cells of its inner lines", () => {
    const { block } = cityByYear("outline", "top");
    expect(memberAt(block, 2, 0)).toEqual({ member: tehran, axis: "rows" });
    expect(memberAt(block, 3, 0)).toEqual({ member: tehran, axis: "rows" });
    expect(memberAt(block, 3, 1)).toEqual({ member: ride, axis: "rows" });
    expect(memberAt(block, 2, 1)).toBeNull(); // Tehran's own line, Service Type column
    expect(valueCellMembers(block, 0, 0)).toEqual([tehran.uniqueName, y2025.uniqueName]);
    expect(valueCellMembers(block, 2, 1)).toEqual([tehran.uniqueName, bike.uniqueName, y2026.uniqueName]);
  });

  it("repeats the outer labels on every line of the group when asked", () => {
    const { block } = cityByYear("outline", "bottom", { repeatLabels: true });
    expect(column(block, 0)).toEqual([
      "Tehran",
      "Tehran",
      "Tehran",
      "Tehran Total",
      "Shiraz",
      "Shiraz",
      "Shiraz",
      "Shiraz Total",
      "Grand Total",
    ]);
    expect(column(block, 1)).toEqual(["", "Ride", "Bike", "", "", "Ride", "Bike", "", ""]);
  });

  it("indents a drilled-down item in its own column only", () => {
    const cal2025 = ref(CAL, "2025", 1, { hasChildren: true, expanded: true });
    const q1 = ref(CAL, "Q1", 2);
    const tuples: Tuple[] = [[cal2025, tehran], [q1, tehran], [null, null]];
    const layout = {
      filters: [],
      rows: [placed(CAL, "Calendar", { expanded: [cal2025.uniqueName] }), placed(CITY, "City")],
      columns: [],
      options: options("outline", "none"),
    };
    const block = buildBlock(helperGrid({ rows: tuples, rowCaptions: ["Calendar", "City"] }), blockFields(layout));
    expect(block.values.slice(1).map((line) => line.slice(0, 2))).toEqual([
      ["2025", ""],
      ["", "Tehran"],
      ["Q1", ""],
      ["", "Tehran"],
      ["Grand Total", ""],
    ]);
    expect(indentAt(block, 3, 0)).toBe(1);
    expect(indentAt(block, 4, 1)).toBe(0);
  });
});

describe("tabular form", () => {
  it("is the grid's own lines, one per tuple, subtotals at the bottom whatever is stored", () => {
    for (const stored of ["top", "bottom"] as const) {
      const { grid, block } = cityByYear("tabular", stored);
      expect(block.values.slice(2)).toEqual([
        ["Tehran", "Ride", 100, 101, 102],
        ["", "Bike", 200, 201, 202],
        ["Tehran Total", "", 300, 301, 302],
        ["Shiraz", "Ride", 400, 401, 402],
        ["", "Bike", 500, 501, 502],
        ["Shiraz Total", "", 600, 601, 602],
        ["Grand Total", "", 700, 701, 702],
      ]);
      expect(block.lines.map((line) => line.grid)).toEqual([0, 1, 2, 3, 4, 5, 6]);
      expect(block.grid).toBe(grid);
    }
  });

  it("is what an old report without a form shows", () => {
    const { layout, grid, block } = cityByYear("tabular", "bottom");
    const old = { ...layout, options: { grandTotals: true, subtotals: true, nonEmpty: true } };
    expect(buildBlock(grid, blockFields(old))).toEqual(block);
    expect(buildBlock(grid, blockFields(old)).values).toEqual(buildBlock(grid, { filters: 0, rows: 2, columns: ["Year"] }).values);
  });

  it("gives Show Details a subtotal's own item", () => {
    const { block } = cityByYear("tabular", "bottom");
    expect(valueCellMembers(block, 2, 0)).toEqual([tehran.uniqueName, y2025.uniqueName]); // Tehran Total
    expect(valueCellMembers(block, 1, 0)).toEqual([tehran.uniqueName, bike.uniqueName, y2025.uniqueName]);
  });

  it("repeats outer labels when asked, as before", () => {
    const { block } = cityByYear("tabular", "none", { repeatLabels: true });
    expect(column(block, 0)).toEqual(["Tehran", "Tehran", "Shiraz", "Shiraz", "Grand Total"]);
  });
});

describe("column subtotals when row subtotals are at the top", () => {
  // Year > Service Type on columns, measures Rides and Fare. The helper puts each year's
  // subtotal first; Excel keeps it after the year's items.
  function twoColumnFields() {
    const tuples: Tuple[] = [];
    for (const year of [y2025, y2026]) tuples.push([year, null], [year, ride], [year, bike]);
    tuples.push([null, null]);
    const grid = helperGrid({
      rows: [[tehran], [null]],
      rowCaptions: ["City"],
      columns: tuples,
      measures: ["Rides", "Fare"],
      value: (g, c) => g * 1000 + c,
    });
    const layout = {
      filters: [],
      rows: [placed(CITY, "City")],
      columns: [placed(YEAR, "Year"), placed(SERVICE, "Service Type")],
      values: [
        { measure: "[Measures].[Rides]", caption: "Rides" },
        { measure: "[Measures].[Fare]", caption: "Fare", numberFormat: "0.00" },
      ],
      options: options("compact", "top"),
    };
    return { grid, layout };
  }

  it("moves each one after its group, labels and values together", () => {
    const { grid, layout } = twoColumnFields();
    expect(grid.colHeaderRows[0].slice(0, 3).map((cell) => cell.text)).toEqual(["2025 Total", "", ""]);
    const block = buildBlock(grid, blockFields(layout));
    expect(block.values[1].slice(1)).toEqual(["2025", "", "", "", "2025 Total", "", "2026", "", "", "", "2026 Total", "", "Grand Total", ""]);
    expect(block.values[2].slice(1)).toEqual(["Ride", "", "Bike", "", "", "", "Ride", "", "Bike", "", "", "", "", ""]);
    expect(block.values[3].slice(1)).toEqual(new Array(7).fill(["Rides", "Fare"]).flat());
    // Original columns: 0-1 2025 Total, 2-3 2025/Ride, 4-5 2025/Bike, 6-7 2026 Total, ….
    expect(block.values[4].slice(1)).toEqual([2, 3, 4, 5, 0, 1, 8, 9, 10, 11, 6, 7, 12, 13]);
    expect(block.grid.colKinds).toEqual([
      "data", "data", "data", "data", "subtotal", "subtotal",
      "data", "data", "data", "data", "subtotal", "subtotal",
      "grandTotal", "grandTotal",
    ]);
    expect(block.grid.colValueIndex).toEqual([0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1]);
    // The value field's own number format follows its columns.
    expect(block.numberFormats[4].slice(1, 7)).toEqual(["#,##0", "0.00", "#,##0", "0.00", "#,##0", "0.00"]);
    // The grid it was given is left alone.
    expect(grid.colHeaderRows[0][0].text).toBe("2025 Total");
  });

  it("keeps the members behind the moved headers, for drill and Show Details", () => {
    const { grid, layout } = twoColumnFields();
    const block = buildBlock(grid, blockFields(layout));
    expect(memberAt(block, 1, 1)).toEqual({ member: y2025, axis: "columns" });
    expect(memberAt(block, 1, 2)).toEqual({ member: y2025, axis: "columns" });
    expect(memberAt(block, 2, 3)).toEqual({ member: bike, axis: "columns" });
    expect(memberAt(block, 1, 5)).toBeNull(); // 2025 Total
    expect(block.columnMembers[0]).toEqual([y2025.uniqueName, ride.uniqueName]);
    expect(block.columnMembers[4]).toEqual([y2025.uniqueName]); // 2025 Total
    expect(block.columnMembers[12]).toEqual([]);
    expect(valueCellMembers(block, 0, 5)).toEqual([tehran.uniqueName, y2025.uniqueName]);
  });

  it("leaves the columns as they are when the row subtotals are at the bottom", () => {
    const { layout } = twoColumnFields();
    const tuples: Tuple[] = [];
    for (const year of [y2025, y2026]) tuples.push([year, ride], [year, bike], [year, null]);
    tuples.push([null, null]);
    const grid = helperGrid({ rows: [[tehran], [null]], rowCaptions: ["City"], columns: tuples, measures: ["Rides", "Fare"] });
    const block = buildBlock(grid, blockFields({ ...layout, options: options("compact", "bottom") }));
    expect(block.grid).toBe(grid);
    expect(block.values[1].slice(1, 7)).toEqual(["2025", "", "", "", "2025 Total", ""]);
    expect(block.columnMembers[4]).toEqual([y2025.uniqueName]);
  });
});

describe("finding the filter cells from the layout alone, in every form", () => {
  const cases: { rows: string[]; columns: string[]; filters: string[] }[] = [
    { rows: ["City", "Service Type"], columns: ["Year"], filters: [] },
    { rows: ["City", "Service Type"], columns: [], filters: ["Segment"] },
    { rows: ["City"], columns: ["Year", "Service Type"], filters: ["Segment", "Region"] },
    { rows: [], columns: ["Year"], filters: [] },
    { rows: ["City"], columns: [], filters: [] },
  ];
  const members: Record<string, MemberRef[]> = {
    City: [tehran, shiraz],
    "Service Type": [ride, bike],
    Year: [y2025, y2026],
    Region: [],
    Segment: [],
  };
  const hierarchyOf = (caption: string) => `[D].[${caption}]`;
  /** Every combination of the fields' members, then their grand total. */
  const tuplesOf = (captions: string[]): Tuple[] => {
    let out: Tuple[] = [[]];
    for (const caption of captions) out = out.flatMap((t) => members[caption].map((m) => [...t, m]));
    return captions.length > 0 ? [...out, captions.map(() => null)] : out;
  };

  for (const form of ["compact", "outline", "tabular"] as const) {
    it(`agrees with the block in ${form} form`, () => {
      for (const c of cases) {
        const layout = {
          filters: c.filters.map((f) => placed(hierarchyOf(f), f)),
          rows: c.rows.map((f) => placed(hierarchyOf(f), f)),
          columns: c.columns.map((f) => placed(hierarchyOf(f), f)),
          options: options(form, "top"),
        };
        const grid = helperGrid({
          rows: c.rows.length > 0 ? tuplesOf(c.rows) : [[]],
          rowCaptions: c.rows.length > 0 ? c.rows : [""],
          columns: tuplesOf(c.columns),
          filters: c.filters.map((caption) => ({ caption, value: "All" })),
        });
        if (c.rows.length === 0) grid.rowHeaders = [[{ text: "Total", indent: 0, kind: "grandTotal", member: null }]];
        const block = buildBlock(grid, blockFields(layout));
        for (let row = 0; row <= block.rows; row++) {
          for (let col = 0; col <= block.cols; col++) {
            expect(guessFieldCell(layout, row, col)).toEqual(fieldCellAt(block, row, col));
          }
        }
      }
    });
  }

  it("goes by tabular form for a layout given without options", () => {
    const layout = { filters: [], rows: [placed(CITY, "City")], columns: [placed(YEAR, "Year")] };
    expect(guessFieldCell(layout, 1, 0)).toEqual({ area: "rows", index: 0 });
    expect(guessFieldCell(layout, 0, 0)).toEqual({ area: "columns", index: 0 });
  });
});

describe("blockFields", () => {
  it("passes the form and subtotals at the top on, and leaves them out for tabular and the bottom", () => {
    const base = { filters: [], rows: [], columns: [] };
    const fields = (o: LayoutOptions): BlockFields => blockFields({ ...base, options: o });
    expect(fields(options("compact", "top"))).toMatchObject({ form: "compact", subtotalsAt: "top" });
    expect(fields(options("outline", "bottom")).form).toBe("outline");
    expect("subtotalsAt" in fields(options("outline", "bottom"))).toBe(false);
    expect("subtotalsAt" in fields(options("compact", "none"))).toBe(false);
    // Tabular form: subtotals at the bottom, whatever is stored.
    const tabular = fields(options("tabular", "top"));
    expect("form" in tabular || "subtotalsAt" in tabular).toBe(false);
    expect("form" in blockFields(base)).toBe(false);
  });
});
