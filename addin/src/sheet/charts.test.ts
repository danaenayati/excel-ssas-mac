import { describe, expect, it } from "vitest";
import { createMockApi } from "../dev/mockApi";
import { layoutReducer, newLayout, type LayoutAction, type ReportForm, type SubtotalsChoice } from "../state/layout";
import type { CellValue, Fields, Grid, HeaderCell, Layout, LineKind, ValueField } from "../types";
import {
  BLANK_LABEL,
  MAX_CHART_SERIES,
  MAX_CHARTS,
  TOTAL_LABEL,
  chartAnchor,
  chartCells,
  chartRefusal,
  chartTable,
  dataBlockPlace,
  parseCharts,
  sameCharts,
  serializeCharts,
  withCharts,
  type ChartTable,
} from "./charts";
import { MAX_COLS } from "./geometry";
import { blockFields, buildBlock } from "./layout";

// -- building grids by hand ---------------------------------------------------------

interface MemberOptions {
  level?: number;
  expanded?: boolean;
  name?: string;
}

function member(caption: string, options: MemberOptions = {}): HeaderCell {
  return {
    text: caption,
    indent: 0,
    kind: "member",
    member: {
      uniqueName: options.name ?? `[${caption}]`,
      hierarchy: "[H]",
      caption,
      level: options.level ?? 1,
      hasChildren: options.expanded ?? false,
      expanded: options.expanded ?? false,
    },
  };
}
/** A repeated outer label: sent blank, with its member. */
function repeat(cell: HeaderCell): HeaderCell {
  return { ...cell, text: "", kind: "blank" };
}
function subtotal(text: string): HeaderCell {
  return { text, indent: 0, kind: "subtotal", member: null };
}
function grand(text = "Grand Total"): HeaderCell {
  return { text, indent: 0, kind: "grandTotal", member: null };
}
function measureCell(text: string): HeaderCell {
  return { text, indent: 0, kind: "measure", member: null };
}

interface GridInput {
  rows: HeaderCell[][];
  rowKinds?: LineKind[];
  /** Column field header rows, without the measure row. */
  columns?: HeaderCell[][];
  colKinds?: LineKind[];
  measures: string[];
  values: CellValue[][];
  colValueIndex?: number[];
  cellFormats?: number[][];
  formatPalette?: string[];
}

function grid(input: GridInput): Grid {
  const C = input.measures.length;
  const N = input.rows.length;
  return {
    filterRows: [],
    rowFieldCaptions: [""],
    colHeaderRows: [...(input.columns ?? []), input.measures.map(measureCell)],
    rowHeaders: input.rows,
    values: input.values,
    formatPalette: input.formatPalette ?? ["General"],
    cellFormats: input.cellFormats ?? input.values.map((line) => line.map(() => 0)),
    rowKinds: input.rowKinds ?? new Array<LineKind>(N).fill("data"),
    colKinds: input.colKinds ?? new Array<LineKind>(C).fill("data"),
    colValueIndex: input.colValueIndex ?? new Array<number>(C).fill(0),
    stats: { rows: N, columns: C, elapsedMs: 0, mdx: "" },
  };
}

function values(...captions: string[]): Pick<Layout, "values"> {
  return { values: captions.map((caption): ValueField => ({ measure: `[Measures].[${caption}]`, caption })) };
}

// -- the table ----------------------------------------------------------------------

describe("chartTable: rows", () => {
  it("makes one category per item and leaves the Grand Total line out", () => {
    const table = chartTable(
      grid({
        rows: [[member("Tehran")], [member("Shiraz")], [grand()]],
        rowKinds: ["data", "data", "grandTotal"],
        measures: ["Rides"],
        values: [[10], [20], [30]],
      }),
      values("Rides"),
    );
    expect(table.categories).toEqual(["Tehran", "Shiraz"]);
    expect(table.series).toEqual([TOTAL_LABEL]);
    expect(table.values).toEqual([[10], [20]]);
    expect(table.title).toBe("Rides");
  });

  it("names the one category Total when there are no row fields", () => {
    const table = chartTable(
      grid({ rows: [[grand("Total")]], rowKinds: ["data"], measures: ["Rides"], values: [[42]] }),
      values("Rides"),
    );
    expect(table.categories).toEqual(["Total"]);
    expect(table.values).toEqual([[42]]);
  });

  it("joins the row labels outer to inner, filling in repeated outer labels", () => {
    const iran = member("Iran");
    const turkey = member("Turkey");
    const table = chartTable(
      grid({
        rows: [
          [iran, member("Tehran")],
          [repeat(iran), member("Shiraz")],
          [turkey, member("Istanbul")],
        ],
        measures: ["Rides"],
        values: [[1], [2], [3]],
      }),
      values("Rides"),
    );
    expect(table.categories).toEqual(["Iran / Tehran", "Iran / Shiraz", "Turkey / Istanbul"]);
  });

  it("leaves subtotal lines out, with subtotals at the bottom", () => {
    const iran = member("Iran");
    const turkey = member("Turkey");
    const table = chartTable(
      grid({
        rows: [
          [iran, member("Tehran")],
          [repeat(iran), member("Shiraz")],
          [subtotal("Iran Total"), subtotal("")],
          [turkey, member("Istanbul")],
          [subtotal("Turkey Total"), subtotal("")],
          [grand(), grand("")],
        ],
        rowKinds: ["data", "data", "subtotal", "data", "subtotal", "grandTotal"],
        measures: ["Rides"],
        values: [[1], [2], [3], [4], [4], [7]],
      }),
      values("Rides"),
    );
    expect(table.categories).toEqual(["Iran / Tehran", "Iran / Shiraz", "Turkey / Istanbul"]);
    expect(table.values).toEqual([[1], [2], [4]]);
  });

  it("leaves subtotal lines out, with subtotals at the top", () => {
    const iran = member("Iran");
    const table = chartTable(
      grid({
        rows: [
          [subtotal("Iran"), subtotal("")],
          [repeat(iran), member("Tehran")],
          [repeat(iran), member("Shiraz")],
          [grand(), grand("")],
        ],
        rowKinds: ["subtotal", "data", "data", "grandTotal"],
        measures: ["Rides"],
        values: [[3], [1], [2], [3]],
      }),
      values("Rides"),
    );
    expect(table.categories).toEqual(["Iran / Tehran", "Iran / Shiraz"]);
    expect(table.values).toEqual([[1], [2]]);
  });

  it("calls an item with an empty caption (blank), as Windows does", () => {
    const table = chartTable(
      grid({ rows: [[member("")], [member("Shiraz")]], measures: ["Rides"], values: [[1], [2]] }),
      values("Rides"),
    );
    expect(table.categories).toEqual([BLANK_LABEL, "Shiraz"]);
  });
});

describe("chartTable: expanded members", () => {
  it("plots the children of an expanded member, labelled with it, and leaves its own line out", () => {
    const table = chartTable(
      grid({
        rows: [
          [member("2025", { level: 1, expanded: true })],
          [member("Q1 2025", { level: 2 })],
          [member("Q2 2025", { level: 2 })],
          [member("2026", { level: 1 })],
          [grand()],
        ],
        rowKinds: ["data", "data", "data", "data", "grandTotal"],
        measures: ["Rides"],
        values: [[30], [10], [20], [40], [70]],
      }),
      values("Rides"),
    );
    expect(table.categories).toEqual(["2025 / Q1 2025", "2025 / Q2 2025", "2026"]);
    expect(table.values).toEqual([[10], [20], [40]]);
  });

  it("goes down several levels", () => {
    const table = chartTable(
      grid({
        rows: [
          [member("2025", { level: 1, expanded: true })],
          [member("Q1 2025", { level: 2, expanded: true })],
          [member("January 2025", { level: 3 })],
          [member("February 2025", { level: 3 })],
          [member("Q2 2025", { level: 2 })],
        ],
        measures: ["Rides"],
        values: [[9], [3], [1], [2], [6]],
      }),
      values("Rides"),
    );
    expect(table.categories).toEqual([
      "2025 / Q1 2025 / January 2025",
      "2025 / Q1 2025 / February 2025",
      "2025 / Q2 2025",
    ]);
  });

  it("keeps an expanded member whose children are all hidden", () => {
    const table = chartTable(
      grid({
        rows: [[member("2025", { level: 1, expanded: true })], [member("2026", { level: 1 })]],
        measures: ["Rides"],
        values: [[1], [2]],
      }),
      values("Rides"),
    );
    expect(table.categories).toEqual(["2025", "2026"]);
  });

  it("drops an expanded outer item's own lines when an inner field follows", () => {
    const iran = member("Iran", { level: 1, expanded: true });
    const tehran = member("Tehran", { level: 2 });
    const table = chartTable(
      grid({
        rows: [
          [iran, member("Eco")],
          [repeat(iran), member("Plus")],
          [tehran, member("Eco")],
          [repeat(tehran), member("Plus")],
        ],
        measures: ["Rides"],
        values: [[5], [6], [1], [2]],
      }),
      values("Rides"),
    );
    expect(table.categories).toEqual(["Iran / Tehran / Eco", "Iran / Tehran / Plus"]);
    expect(table.values).toEqual([[1], [2]]);
  });

  it("tells an inner field's parents apart under each outer item", () => {
    const eco = member("Eco");
    const plus = member("Plus");
    const iran = (expanded: boolean) => member("Iran", { level: 1, expanded });
    const table = chartTable(
      grid({
        rows: [
          [eco, iran(true)],
          [repeat(eco), member("Tehran", { level: 2 })],
          // Under Plus nothing of Iran's is shown: Iran's own line stays.
          [plus, iran(true)],
          [repeat(plus), member("Turkey", { level: 1 })],
        ],
        measures: ["Rides"],
        values: [[3], [3], [4], [5]],
      }),
      values("Rides"),
    );
    expect(table.categories).toEqual(["Eco / Iran / Tehran", "Plus / Iran", "Plus / Turkey"]);
  });

  it("does the same for expanded members on columns", () => {
    const table = chartTable(
      grid({
        rows: [[member("Tehran")]],
        columns: [
          [
            member("2025", { level: 1, expanded: true }),
            member("Q1 2025", { level: 2 }),
            member("Q2 2025", { level: 2 }),
            member("2026", { level: 1 }),
            grand(),
          ],
        ],
        colKinds: ["data", "data", "data", "data", "grandTotal"],
        measures: ["Rides", "Rides", "Rides", "Rides", "Rides"],
        values: [[30, 10, 20, 40, 70]],
      }),
      values("Rides"),
    );
    expect(table.series).toEqual(["2025 / Q1 2025", "2025 / Q2 2025", "2026"]);
    expect(table.values).toEqual([[10, 20, 40]]);
  });
});

describe("chartTable: columns and value fields", () => {
  it("makes one series per column item and leaves the Grand Total column out", () => {
    const table = chartTable(
      grid({
        rows: [[member("Tehran")], [member("Shiraz")]],
        columns: [[member("2025"), member("2026"), grand()]],
        colKinds: ["data", "data", "grandTotal"],
        measures: ["Rides", "Rides", "Rides"],
        values: [
          [1, 2, 3],
          [4, 5, 9],
        ],
      }),
      values("Rides"),
    );
    expect(table.series).toEqual(["2025", "2026"]);
    expect(table.values).toEqual([
      [1, 2],
      [4, 5],
    ]);
    expect(table.title).toBe("Rides");
  });

  it("names each value field's series by its caption when there are several", () => {
    const table = chartTable(
      grid({
        rows: [[member("Tehran")]],
        measures: ["Rides", "Fare"],
        values: [[1, 100]],
        colValueIndex: [0, 1],
      }),
      values("Rides", "Fare"),
    );
    expect(table.series).toEqual(["Rides", "Fare"]);
    expect(table.title).toBe("Rides, Fare");
  });

  it("joins the column labels and the value field's caption", () => {
    const y2025 = member("2025");
    const y2026 = member("2026");
    const table = chartTable(
      grid({
        rows: [[member("Tehran")]],
        columns: [[y2025, repeat(y2025), y2026, repeat(y2026), grand(), grand("")]],
        colKinds: ["data", "data", "data", "data", "grandTotal", "grandTotal"],
        measures: ["Rides", "Fare", "Rides", "Fare", "Rides", "Fare"],
        values: [[1, 10, 2, 20, 3, 30]],
        colValueIndex: [0, 1, 0, 1, 0, 1],
      }),
      values("Rides", "Fare"),
    );
    expect(table.series).toEqual(["2025 / Rides", "2025 / Fare", "2026 / Rides", "2026 / Fare"]);
    expect(table.values).toEqual([[1, 10, 2, 20]]);
  });

  it("joins two column fields and leaves their subtotal columns out", () => {
    const y2025 = member("2025");
    const table = chartTable(
      grid({
        rows: [[member("Tehran")]],
        columns: [
          [y2025, repeat(y2025), subtotal("2025 Total"), grand()],
          [member("Eco"), member("Plus"), subtotal(""), grand("")],
        ],
        colKinds: ["data", "data", "subtotal", "grandTotal"],
        measures: ["Rides", "Rides", "Rides", "Rides"],
        values: [[1, 2, 3, 3]],
      }),
      values("Rides"),
    );
    expect(table.series).toEqual(["2025 / Eco", "2025 / Plus"]);
    expect(table.values).toEqual([[1, 2]]);
  });

  it("titles the chart as the report heads its values (Show Values As included)", () => {
    const table = chartTable(
      grid({
        rows: [[member("Tehran")]],
        measures: ["Rides (% of Grand Total)", "Fare"],
        values: [[0.4, 100]],
        colValueIndex: [0, 1],
      }),
      values("Rides", "Fare"),
    );
    expect(table.title).toBe("Rides (% of Grand Total), Fare");
    expect(table.series).toEqual(["Rides (% of Grand Total)", "Fare"]);
  });

  it("uses a value field's Custom Name for its series and the title", () => {
    const layout: Pick<Layout, "values"> = {
      values: [
        { measure: "[Measures].[Rides]", caption: "Rides", showAs: "percentOfGrandTotal", customName: "Share of rides" },
        { measure: "[Measures].[Fare]", caption: "Fare", customName: "  " },
      ],
    };
    const y2025 = member("2025");
    const table = chartTable(
      grid({
        rows: [[member("Tehran")]],
        columns: [[y2025, repeat(y2025)]],
        measures: ["Rides (% of Grand Total)", "Fare"],
        values: [[0.4, 100]],
        colValueIndex: [0, 1],
      }),
      layout,
    );
    expect(table.series).toEqual(["2025 / Share of rides", "2025 / Fare"]);
    expect(table.title).toBe("Share of rides, Fare");
    // With nothing to plot, the title still uses it.
    expect(chartTable(grid({ rows: [], columns: [[]], measures: [], values: [] }), layout).title).toBe("Share of rides, Fare");
  });

  it("works with the cube's default measure (no value fields)", () => {
    const table = chartTable(
      grid({ rows: [[member("Tehran")]], measures: ["Rides"], values: [[5]], colValueIndex: [-1] }),
      { values: [] },
    );
    expect(table.series).toEqual([TOTAL_LABEL]);
    expect(table.title).toBe("Rides");
  });

  it("falls back to the value fields' captions, then Total, for the title", () => {
    const onlyTotals = grid({
      rows: [[member("Tehran")]],
      columns: [[grand()]],
      colKinds: ["grandTotal"],
      measures: [""],
      values: [[5]],
    });
    expect(chartTable(onlyTotals, values("Rides")).title).toBe("Rides");
    expect(chartTable(onlyTotals, { values: [] }).title).toBe(TOTAL_LABEL);
    expect(chartTable(onlyTotals, values("Rides")).series).toEqual([]);
  });
});

describe("chartTable: values", () => {
  it("turns anything that is not a number into a gap", () => {
    const table = chartTable(
      grid({
        rows: [[member("A")], [member("B")]],
        columns: [[member("x"), member("y"), member("z")]],
        measures: ["Rides", "Rides", "Rides"],
        values: [
          [null, "#N/A", true],
          [1.5, 0, -2],
        ],
      }),
      values("Rides"),
    );
    expect(table.values).toEqual([
      [null, null, null],
      [1.5, 0, -2],
    ]);
  });

  it("copes with a line shorter than the columns", () => {
    const table = chartTable(
      grid({ rows: [[member("A")]], columns: [[member("x"), member("y")]], measures: ["R", "R"], values: [[1]] }),
      values("R"),
    );
    expect(table.values).toEqual([[1, null]]);
  });

  it("is empty when the report has no data", () => {
    const table = chartTable(grid({ rows: [], measures: ["Rides"], values: [] }), values("Rides"));
    expect(table.categories).toEqual([]);
    expect(table.values).toEqual([]);
    expect(table.series).toEqual([TOTAL_LABEL]);
  });

  it("keeps the report's number formats, the user's own first", () => {
    const layout: Pick<Layout, "values"> = {
      values: [
        { measure: "[Measures].[Rides]", caption: "Rides" },
        { measure: "[Measures].[Fare]", caption: "Fare", numberFormat: "#,##0.00 €" },
      ],
    };
    const table = chartTable(
      grid({
        rows: [[member("A")]],
        measures: ["Rides", "Fare"],
        values: [[1, 2]],
        colValueIndex: [0, 1],
        formatPalette: ["General", "#,##0"],
        cellFormats: [[1, 1]],
      }),
      layout,
    );
    expect(table.numberFormats).toEqual([["#,##0", "#,##0.00 €"]]);
  });
});

// -- the same table whatever the report's form ----------------------------------------

describe("chartTable on real results, in every report form", () => {
  async function fieldsOf(): Promise<Fields> {
    return createMockApi({ delayMs: 0 }).fields("demo", "Sample", "Rides");
  }
  function hier(fields: Fields, caption: string) {
    for (const d of fields.dimensions) {
      const found = d.hierarchies.find((h) => h.caption === caption);
      if (found) return found;
    }
    throw new Error("no hierarchy " + caption);
  }
  function measure(fields: Fields, caption: string) {
    const found = fields.measures.find((m) => m.caption === caption);
    if (!found) throw new Error("no measure " + caption);
    return found;
  }
  async function tableFor(form: ReportForm, subtotals: SubtotalsChoice, extra: (f: Fields) => LayoutAction[] = () => []) {
    const api = createMockApi({ delayMs: 0 });
    const fields = await fieldsOf();
    const actions: LayoutAction[] = [
      { type: "addMeasure", measure: measure(fields, "Rides") },
      { type: "addHierarchy", hierarchy: hier(fields, "City") },
      { type: "addHierarchy", hierarchy: hier(fields, "Service Type") },
      { type: "addHierarchy", hierarchy: hier(fields, "Year"), area: "columns" },
      { type: "setForm", form },
      { type: "setSubtotals", at: subtotals },
      ...extra(fields),
    ];
    const layout = actions.reduce(layoutReducer, newLayout());
    const result = await api.query("demo", { catalog: "Sample", cube: "Rides", layout });
    const block = buildBlock(result, blockFields(layout));
    return { table: chartTable(block.grid, layout), result };
  }

  it("is the same table in compact, outline and tabular form, with subtotals anywhere", async () => {
    const reference = (await tableFor("tabular", "none")).table;
    expect(reference.categories.length).toBeGreaterThan(4);
    expect(reference.categories[0]).toMatch(/^[^/]+ \/ [^/]+$/); // "City / Service"
    expect(reference.series).toEqual(["2025", "2026"]);
    for (const form of ["compact", "outline", "tabular"] as ReportForm[]) {
      for (const subtotals of ["none", "top", "bottom"] as SubtotalsChoice[]) {
        const { table } = await tableFor(form, subtotals);
        expect(table, `${form}, subtotals ${subtotals}`).toEqual(reference);
      }
    }
  });

  it("plots exactly the report's data lines and data columns", async () => {
    const { table, result } = await tableFor("compact", "bottom");
    const dataLines = result.rowKinds.filter((k) => k === "data").length;
    const dataColumns = result.colKinds.filter((k) => k === "data").length;
    expect(table.categories).toHaveLength(dataLines);
    expect(table.series).toHaveLength(dataColumns);
    // Every number is one the report shows on a data line, in a data column.
    const shown = new Set<number>();
    result.values.forEach((line, g) =>
      line.forEach((v, c) => {
        if (result.rowKinds[g] === "data" && result.colKinds[c] === "data" && typeof v === "number") shown.add(v);
      }),
    );
    for (const line of table.values) for (const v of line) if (v !== null) expect(shown.has(v)).toBe(true);
  });

  it("follows a drilled-down hierarchy as a PivotChart does", async () => {
    const api = createMockApi({ delayMs: 0 });
    const fields = await fieldsOf();
    const calendar = hier(fields, "Calendar");
    const top = await api.members("demo", { catalog: "Sample", cube: "Rides", hierarchy: calendar.uniqueName });
    const first = top.members[0];
    const layout = [
      { type: "addMeasure", measure: measure(fields, "Rides") },
      { type: "addHierarchy", hierarchy: calendar },
      { type: "expand", hierarchy: calendar.uniqueName, member: first.uniqueName },
    ].reduce((l, a) => layoutReducer(l, a as LayoutAction), newLayout());
    const result = await api.query("demo", { catalog: "Sample", cube: "Rides", layout });
    const table = chartTable(buildBlock(result, blockFields(layout)).grid, layout);
    expect(table.categories).not.toContain(first.caption);
    expect(table.categories.filter((c) => c.indexOf(first.caption + " / ") === 0)).toHaveLength(4);
    expect(table.categories).toContain(top.members[1].caption);
  });
});

// -- cells, refusal, placement ----------------------------------------------------------

describe("chartCells", () => {
  const table: ChartTable = {
    title: "Rides",
    series: ["2025", "=odd"],
    categories: ["Tehran", "2026"],
    values: [
      [1, null],
      [3, 4],
    ],
    numberFormats: [
      ["#,##0", "#,##0"],
      ["#,##0", ""],
    ],
  };

  it("lays the table out with an empty corner, names as text", () => {
    const cells = chartCells(table);
    expect(cells.rows).toBe(3);
    expect(cells.cols).toBe(3);
    expect(cells.values).toEqual([
      ["", "2025", "'=odd"],
      ["Tehran", 1, ""],
      ["2026", 3, 4],
    ]);
    expect(cells.numberFormats).toEqual([
      ["General", "@", "@"],
      ["@", "#,##0", "#,##0"],
      ["@", "#,##0", "General"],
    ]);
  });

  it("gives an empty table one empty category and series", () => {
    const cells = chartCells({ title: "Total", series: [], categories: [], values: [], numberFormats: [] });
    expect(cells.rows).toBe(2);
    expect(cells.cols).toBe(2);
    expect(cells.values).toEqual([
      ["", ""],
      ["", ""],
    ]);
  });
});

describe("chartRefusal", () => {
  it("refuses more series than Excel can draw", () => {
    const series = (n: number): ChartTable => ({
      title: "x",
      series: new Array<string>(n).fill("s"),
      categories: [],
      values: [],
      numberFormats: [],
    });
    expect(chartRefusal(series(MAX_CHART_SERIES))).toBeNull();
    expect(chartRefusal(series(MAX_CHART_SERIES + 1))).toMatch(/at most 255 series/);
  });
});

describe("placement", () => {
  it("puts a chart beside the report, one empty column away, and steps further charts aside", () => {
    const report = { row: 2, col: 1, rows: 10, cols: 3 };
    expect(chartAnchor(report, 0)).toEqual({ row: 2, col: 5 });
    expect(chartAnchor(report, 1)).toEqual({ row: 4, col: 6 });
    expect(chartAnchor({ row: 0, col: MAX_COLS - 2, rows: 1, cols: 2 }, 0).col).toBe(MAX_COLS - 1);
  });

  it("keeps a data block where it was unless it has grown into other cells", () => {
    const previous = { row: 10, col: 0, rows: 3, cols: 2 };
    expect(dataBlockPlace(previous, { rows: 5, cols: 4 }, false, 40)).toEqual({ row: 10, col: 0, rows: 5, cols: 4 });
    expect(dataBlockPlace(previous, { rows: 5, cols: 4 }, true, 40)).toEqual({ row: 41, col: 0, rows: 5, cols: 4 });
    expect(dataBlockPlace(null, { rows: 2, cols: 2 }, false, 0)).toEqual({ row: 0, col: 0, rows: 2, cols: 2 });
    expect(dataBlockPlace(null, { rows: 2, cols: 2 }, false, 7)).toEqual({ row: 8, col: 0, rows: 2, cols: 2 });
  });
});

// -- the stored record ----------------------------------------------------------------

describe("parseCharts", () => {
  it("reads back what was stored", () => {
    const charts = [
      { id: "abc123", type: "Pie", sheetId: "{SHEET-1}", title: "Rides" },
      { id: "def456", type: "ColumnClustered", sheetId: "{SHEET-2}" },
    ];
    expect(parseCharts(JSON.parse(JSON.stringify(serializeCharts(charts as never))))).toEqual(charts);
  });

  it("drops what cannot be read, and repeats", () => {
    expect(parseCharts(undefined)).toEqual([]);
    expect(parseCharts("charts")).toEqual([]);
    expect(
      parseCharts([
        null,
        "x",
        [1],
        { id: "bad id", type: "Pie", sheetId: "s" },
        { id: "ok1", type: "Doughnut", sheetId: "s" },
        { id: "ok2", type: "Line" },
        { id: "ok3", type: "Line", sheetId: "" },
        { id: "ok4", type: "Area", sheetId: "s", title: 5 },
        { id: "ok4", type: "Pie", sheetId: "s" },
        { id: "ok5", type: "BarClustered", sheetId: "s", title: "x".repeat(300), extra: true },
      ]),
    ).toEqual([
      { id: "ok4", type: "Area", sheetId: "s" },
      { id: "ok5", type: "BarClustered", sheetId: "s" },
    ]);
  });

  it("reads at most MAX_CHARTS", () => {
    const many = Array.from({ length: MAX_CHARTS + 5 }, (_, i) => ({ id: "c" + i, type: "Line", sheetId: "s" }));
    expect(parseCharts(many)).toHaveLength(MAX_CHARTS);
  });

  it("stores only the known fields", () => {
    const sneaky = { id: "a", type: "Pie", sheetId: "s", extra: 1 } as never;
    expect(serializeCharts([sneaky])).toEqual([{ id: "a", type: "Pie", sheetId: "s" }]);
  });

  it("tells equal lists apart from changed ones", () => {
    const a = [{ id: "a", type: "Pie" as const, sheetId: "s", title: "T" }];
    expect(sameCharts(a, [{ ...a[0] }])).toBe(true);
    expect(sameCharts(a, [{ ...a[0], title: "U" }])).toBe(false);
    expect(sameCharts(undefined, [])).toBe(true);
    expect(sameCharts(a, [])).toBe(false);
  });

  it("puts charts on a definition, or takes the field away when there are none", () => {
    const def = { id: "r1", connectionId: "c", catalog: "k", cube: "q", layout: newLayout() };
    const chart = { id: "a", type: "Pie" as const, sheetId: "s" };
    const withOne = withCharts(def, [chart]);
    expect(withOne.charts).toEqual([chart]);
    expect(def).not.toHaveProperty("charts");
    expect(withCharts(withOne, [])).not.toHaveProperty("charts");
  });
});
