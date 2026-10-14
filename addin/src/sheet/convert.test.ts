import { beforeEach, describe, expect, it } from "vitest";
import { createMockApi } from "../dev/mockApi";
import { emptyLayout, layoutReducer, type LayoutAction } from "../state/layout";
import { newReportDef } from "../state/report";
import { ReportRunner, type ReportRuntime } from "../state/runner";
import type { Api } from "../api";
import type { Fields, Grid, HeaderCell, Layout, MemberRef } from "../types";
import {
  conversionBlocker,
  convertReport,
  excelString,
  formulaConnection,
  type Conversion,
  type ConversionSource,
} from "./convert";
import { FILTER_FORMAT, TEXT_FORMAT, blockFields, buildBlock } from "./layout";
import { PreviewHost } from "./preview";

const CONN = "Sample cube/Sample/Rides";
const C = '"Sample cube/Sample/Rides"';

let api: Api;
let host: PreviewHost;
let fields: Fields;
let runner: ReportRunner;

function hierarchy(caption: string) {
  for (const d of fields.dimensions) {
    const found = d.hierarchies.find((h) => h.caption === caption);
    if (found) return found;
  }
  throw new Error("no hierarchy " + caption);
}
function measure(caption: string) {
  const found = fields.measures.find((m) => m.caption === caption);
  if (!found) throw new Error("no measure " + caption);
  return found;
}

/** Run a report against the mock cube into the preview, and return what the pane remembers. */
async function run(actions: LayoutAction[]): Promise<ReportRuntime> {
  const layout = actions.reduce(layoutReducer, emptyLayout());
  const def = newReportDef("demo", "Sample", "Rides");
  const outcome = await runner.run({ ...def, layout }, {
    anchor: { sheetId: "preview-sheet-1", sheetName: "Sheet1", row: 0, col: 0 },
    useSelection: false,
  });
  if (outcome.kind !== "written") throw new Error("not written: " + outcome.kind);
  return outcome.runtime;
}

function converted(source: ConversionSource): Conversion {
  const result = convertReport(source, CONN);
  if ("refused" in result) throw new Error("refused: " + result.refused);
  return result;
}

function member(name: string): string {
  return `=CUBE.MEMBER(${C},"${name}")`;
}

function value(...names: string[]): string {
  return `=CUBE.VALUE(${C}${names.map((n) => `,"${n}"`).join("")})`;
}

beforeEach(async () => {
  api = createMockApi({ delayMs: 0 });
  host = new PreviewHost({ persist: false });
  fields = await api.fields("demo", "Sample", "Rides");
  runner = new ReportRunner({ api, host, confirmOverwrite: async () => true });
});

describe("convertReport", () => {
  it("turns members into CUBE.MEMBER and values into CUBE.VALUE, totals at All", async () => {
    const runtime = await run([
      { type: "addMeasure", measure: measure("Rides") },
      { type: "addHierarchy", hierarchy: hierarchy("Year") },
    ]);
    const { formulas, numberFormats, filtered } = converted(runtime);
    // Year | Rides / 2025 / 2026 / Grand Total
    expect(formulas).toEqual([
      [null, null],
      [member("[Date].[Year].&[2025]"), value("[Date].[Year].&[2025]", "[Measures].[Rides]")],
      [member("[Date].[Year].&[2026]"), value("[Date].[Year].&[2026]", "[Measures].[Rides]")],
      [null, value("[Date].[Year].[All]", "[Measures].[Rides]")],
    ]);
    // Header cells were text and become General; the caption loses its filter arrow;
    // value cells keep their own format.
    expect(runtime.sheet!.numberFormats[0][0]).toBe(FILTER_FORMAT);
    expect(numberFormats).toEqual([
      [TEXT_FORMAT, null],
      ["General", null],
      ["General", null],
      [null, null],
    ]);
    expect(filtered).toBe(false);
  });

  it("puts row members, column members, the measure and page filters in each value", async () => {
    const service = hierarchy("Service Type");
    const runtime = await run([
      { type: "addMeasure", measure: measure("Rides") },
      { type: "addMeasure", measure: measure("Fare") },
      { type: "addHierarchy", hierarchy: hierarchy("City") },
      { type: "addHierarchy", hierarchy: hierarchy("Year"), area: "columns" },
      { type: "addHierarchy", hierarchy: service, area: "filters" },
      { type: "addHierarchy", hierarchy: hierarchy("Month Name"), area: "filters" },
      { type: "setMembers", hierarchy: service.uniqueName, members: ["[Service].[Service Type].&[2]"], captions: ["x"] },
    ]);
    const { formulas } = converted(runtime);
    const shape = runtime.shape;
    // Page filters: their item, or All.
    expect(formulas[0][1]).toBe(member("[Service].[Service Type].&[2]"));
    expect(formulas[1][1]).toBe(member("[Date].[Month Name].[All]"));
    const city = runtime.headers.rowHeaders[0][0].member!.uniqueName;
    const slicers = ["[Service].[Service Type].&[2]", "[Date].[Month Name].[All]"];
    expect(formulas[shape.dataTop][shape.headerCols]).toBe(value(city, "[Date].[Year].&[2025]", "[Measures].[Rides]", ...slicers));
    expect(formulas[shape.dataTop][shape.headerCols + 1]).toBe(value(city, "[Date].[Year].&[2025]", "[Measures].[Fare]", ...slicers));
    // The Grand Total column, on the last line (the Grand Total row).
    const last = shape.dataTop + shape.dataRows - 1;
    const lastCol = shape.headerCols + shape.dataCols - 1;
    expect(formulas[last][lastCol]).toBe(
      value("[Geography].[City].[All]", "[Date].[Year].[All]", "[Measures].[Fare]", ...slicers),
    );
    // Column members are written once, over their first measure column; the repeat stays blank.
    expect(formulas[shape.headerTop][shape.headerCols]).toBe(member("[Date].[Year].&[2025]"));
    expect(formulas[shape.headerTop][shape.headerCols + 1]).toBeNull();
  });

  it("reads a subtotal's own member from its group, on rows and on columns", async () => {
    const runtime = await run([
      { type: "addMeasure", measure: measure("Rides") },
      { type: "addMeasure", measure: measure("Fare") },
      { type: "addHierarchy", hierarchy: hierarchy("Year") },
      { type: "addHierarchy", hierarchy: hierarchy("Service Type") },
      { type: "addHierarchy", hierarchy: hierarchy("Month Name"), area: "columns" },
      { type: "addHierarchy", hierarchy: hierarchy("City"), area: "columns" },
      { type: "setOption", option: "subtotals", value: true },
    ]);
    const { formulas } = converted(runtime);
    const { rowHeaders, colHeaderRows } = runtime.headers;
    const shape = runtime.shape;
    const subtotalRow = rowHeaders.findIndex((line) => line[0].kind === "subtotal");
    const subtotalCol = colHeaderRows[0].findIndex((cell) => cell.kind === "subtotal");
    expect(subtotalRow).toBeGreaterThan(0);
    expect(subtotalCol).toBeGreaterThan(0);
    expect(rowHeaders[subtotalRow][0].member).toBeNull(); // "2025 Total" names no member
    // Its second measure column: the same subtotal, repeated.
    expect(colHeaderRows[0][subtotalCol + 1].kind).toBe("subtotal");
    expect(formulas[shape.dataTop + subtotalRow][shape.headerCols + subtotalCol + 1]).toBe(
      value("[Date].[Year].&[2025]", "[Service].[Service Type].[All]", "[Date].[Month Name].&[1]", "[Geography].[City].[All]", "[Measures].[Fare]"),
    );
  });

  it("leaves blank repeats blank, and writes them when labels are repeated", async () => {
    const actions: LayoutAction[] = [
      { type: "addMeasure", measure: measure("Rides") },
      { type: "addHierarchy", hierarchy: hierarchy("Year") },
      { type: "addHierarchy", hierarchy: hierarchy("Service Type") },
    ];
    const plain = await run(actions);
    const top = plain.shape.dataTop;
    expect(plain.headers.rowHeaders[1][0].kind).toBe("blank");
    expect(converted(plain).formulas[top + 1][0]).toBeNull();
    const repeated = await run([...actions, { type: "setOption", option: "repeatLabels", value: true }]);
    expect(converted(repeated).formulas[top + 1][0]).toBe(member("[Date].[Year].&[2025]"));
  });

  it("says when filters hide items, since totals will count them all", async () => {
    const city = hierarchy("City");
    const runtime = await run([
      { type: "addMeasure", measure: measure("Rides") },
      { type: "addHierarchy", hierarchy: city },
      { type: "setLabelFilter", hierarchy: city.uniqueName, filter: { operator: "beginsWith", value: "T" } },
    ]);
    expect(converted(runtime).filtered).toBe(true);
  });

  it("asks for a refresh when it has no sheet to read", async () => {
    const runtime = await run([{ type: "addMeasure", measure: measure("Rides") }]);
    const result = convertReport({ ...runtime, sheet: undefined }, CONN);
    expect(result).toEqual({ refused: expect.stringContaining("Refresh the report first") });
  });
});

// -- subtotals at the top, from a hand-made grid (the sample cube has no such option) --
function ref(uniqueName: string, hierarchy: string): MemberRef {
  return { uniqueName, hierarchy, caption: uniqueName, level: 1, hasChildren: false, expanded: false };
}
function cell(kind: HeaderCell["kind"], text: string, m: MemberRef | null = null): HeaderCell {
  return { text, indent: 0, kind, member: m };
}

describe("subtotals at the top", () => {
  it("read their member from the line after them", () => {
    const Y = "[Y].[Y]";
    const S = "[S].[S]";
    const y1 = ref("[Y].[Y].&[1]", Y);
    const s1 = ref("[S].[S].&[1]", S);
    const s2 = ref("[S].[S].&[2]", S);
    const grid: Grid = {
      filterRows: [],
      rowFieldCaptions: ["Y", "S"],
      colHeaderRows: [[cell("measure", "Rides")]],
      rowHeaders: [
        [cell("subtotal", "1 Total"), cell("subtotal", "")],
        [cell("member", "1", y1), cell("member", "a", s1)],
        [cell("blank", "", y1), cell("member", "b", s2)],
        [cell("grandTotal", "Grand Total"), cell("grandTotal", "")],
      ],
      values: [[3], [1], [2], [3]],
      formatPalette: ["General"],
      cellFormats: [[0], [0], [0], [0]],
      rowKinds: ["subtotal", "data", "data", "grandTotal"],
      colKinds: ["data"],
      colValueIndex: [0],
      stats: { rows: 4, columns: 1, elapsedMs: 0, mdx: "" },
    };
    const layout: Layout = {
      ...emptyLayout(),
      rows: [
        { hierarchy: Y, caption: "Y", allMember: "[Y].[Y].[All]", expanded: [], members: [], memberCaptions: [] },
        { hierarchy: S, caption: "S", allMember: "[S].[S].[All]", expanded: [], members: [], memberCaptions: [] },
      ],
      values: [{ measure: "[Measures].[Rides]", caption: "Rides" }],
      options: { grandTotals: true, subtotals: true, nonEmpty: true, form: "outline", subtotalsAt: "top" },
    };
    const block = buildBlock(grid, blockFields(layout));
    const source: ConversionSource = {
      layout,
      headers: { rowHeaders: grid.rowHeaders, colHeaderRows: grid.colHeaderRows },
      shape: block.shape,
      rows: block.rows,
      cols: block.cols,
      members: block.members,
      colValueIndex: grid.colValueIndex,
      lines: block.lines,
      sheet: block,
    };
    const { formulas } = converted(source);
    // The sheet line that shows grid line g, and its value cell.
    const at = (g: number) => formulas[block.shape.dataTop + block.lines.findIndex((l) => l.grid === g)][block.shape.headerCols];
    expect(at(0)).toBe(value("[Y].[Y].&[1]", "[S].[S].[All]", "[Measures].[Rides]"));
    expect(at(2)).toBe(value("[Y].[Y].&[1]", "[S].[S].&[2]", "[Measures].[Rides]"));
    expect(at(3)).toBe(value("[Y].[Y].[All]", "[S].[S].[All]", "[Measures].[Rides]"));
    // At the bottom instead, the line before it is in no group: refused, not guessed.
    const bottom = { ...source, layout: { ...layout, options: { ...layout.options, subtotalsAt: "bottom" as const } } };
    expect(convertReport(bottom, CONN)).toEqual({ refused: expect.stringContaining("row") });
  });
});

// -- every form, against what Show Details reads for the same cell -----------------
describe("every report form", () => {
  const forms: [string, LayoutAction[]][] = [
    ["tabular", []],
    ["outline, subtotals at the bottom", [{ type: "setForm", form: "outline" }, { type: "setSubtotals", at: "bottom" }]],
    ["outline, subtotals at the top", [{ type: "setForm", form: "outline" }, { type: "setSubtotals", at: "top" }]],
    ["compact, subtotals at the top", [{ type: "setForm", form: "compact" }, { type: "setSubtotals", at: "top" }]],
    ["compact, no subtotals", [{ type: "setForm", form: "compact" }, { type: "setSubtotals", at: "none" }]],
    ["tabular, subtotals", [{ type: "setForm", form: "tabular" }, { type: "setSubtotals", at: "bottom" }]],
  ];
  for (const [name, choice] of forms) {
    it(name, async () => {
      const runtime = await run([
        { type: "addMeasure", measure: measure("Rides") },
        { type: "addMeasure", measure: measure("Fare") },
        { type: "addHierarchy", hierarchy: hierarchy("Year") },
        { type: "addHierarchy", hierarchy: hierarchy("Service Type") },
        { type: "addHierarchy", hierarchy: hierarchy("Month Name"), area: "columns" },
        { type: "addHierarchy", hierarchy: hierarchy("City"), area: "columns" },
        ...choice,
      ]);
      const { formulas } = converted(runtime);
      const { shape, layout } = runtime;
      // Show Details' members leave totals out; a formula names the All member instead.
      const padded = (names: string[], fields: { allMember: string | null }[]) => [
        ...names,
        ...fields.slice(names.length).map((f) => f.allMember as string),
      ];
      let checked = 0;
      for (let n = 0; n < shape.dataRows; n++) {
        for (let c = 0; c < shape.dataCols; c++) {
          const formula = formulas[shape.dataTop + n][shape.headerCols + c];
          const line = runtime.lines[n];
          if (line.grid === null) {
            expect(formula).toBeNull();
            continue;
          }
          const index = runtime.colValueIndex[c];
          expect(formula).toBe(
            value(
              ...padded(line.members, layout.rows),
              ...padded(runtime.columnMembers[c], layout.columns),
              layout.values[index].measure,
            ),
          );
          checked++;
        }
      }
      expect(checked).toBeGreaterThan(0);
      // Every member label became CUBE.MEMBER, every other label stayed.
      runtime.members.forEach((hit, key) => {
        const [row, col] = key.split(",").map(Number);
        const shown = runtime.sheet!.values[row][col];
        expect(formulas[row][col]).toBe(shown === "" ? null : member(hit.member.uniqueName));
      });
    });
  }
});

describe("conversionBlocker", () => {
  const base = (): Layout => ({
    ...emptyLayout(),
    rows: [{ hierarchy: "[G].[C]", caption: "City", allMember: "[G].[C].[All]", expanded: [], members: [], memberCaptions: [] }],
    values: [{ measure: "[Measures].[Rides]", caption: "Rides" }],
  });

  it("allows a plain report", () => {
    expect(conversionBlocker(base())).toBeNull();
  });

  it("refuses what one cube tuple cannot give", () => {
    const cases: [Partial<Layout> | ((l: Layout) => Layout), string][] = [
      [{ values: [{ measure: "[Measures].[Rides]", caption: "Rides %", showAs: "percentOfGrandTotal" }] }, "Show Values As"],
      [{ values: [{ measure: "[Measures].[Goal]", caption: "Goal", kpi: { name: "Growth", part: "goal" } }] }, "KPI"],
      [
        {
          values: [{ measure: "[Measures].[Per Ride]", caption: "Per Ride" }],
          calculatedMeasures: [{ name: "Per Ride", formula: "[Measures].[Fare] / [Measures].[Rides]" }],
        },
        "calculated measure",
      ],
      [(l) => ({ ...l, rows: [{ ...l.rows[0], customItems: [{ name: "North", members: ["[G].[C].&[1]", "[G].[C].&[2]"] }] }] }), "groups"],
      [(l) => ({ ...l, rows: [{ ...l.rows[0], set: "[Top Cities]" }] }), "named set"],
      [
        (l) => ({
          ...l,
          filters: [{ hierarchy: "[D].[Y]", caption: "Year", allMember: null, expanded: [], members: ["[D].[Y].&[1]", "[D].[Y].&[2]"], memberCaptions: [] }],
        }),
        "several items",
      ],
    ];
    for (const [change, needle] of cases) {
      const layout = typeof change === "function" ? change(base()) : { ...base(), ...change };
      expect(conversionBlocker(layout)).toContain(needle);
    }
  });

  it("allows label, value and date filters, and a filter with one item", () => {
    const layout = base();
    layout.rows[0] = {
      ...layout.rows[0],
      labelFilter: { operator: "contains", value: "a" },
      dateFilter: { operator: "thisYear" },
    };
    layout.filters = [{ hierarchy: "[D].[Y]", caption: "Year", allMember: null, expanded: [], members: ["[D].[Y].&[1]"], memberCaptions: [] }];
    expect(conversionBlocker(layout)).toBeNull();
  });
});

describe("formula text", () => {
  it("writes Excel string constants", () => {
    expect(excelString('[A].&[say "hi"]')).toBe('"[A].&[say ""hi""]"');
    expect(excelString("")).toBe('""');
    const long = "x".repeat(600);
    expect(excelString(long)).toBe(['"' + "x".repeat(250) + '"', '"' + "x".repeat(250) + '"', '"' + "x".repeat(100) + '"'].join("&"));
  });

  it("names the connection with its database and cube", () => {
    expect(formulaConnection(" Sales Cube ", "BI", "Model")).toBe("Sales Cube/BI/Model");
    expect(formulaConnection("Sales Cube", "BI/2024", "Model")).toBeNull();
    expect(formulaConnection("Sales Cube", "", "Model")).toBeNull();
  });
});

describe("PreviewHost.writeFormulas", () => {
  it("records formulas and formats, leaves other cells, and checks the size", async () => {
    const runtime = await run([
      { type: "addMeasure", measure: measure("Rides") },
      { type: "addHierarchy", hierarchy: hierarchy("Year") },
    ]);
    const before = host.snapshot().cells.get("1,1")!.value;
    const { formulas, numberFormats } = converted(runtime);
    const id = Array.from(host.snapshot().names.keys())[0];
    await host.writeFormulas(id, formulas, numberFormats);
    const cells = host.snapshot().cells;
    expect(cells.get("1,1")).toMatchObject({ value: before, formula: formulas[1][1] });
    expect(cells.get("1,0")).toMatchObject({ value: "2025", format: "General", formula: formulas[1][0] });
    expect(cells.get("3,0")).toMatchObject({ value: "Grand Total" });
    expect(cells.get("3,0")!.formula).toBeUndefined();
    await expect(host.writeFormulas(id, formulas.slice(1), numberFormats.slice(1))).rejects.toThrow(/Refresh it/);
    await expect(host.writeFormulas("nope", formulas, numberFormats)).rejects.toThrow(/no longer on a sheet/);
  });
});
