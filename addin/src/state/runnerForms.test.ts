// Compact, outline and tabular form end to end: the sample cube, the runner and the
// preview host. Sheet rows no longer map 1:1 to grid lines in compact and outline form;
// everything that turns a cell back into cube data must still find the right one.

import { beforeEach, describe, expect, it } from "vitest";
import type { Api } from "../api";
import { createMockApi } from "../dev/mockApi";
import type { ReportDef } from "../sheet/host";
import { COLUMN_LABELS, FILTER_FORMAT, ROW_LABELS, guessFieldCell } from "../sheet/layout";
import { PreviewHost } from "../sheet/preview";
import type { Fields, Layout } from "../types";
import { layoutReducer, newLayout, type LayoutAction, type ReportForm, type SubtotalsChoice } from "./layout";
import { newReportDef } from "./report";
import { ReportRunner } from "./runner";

const CATALOG = "Sample";
const CUBE = "Rides";
const TEHRAN = "[Geography].[City].&[1]";

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

/** City > Service Type on rows, Year on columns, Rides; in `form` with `subtotals`. */
function cityServiceByYear(form: ReportForm, subtotals: SubtotalsChoice, extra: LayoutAction[] = []): Layout {
  const actions: LayoutAction[] = [
    { type: "addMeasure", measure: measure("Rides") },
    { type: "addHierarchy", hierarchy: hierarchy("City") },
    { type: "addHierarchy", hierarchy: hierarchy("Service Type") },
    { type: "addHierarchy", hierarchy: hierarchy("Year"), area: "columns" },
    { type: "setForm", form },
    { type: "setSubtotals", at: subtotals },
    ...extra,
  ];
  return actions.reduce(layoutReducer, newLayout());
}

async function write(layout: Layout): Promise<{ def: ReportDef; rows: number; cols: number }> {
  host.reset();
  host.select(0, 0);
  const def = { ...newReportDef("demo", CATALOG, CUBE), layout };
  const outcome = await runner.run(def, { anchor: null, useSelection: true });
  if (outcome.kind !== "written") throw new Error("expected written, got " + outcome.kind);
  return { def, rows: outcome.extent.rows, cols: outcome.extent.cols };
}

function text(row: number, col: number) {
  const cell = host.snapshot().cells.get(row + "," + col);
  return cell ? cell.value : undefined;
}

beforeEach(async () => {
  api = createMockApi({ delayMs: 0 });
  host = new PreviewHost({ persist: false });
  fields = await api.fields("demo", CATALOG, CUBE);
  runner = new ReportRunner({ api, host, confirmOverwrite: async () => true });
});

describe("a new report", () => {
  it("is written in Compact Form with subtotals at the top of each group", async () => {
    const layout = [
      { type: "addMeasure", measure: measure("Rides") },
      { type: "addHierarchy", hierarchy: hierarchy("City") },
      { type: "addHierarchy", hierarchy: hierarchy("Service Type") },
    ].reduce((l, a) => layoutReducer(l, a as LayoutAction), newLayout());
    await write(layout);
    expect(text(0, 0)).toBe(ROW_LABELS);
    expect(text(1, 0)).toBe("Tehran");
    expect(text(2, 0)).toBe("Ride");
    // Tehran's line holds Tehran's subtotal: the sum of its services.
    const services = [2, 3, 4].map((row) => Number(text(row, 1)));
    expect(Number(text(1, 1))).toBe(services.reduce((a, b) => a + b, 0));
    expect(text(5, 0)).toBe("Rey");
  });
});

describe("compact form through the runner", () => {
  // Row 0: Column Labels. Row 1: years. Row 2: Row Labels + measures. Row 3: Tehran.
  it("writes Row Labels and Column Labels with the arrow", async () => {
    await write(cityServiceByYear("compact", "top"));
    const cell = (row: number, col: number) => host.snapshot().cells.get(row + "," + col);
    expect(cell(0, 1)).toMatchObject({ value: COLUMN_LABELS, format: FILTER_FORMAT });
    expect(cell(2, 0)).toMatchObject({ value: ROW_LABELS, format: FILTER_FORMAT });
    expect(text(1, 1)).toBe("2025");
    expect(text(3, 0)).toBe("Tehran");
    expect(text(4, 0)).toBe("Ride");
  });

  it("knows the fields behind Row Labels and Column Labels, and that they stand for their whole area", async () => {
    const { def, rows, cols } = await write(cityServiceByYear("compact", "top"));
    const at = (row: number, col: number) => ({ row, col, rows, cols });
    expect(runner.buttonAt(def.id, at(2, 0))?.caption).toBe("City");
    expect(runner.labelsAt(def.id, at(2, 0))).toBe("rows");
    expect(runner.buttonAt(def.id, at(0, 1))?.caption).toBe("Year");
    expect(runner.labelsAt(def.id, at(0, 1))).toBe("columns");
    expect(runner.labelsAt(def.id, at(1, 1))).toBeNull();
    expect(runner.labelsAt(def.id, at(3, 0))).toBeNull();
    expect(runner.labelsAt(def.id, { row: 2, col: 0, rows: rows + 1, cols })).toBeNull(); // block changed size
    // Status bar: the year cells offer Year, the labels cells their area's first field.
    expect(runner.fieldAt(def.id, at(1, 1))?.caption).toBe("Year");
    expect(runner.fieldAt(def.id, at(2, 0))?.caption).toBe("City");
    expect(runner.fieldAt(def.id, at(0, 1))?.caption).toBe("Year");
    expect(runner.fieldAt(def.id, at(0, 2))).toBeNull();
    expect(runner.fieldAt(def.id, at(3, 0))).toBeNull();
  });

  it("finds the member behind outer and inner items, for drill", async () => {
    const { def } = await write(cityServiceByYear("compact", "bottom"));
    expect(runner.memberAt(def.id, 3, 0)?.member.uniqueName).toBe(TEHRAN);
    expect(runner.memberAt(def.id, 4, 0)?.member.caption).toBe("Ride");
    expect(runner.memberAt(def.id, 1, 1)?.member.caption).toBe("2025");
  });

  it("Show Details: an outer item's line is its subtotal at the top, and nothing at the bottom", async () => {
    const top = await write(cityServiceByYear("compact", "top"));
    const at = (row: number, col: number, size: { rows: number; cols: number }) => ({ row, col, ...size });
    const tehran = runner.detailsFor(top.def.id, at(3, 1, top));
    expect(tehran?.members).toEqual([TEHRAN, "[Date].[Year].&[2025]"]);
    const details = await api.details("demo", { catalog: CATALOG, cube: CUBE, ...tehran! });
    const sum = details.rows.reduce((total, row) => total + Number(row[row.length - 1]), 0);
    expect(sum).toBe(Number(text(3, 1)));
    // Tehran / Ride, 2026.
    expect(runner.detailsFor(top.def.id, at(4, 2, top))?.members).toEqual([
      TEHRAN,
      "[Service].[Service Type].&[1]",
      "[Date].[Year].&[2026]",
    ]);
    expect(runner.detailsFor(top.def.id, at(2, 1, top))).toBeNull(); // header
    expect(runner.detailsFor(top.def.id, at(3, 0, top))).toBeNull(); // a row label

    const bottom = await write(cityServiceByYear("compact", "bottom"));
    expect(runner.detailsFor(bottom.def.id, at(3, 1, bottom))).toBeNull(); // Tehran's own line: no values
    expect(text(4, 0)).toBe("Ride");
    const total = [...Array(bottom.rows).keys()].find((row) => text(row, 0) === "Tehran Total") as number;
    expect(runner.detailsFor(bottom.def.id, at(total, 1, bottom))?.members).toEqual([TEHRAN, "[Date].[Year].&[2025]"]);
  });

  it("expands and collapses through outer items, with the grid's own lines behind", async () => {
    const calendar = hierarchy("Calendar").uniqueName;
    const layout = [
      { type: "addMeasure", measure: measure("Rides") },
      { type: "addHierarchy", hierarchy: hierarchy("Calendar") },
      { type: "addHierarchy", hierarchy: hierarchy("City") },
      { type: "setLabelFilter", hierarchy: hierarchy("City").uniqueName, filter: { operator: "equals", value: "Tehran" } },
    ].reduce((l, a) => layoutReducer(l, a as LayoutAction), newLayout());
    const first = await write(layout);
    expect(text(1, 0)).toBe("2025");
    const year = runner.memberAt(first.def.id, 1, 0)!;
    expect(year.member.hasChildren).toBe(true);
    const opened = layoutReducer(layout, { type: "expand", hierarchy: calendar, member: year.member.uniqueName });
    const second = await write(opened);
    const captions = [...Array(second.rows).keys()].map((row) => text(row, 0));
    expect(captions.slice(1, 6)).toEqual(["2025", "Tehran", "Q1 2025", "Tehran", "Q2 2025"]);
    const expanded = runner.memberAt(second.def.id, 1, 0)!;
    expect(expanded.member.expanded).toBe(true);
    expect(runner.descendants(second.def.id, "rows", expanded.member)).toEqual([]);
    expect(runner.expandable(second.def.id, calendar)?.members.length).toBeGreaterThan(0);
  });
});

describe("outline form through the runner", () => {
  it("writes a column per row field and each outer item on its own line", async () => {
    const { def, rows, cols } = await write(cityServiceByYear("outline", "top"));
    // Row 0: Year caption + years. Row 1: City, Service Type, measures. Row 2: Tehran.
    expect([text(1, 0), text(1, 1)]).toEqual(["City", "Service Type"]);
    expect([text(2, 0), text(2, 1)]).toEqual(["Tehran", ""]);
    expect([text(3, 0), text(3, 1)]).toEqual(["", "Ride"]);
    const at = (row: number, col: number) => ({ row, col, rows, cols });
    expect(runner.buttonAt(def.id, at(1, 1))?.caption).toBe("Service Type");
    expect(runner.labelsAt(def.id, at(1, 1))).toBeNull();
    expect(runner.memberAt(def.id, 3, 0)?.member.uniqueName).toBe(TEHRAN); // blank repeat
    expect(runner.detailsFor(def.id, at(2, 2))?.members).toEqual([TEHRAN, "[Date].[Year].&[2025]"]);
  });
});

describe("the guess from the layout alone, in every form", () => {
  for (const form of ["compact", "outline", "tabular"] as const) {
    it(`agrees with the written ${form} block, cell by cell`, async () => {
      const layouts = [
        cityServiceByYear(form, "top"),
        cityServiceByYear(form, "none", [{ type: "addHierarchy", hierarchy: hierarchy("Segment"), area: "filters" }]),
        cityServiceByYear(form, "bottom", [{ type: "move", name: hierarchy("City").uniqueName, area: "columns" }]),
        cityServiceByYear(form, "top", [
          { type: "remove", name: hierarchy("City").uniqueName },
          { type: "remove", name: hierarchy("Service Type").uniqueName },
        ]),
      ];
      for (const layout of layouts) {
        const { def, rows, cols } = await write(layout);
        for (let row = 0; row < rows; row++) {
          for (let col = 0; col < cols; col++) {
            const guess = guessFieldCell(layout, row, col);
            const guessed = guess ? layout[guess.area][guess.index] : null;
            expect(runner.buttonAt(def.id, { row, col, rows, cols }) ?? null).toEqual(guessed ?? null);
            expect(runner.labelsAt(def.id, { row, col, rows, cols })).toBe(
              guess && guess.labels && guess.area !== "filters" ? guess.area : null,
            );
          }
        }
      }
    });
  }
});

describe("column subtotals with row subtotals at the top", () => {
  it("stay after their group, and Show Details still finds each column's members", async () => {
    const layout = cityServiceByYear("compact", "top", [
      { type: "move", name: hierarchy("Service Type").uniqueName, area: "columns" },
    ]);
    const { def, rows, cols } = await write(layout);
    // Columns: Year > Service Type. Row 0 Column Labels, 1 years, 2 services, 3 measures.
    const years = [...Array(cols).keys()].map((col) => text(1, col));
    const first = years.indexOf("2025 Total");
    expect(years.indexOf("2025")).toBe(1);
    expect(first).toBeGreaterThan(1);
    expect(text(2, first)).toBe("");
    const row = 4; // Tehran
    const target = runner.detailsFor(def.id, { row, col: first, rows, cols });
    expect(target?.members).toEqual([TEHRAN, "[Date].[Year].&[2025]"]);
    const sum = [...Array(first - 1).keys()].reduce((total, i) => total + Number(text(row, 1 + i) || 0), 0);
    expect(Number(text(row, first))).toBe(sum);
  });
});
