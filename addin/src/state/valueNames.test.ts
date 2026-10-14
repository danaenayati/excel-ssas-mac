import { describe, expect, it } from "vitest";
import { blockFields, buildBlock } from "../sheet/layout";
import { createMockApi } from "../dev/mockApi";
import type { Layout, ValueField } from "../types";
import { layoutReducer, newLayout, valueKey, type LayoutAction } from "./layout";
import { parseLayout } from "./report";
import { customNameOf, customNameProblem, valueFieldNames } from "./valueNames";

const RIDES = "[Measures].[Rides]";
const FARE = "[Measures].[Fare]";
const city = { uniqueName: "[Geography].[City]", caption: "City", allMember: "[Geography].[City].[All]" };

const run = (actions: LayoutAction[], start: Layout = newLayout()): Layout => actions.reduce(layoutReducer, start);

/** Rides, a second Rides, Fare as % of Grand Total; City on rows. */
function sample(): Layout {
  return run([
    { type: "addHierarchy", hierarchy: city },
    { type: "addMeasure", measure: { uniqueName: RIDES, caption: "Rides" } },
    { type: "duplicateValue", key: RIDES },
    { type: "addMeasure", measure: { uniqueName: FARE, caption: "Fare" } },
    { type: "setShowAs", measure: FARE, showAs: "percentOfGrandTotal" },
  ]);
}

describe("naming a value field", () => {
  it("stores the name, trimmed, on that field only; empty gives the cube's back", () => {
    const named = layoutReducer(sample(), { type: "setCustomName", measure: `${RIDES}#2`, name: "  Trips  " });
    expect(named.values.map((v) => v.customName)).toEqual([undefined, "Trips", undefined]);
    const back = layoutReducer(named, { type: "setCustomName", measure: `${RIDES}#2`, name: "" });
    expect("customName" in back.values[1]).toBe(false);
    expect("customName" in layoutReducer(named, { type: "setCustomName", measure: `${RIDES}#2`, name: null }).values[1]).toBe(false);
  });

  it("changes nothing for the same name or an unknown field", () => {
    const named = layoutReducer(sample(), { type: "setCustomName", measure: RIDES, name: "Trips" });
    expect(layoutReducer(named, { type: "setCustomName", measure: RIDES, name: "Trips " })).toBe(named);
    expect(layoutReducer(named, { type: "setCustomName", measure: "[Measures].[Nope]", name: "X" })).toBe(named);
    const plain = sample();
    expect(layoutReducer(plain, { type: "setCustomName", measure: RIDES, name: "  " })).toBe(plain);
  });

  it("keeps the copies' keys: a renamed first copy is still the measure's key", () => {
    const named = layoutReducer(sample(), { type: "setCustomName", measure: RIDES, name: "Trips" });
    expect(named.values.map((_, i) => valueKey(named.values, i))).toEqual([RIDES, `${RIDES}#2`, FARE]);
    // A rename keeps Show Values As, and Show Values As keeps the rename.
    const both = run(
      [
        { type: "setCustomName", measure: FARE, name: "Fare share" },
        { type: "setShowAs", measure: FARE, showAs: "percentOfColumnTotal" },
      ],
      named,
    );
    expect(both.values[2]).toMatchObject({ customName: "Fare share", showAs: "percentOfColumnTotal" });
  });

  it("names chips by the custom name, else as the helper heads them", () => {
    const named = layoutReducer(sample(), { type: "setCustomName", measure: `${RIDES}#2`, name: "Trips" });
    expect(valueFieldNames(sample().values)).toEqual(["Rides", "Rides 2", "Fare"]);
    expect(valueFieldNames(named.values)).toEqual(["Rides", "Trips", "Fare"]);
    expect(customNameOf({ customName: "  " })).toBeNull();
    expect(customNameOf({})).toBeNull();
  });
});

describe("refusing a name in use, as Excel does", () => {
  const layout = sample();

  it("refuses another value field's name, whatever the case", () => {
    expect(customNameProblem(layout, `${RIDES}#2`, "rides")).toMatch(/already has a field called “rides”/);
    expect(customNameProblem(layout, RIDES, "Rides 2")).not.toBeNull();
    expect(customNameProblem(layout, RIDES, "fare")).not.toBeNull();
  });

  it("refuses the name of a field in Rows, Columns, Filters or a slicer", () => {
    expect(customNameProblem(layout, RIDES, "City")).not.toBeNull();
    const withSlicer: Layout = {
      ...layout,
      slicers: [{ hierarchy: "[Service].[Type]", caption: "Service Type", allMember: null, expanded: [], members: [], memberCaptions: [] }],
    };
    expect(customNameProblem(withSlicer, RIDES, "service type")).not.toBeNull();
  });

  it("accepts a new name, the field's own name, and empty", () => {
    expect(customNameProblem(layout, RIDES, "Trips")).toBeNull();
    expect(customNameProblem(layout, `${RIDES}#2`, "Rides 2")).toBeNull();
    expect(customNameProblem(layout, RIDES, "   ")).toBeNull();
  });

  it("compares with other fields' custom names, not their old ones", () => {
    const named = layoutReducer(layout, { type: "setCustomName", measure: FARE, name: "Money" });
    expect(customNameProblem(named, RIDES, "Fare")).toBeNull();
    expect(customNameProblem(named, RIDES, "Money")).not.toBeNull();
  });

  it("refuses a name longer than Excel allows", () => {
    expect(customNameProblem(layout, RIDES, "x".repeat(256))).toMatch(/at most 255/);
  });
});

describe("reading custom names back", () => {
  it("keeps a text name, trimmed, also on a KPI part", () => {
    const layout = parseLayout({
      values: [
        { measure: RIDES, caption: "Rides", customName: " Trips " },
        { measure: "[Measures].[Score]", caption: "Score Value", kpi: { name: "Score", part: "value" }, customName: "Grade" },
        { measure: FARE, caption: "Fare", customName: "   " },
        { measure: FARE, caption: "Fare", customName: 7 },
      ],
    });
    expect(layout.values.map((v: ValueField) => v.customName)).toEqual(["Trips", "Grade", undefined, undefined]);
  });
});

describe("the header cells say the custom name", () => {
  it("in every column of that value field, the others as the helper heads them", async () => {
    const api = createMockApi({ delayMs: 0 });
    const layout = run(
      [
        { type: "addHierarchy", hierarchy: { uniqueName: "[Date].[Year]", caption: "Year", allMember: "[Date].[Year].[All]" }, area: "columns" },
        { type: "setCustomName", measure: `${RIDES}#2`, name: "Trips" },
        { type: "setCustomName", measure: FARE, name: "Fare share" },
      ],
      sample(),
    );
    const grid = await api.query("demo", { catalog: "Sample", cube: "Rides", layout });
    const block = buildBlock(grid, blockFields(layout));
    const measureRow = block.values[block.shape.dataTop - 1].slice(block.shape.headerCols);
    // 2025, 2026 and Grand Total, each with the three value fields.
    expect(measureRow).toEqual(["Rides", "Trips", "Fare share", "Rides", "Trips", "Fare share", "Rides", "Trips", "Fare share"]);
    // Without names the helper's own headers stay, Show Values As caption included.
    const plain = buildBlock(grid, blockFields(sample()));
    const helperHeaders = grid.colHeaderRows[grid.colHeaderRows.length - 1].map((cell) => cell.text);
    expect(plain.values[plain.shape.dataTop - 1].slice(plain.shape.headerCols)).toEqual(helperHeaders);
    expect(helperHeaders[2]).toBe("Fare (% of Grand Total)");
  });
});
