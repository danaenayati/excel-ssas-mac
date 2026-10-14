// The layout reducer's actions for KPI parts, repeated measures, named sets, Date
// Filters, groups and calculated items, and calculated measures.

import { describe, expect, it } from "vitest";
import type { Layout } from "../types";
import {
  calculatedMeasureName,
  emptyLayout,
  isFiltered,
  layoutReducer,
  locate,
  valueKey,
  valueLabels,
  type LayoutAction,
} from "./layout";

const rides = { uniqueName: "[Measures].[Rides]", caption: "Rides" };
const city = { uniqueName: "[Geo].[City]", caption: "City", allMember: "[Geo].[City].[All]" };
const kpi = { name: "Sales", caption: "Sales", measure: "[Measures].[Sales]" };
const topCities = { uniqueName: "[Top Cities]", caption: "Top Cities", hierarchy: "[Geo].[City]", allMember: "[Geo].[City].[All]" };

function apply(actions: LayoutAction[], start: Layout = emptyLayout()): Layout {
  return actions.reduce(layoutReducer, start);
}

describe("value fields: copies and KPI parts", () => {
  it("keys each copy of a measure, the first by its measure", () => {
    const layout = apply([
      { type: "addMeasure", measure: rides },
      { type: "duplicateValue", key: "[Measures].[Rides]" },
      { type: "duplicateValue", key: "[Measures].[Rides]" },
      { type: "addKpi", kpi, part: "status" },
    ]);
    expect(layout.values.map((_, i) => valueKey(layout.values, i))).toEqual([
      "[Measures].[Rides]", "[Measures].[Rides]#2", "[Measures].[Rides]#3", "kpi:Sales:status",
    ]);
    expect(valueLabels(layout.values)).toEqual(["Rides", "Rides 2", "Rides 3", "Sales Status"]);
    expect(locate(layout, "[Measures].[Rides]#3")).toEqual({ area: "values", index: 2 });
  });

  it("shows, formats, moves and removes one copy without touching the others", () => {
    let layout = apply([{ type: "addMeasure", measure: rides }, { type: "duplicateValue", key: "[Measures].[Rides]" }]);
    layout = apply(
      [
        { type: "setShowAs", measure: "[Measures].[Rides]#2", showAs: "percentOfGrandTotal" },
        { type: "setNumberFormat", measure: "[Measures].[Rides]#2", format: "0.0%" },
      ],
      layout,
    );
    expect(layout.values[0]).toEqual({ measure: "[Measures].[Rides]", caption: "Rides" });
    expect(layout.values[1]).toMatchObject({ showAs: "percentOfGrandTotal", numberFormat: "0.0%" });
    layout = layoutReducer(layout, { type: "move", name: "[Measures].[Rides]#2", area: "values", index: 0 });
    expect(layout.values[0].showAs).toBe("percentOfGrandTotal");
    // Keys follow position: the copy now first is "[Measures].[Rides]".
    const removed = layoutReducer(layout, { type: "remove", name: "[Measures].[Rides]#2" });
    expect(removed.values).toEqual([layout.values[0]]);
    // Unticking the measure in the field list takes every plain copy, as in Excel.
    const kpiToo = layoutReducer(layout, { type: "addKpi", kpi: { ...kpi, measure: "[Measures].[Rides]" }, part: "goal" });
    expect(layoutReducer(kpiToo, { type: "removeMeasure", measure: "[Measures].[Rides]" }).values).toEqual([kpiToo.values[2]]);
  });

  it("does not take a KPI part for the measure it shows", () => {
    const layout = apply([
      { type: "addKpi", kpi: { ...kpi, measure: "[Measures].[Rides]" }, part: "value" },
      { type: "addMeasure", measure: rides },
      { type: "addKpi", kpi: { ...kpi, measure: "[Measures].[Rides]" }, part: "value" },
    ]);
    expect(layout.values).toEqual([
      { measure: "[Measures].[Rides]", caption: "Sales Value", kpi: { name: "Sales", part: "value" } },
      { measure: "[Measures].[Rides]", caption: "Rides" },
    ]);
    expect(layoutReducer(layout, { type: "duplicateValue", key: "kpi:Sales:value" })).toBe(layout);
  });
});

describe("named sets", () => {
  it("places a set as its hierarchy's field, never as a page filter", () => {
    const layout = apply([{ type: "addSet", set: topCities }]);
    expect(layout.rows).toEqual([
      { hierarchy: "[Geo].[City]", caption: "Top Cities", allMember: "[Geo].[City].[All]", expanded: [], members: [], memberCaptions: [], set: "[Top Cities]" },
    ]);
    expect(layoutReducer(layout, { type: "move", name: "[Geo].[City]", area: "filters" })).toBe(layout);
    expect(layoutReducer(layout, { type: "move", name: "[Geo].[City]", area: "columns" }).columns[0].set).toBe("[Top Cities]");
    expect(layoutReducer(layout, { type: "addSet", set: topCities })).toBe(layout);
    expect(layoutReducer(apply([{ type: "addHierarchy", hierarchy: city }]), { type: "addSet", set: topCities }).rows[0].set).toBeUndefined();
  });
});

describe("Date Filters", () => {
  it("replaces a label filter and is replaced by one, as in Excel", () => {
    let layout = apply([
      { type: "addHierarchy", hierarchy: city },
      { type: "setLabelFilter", hierarchy: city.uniqueName, filter: { operator: "contains", value: "a" } },
      { type: "setDateFilter", hierarchy: city.uniqueName, filter: { operator: "thisMonth" } },
    ]);
    expect(layout.rows[0].labelFilter).toBeUndefined();
    expect(layout.rows[0].dateFilter).toEqual({ operator: "thisMonth" });
    expect(isFiltered(layout.rows[0])).toBe(true);
    layout = layoutReducer(layout, { type: "setLabelFilter", hierarchy: city.uniqueName, filter: { operator: "equals", value: "x" } });
    expect(layout.rows[0].dateFilter).toBeUndefined();
    layout = layoutReducer(layout, { type: "setDateFilter", hierarchy: city.uniqueName, filter: { operator: "today" } });
    layout = layoutReducer(layout, { type: "clearFilters", hierarchy: city.uniqueName });
    expect(layout.rows[0].dateFilter).toBeUndefined();
  });
});

describe("groups and calculated items", () => {
  it("sets and clears a field's items", () => {
    const items = [{ name: "Pair", members: ["[Geo].[City].&[1]", "[Geo].[City].&[2]"] }];
    let layout = apply([{ type: "addHierarchy", hierarchy: city }, { type: "setCustomItems", hierarchy: city.uniqueName, items }]);
    expect(layout.rows[0].customItems).toEqual(items);
    expect(layoutReducer(layout, { type: "setCustomItems", hierarchy: city.uniqueName, items })).toBe(layout);
    layout = layoutReducer(layout, { type: "setCustomItems", hierarchy: city.uniqueName, items: [] });
    expect("customItems" in layout.rows[0]).toBe(false);
  });
});

describe("calculated measures", () => {
  const fare = { name: "Fare", formula: "[Measures].[Revenue] / [Measures].[Rides]" };
  const fareName = calculatedMeasureName("Fare");

  it("adds, renames with its uses, and deletes with its uses", () => {
    let layout = apply([
      { type: "setCalculatedMeasure", measure: fare },
      { type: "addMeasure", measure: { uniqueName: fareName, caption: "Fare" } },
      { type: "addHierarchy", hierarchy: city },
      { type: "setSort", hierarchy: city.uniqueName, sort: { by: "measure", measure: fareName, order: "desc" } },
      { type: "setValueFilter", hierarchy: city.uniqueName, filter: { operator: "top", measure: fareName, value: 3 } },
    ]);
    expect(layout.calculatedMeasures).toEqual([fare]);
    layout = layoutReducer(layout, { type: "setCalculatedMeasure", measure: { ...fare, name: "Avg ]Fare" }, previousName: "Fare" });
    const renamed = "[Measures].[Avg ]]Fare]";
    expect(layout.calculatedMeasures).toEqual([{ ...fare, name: "Avg ]Fare" }]);
    expect(layout.values[0]).toEqual({ measure: renamed, caption: "Avg ]Fare" });
    expect(layout.rows[0].sort?.measure).toBe(renamed);
    expect(layout.rows[0].valueFilter?.measure).toBe(renamed);
    layout = layoutReducer(layout, { type: "removeCalculatedMeasure", name: "Avg ]Fare" });
    expect(layout.values).toEqual([]);
    expect(layout.rows[0].sort).toBeUndefined();
    expect(layout.rows[0].valueFilter).toBeUndefined();
    expect("calculatedMeasures" in layout).toBe(false);
  });

  it("changing nothing changes nothing", () => {
    const layout = apply([{ type: "setCalculatedMeasure", measure: fare }]);
    expect(layoutReducer(layout, { type: "setCalculatedMeasure", measure: { ...fare } })).toBe(layout);
    expect(layoutReducer(layout, { type: "removeCalculatedMeasure", name: "Nope" })).toBe(layout);
  });
});

describe("one level of a user hierarchy", () => {
  const calendar = { uniqueName: "[Date].[Calendar]", caption: "Calendar", allMember: "[Date].[Calendar].[All]" };
  const month = { uniqueName: "[Date].[Calendar].[Month]", number: 3, caption: "Month" };
  const year = { uniqueName: "[Date].[Calendar].[Year]", number: 1, caption: "Year" };

  it("places a level alone, swaps it for another level, and back to the whole field", () => {
    let layout = apply([{ type: "addHierarchy", hierarchy: { ...calendar, level: month } }]);
    expect(layout.rows[0]).toMatchObject({ hierarchy: "[Date].[Calendar]", caption: "Month", level: month });
    layout = apply(
      [
        { type: "expand", hierarchy: calendar.uniqueName, member: "[Date].[Calendar].[Month].&[1]" },
        { type: "move", name: calendar.uniqueName, area: "columns" },
        { type: "addHierarchy", hierarchy: { ...calendar, level: year } },
      ],
      layout,
    );
    expect(layout.columns[0]).toMatchObject({ caption: "Year", level: year, expanded: [] });
    layout = layoutReducer(layout, { type: "addHierarchy", hierarchy: calendar });
    expect(layout.columns[0].caption).toBe("Calendar");
    expect("level" in layout.columns[0]).toBe(false);
  });
});

describe("named sets spanning several fields", () => {
  const pairs = {
    uniqueName: "[Pairs]",
    caption: "Pairs",
    hierarchy: "[Geo].[City]",
    allMember: "[Geo].[City].[All]",
    hierarchies: ["[Geo].[City]", "[S].[T]"],
    allMembers: ["[Geo].[City].[All]", "[S].[T].[All]"],
  };

  it("keeps every field of the set, and blocks those fields elsewhere", () => {
    const layout = apply([{ type: "addSet", set: pairs }]);
    expect(layout.rows[0]).toMatchObject({ set: "[Pairs]", setHierarchies: pairs.hierarchies, setAllMembers: pairs.allMembers });
    const service = { uniqueName: "[S].[T]", caption: "Type", allMember: "[S].[T].[All]" };
    expect(layoutReducer(layout, { type: "addHierarchy", hierarchy: service })).toBe(layout);
    const typeFirst = apply([{ type: "addHierarchy", hierarchy: service }]);
    expect(layoutReducer(typeFirst, { type: "addSet", set: pairs })).toBe(typeFirst);
  });

  it("reads them back from a saved report", async () => {
    const { parseReportDef } = await import("./report");
    const def = parseReportDef(
      JSON.stringify({ id: "abc", layout: { rows: [{ hierarchy: "[Geo].[City]", set: "[Pairs]", setHierarchies: pairs.hierarchies, setAllMembers: pairs.allMembers, level: null }], options: { includeFilteredItems: true } } }),
    )!;
    expect(def.layout.rows[0]).toMatchObject({ setHierarchies: pairs.hierarchies, setAllMembers: pairs.allMembers });
    expect(def.layout.options.includeFilteredItems).toBe(true);
  });
});
