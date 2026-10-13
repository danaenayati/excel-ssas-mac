import { describe, expect, it } from "vitest";
import type { Layout } from "../types";
import {
  emptyLayout,
  filterValueText,
  findField,
  isEmptyLayout,
  layoutReducer,
  locate,
  noRestrictionLabel,
  selectionSummary,
  type HierarchyInput,
  type LayoutAction,
  type MeasureInput,
} from "./layout";

const rides: MeasureInput = { uniqueName: "[Measures].[Rides]", caption: "Rides" };
const fare: MeasureInput = { uniqueName: "[Measures].[Fare]", caption: "Fare" };
const avg: MeasureInput = { uniqueName: "[Measures].[Avg Fare]", caption: "Avg Fare" };

const calendar: HierarchyInput = {
  uniqueName: "[Date].[Calendar]",
  caption: "Calendar",
  allMember: "[Date].[Calendar].[All]",
};
const city: HierarchyInput = {
  uniqueName: "[Geography].[City]",
  caption: "City",
  allMember: "[Geography].[City].[All]",
};
const segment: HierarchyInput = {
  uniqueName: "[Customer].[Segment]",
  caption: "Segment",
  allMember: null,
};

function run(actions: LayoutAction[], start: Layout = emptyLayout()): Layout {
  return actions.reduce(layoutReducer, start);
}

const names = (layout: Layout) => ({
  rows: layout.rows.map((f) => f.caption),
  columns: layout.columns.map((f) => f.caption),
  filters: layout.filters.map((f) => f.caption),
  values: layout.values.map((v) => v.caption),
});

describe("empty layout", () => {
  it("starts with nothing placed and the default options", () => {
    const layout = emptyLayout();
    expect(names(layout)).toEqual({ rows: [], columns: [], filters: [], values: [] });
    expect(layout.options).toEqual({ grandTotals: true, subtotals: false, nonEmpty: true });
    expect(isEmptyLayout(layout)).toBe(true);
  });

  it("is still empty with only a filter, and not empty with a row, column or value", () => {
    const onlyFilter = run([{ type: "addHierarchy", hierarchy: city, area: "filters" }]);
    expect(isEmptyLayout(onlyFilter)).toBe(true);
    expect(isEmptyLayout(run([{ type: "addMeasure", measure: rides }]))).toBe(false);
    expect(isEmptyLayout(run([{ type: "addHierarchy", hierarchy: city }]))).toBe(false);
    expect(
      isEmptyLayout(run([{ type: "addHierarchy", hierarchy: city, area: "columns" }])),
    ).toBe(false);
  });

  it("gives each new layout its own options object", () => {
    const a = emptyLayout();
    const b = emptyLayout();
    expect(a.options).not.toBe(b.options);
  });
});

describe("adding fields", () => {
  it("puts a ticked measure in Values and a ticked hierarchy in Rows", () => {
    const layout = run([
      { type: "addMeasure", measure: rides },
      { type: "addHierarchy", hierarchy: calendar },
    ]);
    expect(names(layout)).toEqual({
      rows: ["Calendar"],
      columns: [],
      filters: [],
      values: ["Rides"],
    });
    expect(layout.values[0]).toEqual({ measure: "[Measures].[Rides]", caption: "Rides" });
    expect(layout.rows[0]).toEqual({
      hierarchy: "[Date].[Calendar]",
      caption: "Calendar",
      allMember: "[Date].[Calendar].[All]",
      expanded: [],
      members: [],
      memberCaptions: [],
    });
  });

  it("keeps a null All member for hierarchies that have none", () => {
    const layout = run([{ type: "addHierarchy", hierarchy: segment }]);
    expect(layout.rows[0].allMember).toBeNull();
  });

  it("appends in the order ticked, or inserts at the given index", () => {
    const layout = run([
      { type: "addMeasure", measure: rides },
      { type: "addMeasure", measure: fare },
      { type: "addMeasure", measure: avg, index: 0 },
      { type: "addHierarchy", hierarchy: calendar },
      { type: "addHierarchy", hierarchy: city, index: 0 },
    ]);
    expect(names(layout).values).toEqual(["Avg Fare", "Rides", "Fare"]);
    expect(names(layout).rows).toEqual(["City", "Calendar"]);
  });

  it("can add a hierarchy straight into Columns or Filters", () => {
    const layout = run([
      { type: "addHierarchy", hierarchy: calendar, area: "columns" },
      { type: "addHierarchy", hierarchy: city, area: "filters" },
    ]);
    expect(names(layout)).toEqual({
      rows: [],
      columns: ["Calendar"],
      filters: ["City"],
      values: [],
    });
  });

  it("does nothing when the same measure or hierarchy is ticked twice", () => {
    const once = run([
      { type: "addMeasure", measure: rides },
      { type: "addHierarchy", hierarchy: calendar },
    ]);
    expect(layoutReducer(once, { type: "addMeasure", measure: rides })).toBe(once);
    expect(layoutReducer(once, { type: "addHierarchy", hierarchy: calendar })).toBe(once);
  });

  it("clamps an out-of-range insertion index", () => {
    const layout = run([
      { type: "addMeasure", measure: rides },
      { type: "addMeasure", measure: fare, index: 99 },
      { type: "addMeasure", measure: avg, index: -3 },
    ]);
    expect(names(layout).values).toEqual(["Avg Fare", "Rides", "Fare"]);
  });
});

describe("one area per hierarchy", () => {
  it("moves a hierarchy instead of duplicating it when added to another area", () => {
    const layout = run([
      { type: "addHierarchy", hierarchy: calendar },
      { type: "addHierarchy", hierarchy: calendar, area: "columns" },
    ]);
    expect(names(layout)).toEqual({
      rows: [],
      columns: ["Calendar"],
      filters: [],
      values: [],
    });
  });

  it("never leaves a hierarchy in two areas, whatever the sequence", () => {
    const layout = run([
      { type: "addHierarchy", hierarchy: calendar },
      { type: "addHierarchy", hierarchy: city },
      { type: "move", name: calendar.uniqueName, area: "filters" },
      { type: "addHierarchy", hierarchy: calendar, area: "rows" },
      { type: "move", name: city.uniqueName, area: "columns" },
      { type: "addHierarchy", hierarchy: city, area: "columns" },
      { type: "move", name: calendar.uniqueName, area: "columns", index: 0 },
    ]);
    const all = [...layout.rows, ...layout.columns, ...layout.filters].map((f) => f.hierarchy);
    expect(all.slice().sort()).toEqual([calendar.uniqueName, city.uniqueName].sort());
    expect(new Set(all).size).toBe(all.length);
    expect(names(layout).columns).toEqual(["Calendar", "City"]);
  });

  it("reports where a field sits", () => {
    const layout = run([
      { type: "addMeasure", measure: rides },
      { type: "addMeasure", measure: fare },
      { type: "addHierarchy", hierarchy: calendar, area: "filters" },
    ]);
    expect(locate(layout, fare.uniqueName)).toEqual({ area: "values", index: 1 });
    expect(locate(layout, calendar.uniqueName)).toEqual({ area: "filters", index: 0 });
    expect(locate(layout, city.uniqueName)).toBeNull();
    expect(findField(layout, calendar.uniqueName)?.caption).toBe("Calendar");
    expect(findField(layout, rides.uniqueName)).toBeNull();
  });
});

describe("removing fields", () => {
  it("removes a hierarchy from whichever area holds it", () => {
    const start = run([
      { type: "addHierarchy", hierarchy: calendar, area: "columns" },
      { type: "addHierarchy", hierarchy: city, area: "filters" },
      { type: "addHierarchy", hierarchy: segment },
    ]);
    const layout = run(
      [
        { type: "remove", name: calendar.uniqueName },
        { type: "remove", name: city.uniqueName },
        { type: "remove", name: segment.uniqueName },
      ],
      start,
    );
    expect(names(layout)).toEqual({ rows: [], columns: [], filters: [], values: [] });
  });

  it("removes a measure and keeps the others in order", () => {
    const layout = run([
      { type: "addMeasure", measure: rides },
      { type: "addMeasure", measure: fare },
      { type: "addMeasure", measure: avg },
      { type: "remove", name: fare.uniqueName },
    ]);
    expect(names(layout).values).toEqual(["Rides", "Avg Fare"]);
  });

  it("returns the same layout when the field is not placed", () => {
    const layout = run([{ type: "addMeasure", measure: rides }]);
    expect(layoutReducer(layout, { type: "remove", name: "[Nope].[Nope]" })).toBe(layout);
  });
});

describe("moving fields", () => {
  const start = run([
    { type: "addHierarchy", hierarchy: calendar },
    { type: "addHierarchy", hierarchy: city },
    { type: "addHierarchy", hierarchy: segment },
    { type: "addMeasure", measure: rides },
    { type: "addMeasure", measure: fare },
    { type: "addMeasure", measure: avg },
  ]);

  it("moves a hierarchy between Rows, Columns and Filters", () => {
    const toColumns = layoutReducer(start, {
      type: "move",
      name: city.uniqueName,
      area: "columns",
    });
    expect(names(toColumns)).toMatchObject({ rows: ["Calendar", "Segment"], columns: ["City"] });

    const toFilters = layoutReducer(toColumns, {
      type: "move",
      name: city.uniqueName,
      area: "filters",
    });
    expect(names(toFilters)).toMatchObject({ columns: [], filters: ["City"] });

    const back = layoutReducer(toFilters, {
      type: "move",
      name: city.uniqueName,
      area: "rows",
      index: 0,
    });
    expect(names(back)).toMatchObject({ rows: ["City", "Calendar", "Segment"], filters: [] });
  });

  it("keeps expanded and members when a hierarchy moves to another area", () => {
    const prepared = run(
      [
        { type: "expand", hierarchy: calendar.uniqueName, member: "[Date].[Calendar].[2025]" },
        {
          type: "setMembers",
          hierarchy: calendar.uniqueName,
          members: ["[Date].[Calendar].[2025]", "[Date].[Calendar].[2026]"],
          captions: ["2025", "2026"],
        },
      ],
      start,
    );
    const before = findField(prepared, calendar.uniqueName)!;

    for (const area of ["columns", "filters", "rows"] as const) {
      const moved = layoutReducer(prepared, { type: "move", name: calendar.uniqueName, area });
      const after = findField(moved, calendar.uniqueName)!;
      expect(locate(moved, calendar.uniqueName)?.area).toBe(area);
      expect(after.expanded).toEqual(["[Date].[Calendar].[2025]"]);
      expect(after.members).toEqual(before.members);
      expect(after.memberCaptions).toEqual(["2025", "2026"]);
      expect(after.allMember).toBe(before.allMember);
    }
  });

  it("keeps expanded and members when re-added through addHierarchy with an area", () => {
    const prepared = run(
      [{ type: "expand", hierarchy: city.uniqueName, member: "[Geography].[City].[Tehran]" }],
      start,
    );
    const moved = layoutReducer(prepared, {
      type: "addHierarchy",
      hierarchy: city,
      area: "columns",
    });
    expect(findField(moved, city.uniqueName)!.expanded).toEqual(["[Geography].[City].[Tehran]"]);
  });

  it("reorders inside an area with an index counted after removal", () => {
    const last = layoutReducer(start, { type: "move", name: calendar.uniqueName, area: "rows" });
    expect(names(last).rows).toEqual(["City", "Segment", "Calendar"]);

    const middle = layoutReducer(start, {
      type: "move",
      name: segment.uniqueName,
      area: "rows",
      index: 1,
    });
    expect(names(middle).rows).toEqual(["Calendar", "Segment", "City"]);
  });

  it("returns the same layout when a move would not change anything", () => {
    expect(
      layoutReducer(start, { type: "move", name: city.uniqueName, area: "rows", index: 1 }),
    ).toBe(start);
    expect(
      layoutReducer(start, { type: "move", name: fare.uniqueName, area: "values", index: 1 }),
    ).toBe(start);
    expect(layoutReducer(start, { type: "move", name: "[X].[Y]", area: "rows" })).toBe(start);
  });

  it("moves up and down one step and stops at the ends", () => {
    const down = layoutReducer(start, { type: "shift", name: calendar.uniqueName, delta: 1 });
    expect(names(down).rows).toEqual(["City", "Calendar", "Segment"]);

    const up = layoutReducer(start, { type: "shift", name: segment.uniqueName, delta: -1 });
    expect(names(up).rows).toEqual(["Calendar", "Segment", "City"]);

    expect(layoutReducer(start, { type: "shift", name: calendar.uniqueName, delta: -1 })).toBe(
      start,
    );
    expect(layoutReducer(start, { type: "shift", name: segment.uniqueName, delta: 1 })).toBe(
      start,
    );
  });

  it("reorders measures inside Values", () => {
    const down = layoutReducer(start, { type: "shift", name: rides.uniqueName, delta: 1 });
    expect(names(down).values).toEqual(["Fare", "Rides", "Avg Fare"]);
    const first = layoutReducer(start, {
      type: "move",
      name: avg.uniqueName,
      area: "values",
      index: 0,
    });
    expect(names(first).values).toEqual(["Avg Fare", "Rides", "Fare"]);
  });

  it("refuses to put a measure in a hierarchy area or a hierarchy in Values", () => {
    expect(layoutReducer(start, { type: "move", name: rides.uniqueName, area: "rows" })).toBe(
      start,
    );
    expect(layoutReducer(start, { type: "move", name: city.uniqueName, area: "values" })).toBe(
      start,
    );
  });

  it("does not mutate the previous layout", () => {
    const snapshot = JSON.stringify(start);
    run(
      [
        { type: "move", name: city.uniqueName, area: "columns" },
        { type: "shift", name: rides.uniqueName, delta: 1 },
        { type: "remove", name: segment.uniqueName },
        { type: "expand", hierarchy: calendar.uniqueName, member: "[Date].[Calendar].[2025]" },
        { type: "setOption", option: "subtotals", value: true },
      ],
      start,
    );
    expect(JSON.stringify(start)).toBe(snapshot);
  });
});

describe("member filter", () => {
  const start = run([{ type: "addHierarchy", hierarchy: city, area: "filters" }]);

  it("stores the selected members with their captions", () => {
    const layout = layoutReducer(start, {
      type: "setMembers",
      hierarchy: city.uniqueName,
      members: ["[Geography].[City].[Tehran]", "[Geography].[City].[Shiraz]"],
      captions: ["Tehran", "Shiraz"],
    });
    expect(layout.filters[0].members).toEqual([
      "[Geography].[City].[Tehran]",
      "[Geography].[City].[Shiraz]",
    ]);
    expect(layout.filters[0].memberCaptions).toEqual(["Tehran", "Shiraz"]);
  });

  it("clears the restriction with an empty list", () => {
    const filtered = layoutReducer(start, {
      type: "setMembers",
      hierarchy: city.uniqueName,
      members: ["[Geography].[City].[Tehran]"],
      captions: ["Tehran"],
    });
    const cleared = layoutReducer(filtered, {
      type: "setMembers",
      hierarchy: city.uniqueName,
      members: [],
      captions: [],
    });
    expect(cleared.filters[0].members).toEqual([]);
    expect(cleared.filters[0].memberCaptions).toEqual([]);
  });

  it("keeps captions the same length as members", () => {
    const layout = layoutReducer(start, {
      type: "setMembers",
      hierarchy: city.uniqueName,
      members: ["[Geography].[City].[Tehran]", "[Geography].[City].[Shiraz]"],
      captions: ["Tehran"],
    });
    expect(layout.filters[0].memberCaptions).toEqual(["Tehran", "[Geography].[City].[Shiraz]"]);
  });

  it("returns the same layout when the selection is unchanged or the field is not placed", () => {
    const filtered = layoutReducer(start, {
      type: "setMembers",
      hierarchy: city.uniqueName,
      members: ["[Geography].[City].[Tehran]"],
      captions: ["Tehran"],
    });
    expect(
      layoutReducer(filtered, {
        type: "setMembers",
        hierarchy: city.uniqueName,
        members: ["[Geography].[City].[Tehran]"],
        captions: ["Tehran"],
      }),
    ).toBe(filtered);
    expect(
      layoutReducer(start, {
        type: "setMembers",
        hierarchy: calendar.uniqueName,
        members: ["x"],
        captions: ["x"],
      }),
    ).toBe(start);
  });
});

describe("expand and collapse", () => {
  const y2025 = "[Date].[Calendar].[2025]";
  const q1 = "[Date].[Calendar].[2025].[Q1]";
  const q2 = "[Date].[Calendar].[2025].[Q2]";
  const y2026 = "[Date].[Calendar].[2026]";
  const start = run([
    { type: "addHierarchy", hierarchy: calendar },
    { type: "addHierarchy", hierarchy: city, area: "columns" },
  ]);

  it("expand adds the member to that field only", () => {
    const layout = layoutReducer(start, {
      type: "expand",
      hierarchy: calendar.uniqueName,
      member: y2025,
    });
    expect(layout.rows[0].expanded).toEqual([y2025]);
    expect(layout.columns[0].expanded).toEqual([]);
    expect(layout.columns).toBe(start.columns);
  });

  it("expand is idempotent", () => {
    const once = layoutReducer(start, {
      type: "expand",
      hierarchy: calendar.uniqueName,
      member: y2025,
    });
    expect(
      layoutReducer(once, { type: "expand", hierarchy: calendar.uniqueName, member: y2025 }),
    ).toBe(once);
  });

  it("works for a hierarchy on Columns too", () => {
    const layout = layoutReducer(start, {
      type: "expand",
      hierarchy: city.uniqueName,
      member: "[Geography].[City].[Tehran]",
    });
    expect(layout.columns[0].expanded).toEqual(["[Geography].[City].[Tehran]"]);
  });

  it("collapse removes the member", () => {
    const expanded = run(
      [
        { type: "expand", hierarchy: calendar.uniqueName, member: y2025 },
        { type: "expand", hierarchy: calendar.uniqueName, member: y2026 },
      ],
      start,
    );
    const layout = layoutReducer(expanded, {
      type: "collapse",
      hierarchy: calendar.uniqueName,
      member: y2025,
    });
    expect(layout.rows[0].expanded).toEqual([y2026]);
  });

  it("collapse also removes the expanded descendants it is told about", () => {
    const expanded = run(
      [
        { type: "expand", hierarchy: calendar.uniqueName, member: y2025 },
        { type: "expand", hierarchy: calendar.uniqueName, member: q1 },
        { type: "expand", hierarchy: calendar.uniqueName, member: q2 },
        { type: "expand", hierarchy: calendar.uniqueName, member: y2026 },
      ],
      start,
    );
    const layout = layoutReducer(expanded, {
      type: "collapse",
      hierarchy: calendar.uniqueName,
      member: y2025,
      descendants: [q1, q2, "[Date].[Calendar].[2025].[Q3]"],
    });
    expect(layout.rows[0].expanded).toEqual([y2026]);
  });

  it("returns the same layout when there is nothing to collapse", () => {
    expect(
      layoutReducer(start, { type: "collapse", hierarchy: calendar.uniqueName, member: y2025 }),
    ).toBe(start);
    expect(layoutReducer(start, { type: "expand", hierarchy: "[X].[Y]", member: y2025 })).toBe(
      start,
    );
  });
});

describe("options", () => {
  it("toggles each option on its own", () => {
    const start = emptyLayout();
    const subtotals = layoutReducer(start, { type: "setOption", option: "subtotals", value: true });
    expect(subtotals.options).toEqual({ grandTotals: true, subtotals: true, nonEmpty: true });

    const noTotals = layoutReducer(subtotals, {
      type: "setOption",
      option: "grandTotals",
      value: false,
    });
    expect(noTotals.options).toEqual({ grandTotals: false, subtotals: true, nonEmpty: true });

    const showEmpty = layoutReducer(noTotals, {
      type: "setOption",
      option: "nonEmpty",
      value: false,
    });
    expect(showEmpty.options).toEqual({ grandTotals: false, subtotals: true, nonEmpty: false });
  });

  it("returns the same layout when the value does not change", () => {
    const start = emptyLayout();
    expect(layoutReducer(start, { type: "setOption", option: "grandTotals", value: true })).toBe(
      start,
    );
  });

  it("keeps the fields when an option changes", () => {
    const start = run([
      { type: "addMeasure", measure: rides },
      { type: "addHierarchy", hierarchy: calendar },
    ]);
    const next = layoutReducer(start, { type: "setOption", option: "subtotals", value: true });
    expect(next.rows).toBe(start.rows);
    expect(next.values).toBe(start.values);
  });
});

describe("replace", () => {
  it("swaps in a whole layout, as when another report becomes active", () => {
    const other = run([{ type: "addMeasure", measure: fare }]);
    expect(layoutReducer(emptyLayout(), { type: "replace", layout: other })).toBe(other);
  });
});

describe("how a field's selection reads", () => {
  const filtered = (hierarchy: HierarchyInput, captions: string[]) => {
    const layout = run([
      { type: "addHierarchy", hierarchy, area: "filters" },
      {
        type: "setMembers",
        hierarchy: hierarchy.uniqueName,
        members: captions.map((c) => `${hierarchy.uniqueName}.[${c}]`),
        captions,
      },
    ]);
    return layout.filters[0];
  };

  it("is All, or (Default) without an All member, when nothing is selected", () => {
    expect(noRestrictionLabel(city)).toBe("All");
    expect(noRestrictionLabel(segment)).toBe("(Default)");
    expect(filterValueText(filtered(city, []))).toBe("All");
    expect(filterValueText(filtered(segment, []))).toBe("(Default)");
    expect(selectionSummary(filtered(city, []))).toBe("All");
    expect(selectionSummary(filtered(segment, []))).toBe("(Default)");
  });

  it("is the member's caption for one member, on the sheet and on the chip", () => {
    expect(filterValueText(filtered(city, ["Tehran"]))).toBe("Tehran");
    expect(selectionSummary(filtered(city, ["Tehran"]))).toBe("Tehran");
  });

  it("is (Multiple Items) on the sheet and a count on the chip for several", () => {
    const three = filtered(city, ["Tehran", "Shiraz", "Rey"]);
    expect(filterValueText(three)).toBe("(Multiple Items)");
    expect(selectionSummary(three)).toBe("3 selected");
    expect(selectionSummary(filtered(city, ["Tehran", "Shiraz"]))).toBe("2 selected");
  });

  it("is unchanged by moving the field to another area", () => {
    const start = run([
      { type: "addHierarchy", hierarchy: city, area: "filters" },
      {
        type: "setMembers",
        hierarchy: city.uniqueName,
        members: ["[Geography].[City].[Tehran]"],
        captions: ["Tehran"],
      },
    ]);
    const moved = layoutReducer(start, { type: "move", name: city.uniqueName, area: "rows" });
    expect(selectionSummary(moved.rows[0])).toBe("Tehran");
  });
});
