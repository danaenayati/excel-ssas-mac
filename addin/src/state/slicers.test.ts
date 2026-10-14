import { describe, expect, it } from "vitest";
import { createMockApi } from "../dev/mockApi";
import type { Grid, Layout, PlacedField } from "../types";
import {
  MAX_SLICERS,
  findField,
  layoutReducer,
  locate,
  newLayout,
  slicerIndex,
  type HierarchyInput,
  type LayoutAction,
} from "./layout";
import { parseLayout } from "./report";
import { isSlicerSelected, slicerClick, slicerExtras, type SlicerItem } from "./slicers";

const CITY = "[Geography].[City]";
const city: HierarchyInput = { uniqueName: CITY, caption: "City", allMember: "[Geography].[City].[All]" };
const year: HierarchyInput = { uniqueName: "[Date].[Year]", caption: "Year", allMember: "[Date].[Year].[All]" };
const service: HierarchyInput = {
  uniqueName: "[Service].[Service Type]",
  caption: "Service Type",
  allMember: "[Service].[Service Type].[All]",
};
const rides = { uniqueName: "[Measures].[Rides]", caption: "Rides" };

const run = (actions: LayoutAction[], start: Layout = newLayout()): Layout => actions.reduce(layoutReducer, start);
const slicerNames = (layout: Layout) => (layout.slicers ?? []).map((s) => s.caption);

describe("adding and removing slicers", () => {
  it("adds a field as a slicer at the end, with nothing chosen", () => {
    const layout = run([
      { type: "addSlicer", hierarchy: city },
      { type: "addSlicer", hierarchy: year },
    ]);
    expect(slicerNames(layout)).toEqual(["City", "Year"]);
    expect(layout.slicers?.[0]).toEqual({
      hierarchy: CITY,
      caption: "City",
      allMember: city.allMember,
      expanded: [],
      members: [],
      memberCaptions: [],
    });
    expect(slicerIndex(layout, year.uniqueName)).toBe(1);
  });

  it("changes nothing for a field that already has a slicer", () => {
    const layout = run([{ type: "addSlicer", hierarchy: city }]);
    expect(layoutReducer(layout, { type: "addSlicer", hierarchy: city })).toBe(layout);
  });

  it("moves a field out of Rows, Columns or Filters into a slicer, keeping its chosen items", () => {
    const start = run([
      { type: "addHierarchy", hierarchy: city, area: "filters" },
      { type: "setMembers", hierarchy: CITY, members: ["[Geography].[City].&[1]"], captions: ["Tehran"] },
      { type: "addHierarchy", hierarchy: year, area: "columns" },
    ]);
    const layout = run([{ type: "addSlicer", hierarchy: city }, { type: "addSlicer", hierarchy: year }], start);
    expect(layout.filters).toEqual([]);
    expect(layout.columns).toEqual([]);
    expect(slicerNames(layout)).toEqual(["City", "Year"]);
    expect(layout.slicers?.[0]).toMatchObject({ members: ["[Geography].[City].&[1]"], memberCaptions: ["Tehran"] });
  });

  it("keeps only what a slicer uses of a row field: no drill state, sort or filters", () => {
    const start = run([
      { type: "addHierarchy", hierarchy: city },
      { type: "expand", hierarchy: CITY, member: "[Geography].[City].&[1]" },
      { type: "setSort", hierarchy: CITY, sort: { by: "caption", order: "desc" } },
      { type: "setLabelFilter", hierarchy: CITY, filter: { operator: "contains", value: "a" } },
    ]);
    const slicer = run([{ type: "addSlicer", hierarchy: city }], start).slicers?.[0] as PlacedField;
    expect(slicer.expanded).toEqual([]);
    expect(slicer.sort).toBeUndefined();
    expect(slicer.labelFilter).toBeUndefined();
  });

  it("does not make a named set a slicer", () => {
    const start = run([
      { type: "addSet", set: { uniqueName: "[Top Cities]", caption: "Top Cities", hierarchy: CITY, allMember: null } },
    ]);
    expect(layoutReducer(start, { type: "addSlicer", hierarchy: city })).toBe(start);
  });

  it("removes a slicer, and leaves no slicers key behind", () => {
    const layout = run([{ type: "addSlicer", hierarchy: city }, { type: "removeSlicer", hierarchy: CITY }]);
    expect("slicers" in layout).toBe(false);
    expect(layoutReducer(layout, { type: "removeSlicer", hierarchy: CITY })).toBe(layout);
    // "remove" works on a slicer too.
    expect("slicers" in run([{ type: "addSlicer", hierarchy: city }, { type: "remove", name: CITY }])).toBe(false);
  });

  it("stops at 16 slicers", () => {
    const many = Array.from({ length: MAX_SLICERS + 2 }, (_, i): LayoutAction => ({
      type: "addSlicer",
      hierarchy: { uniqueName: `[D].[H${i}]`, caption: `H${i}`, allMember: null },
    }));
    expect(run(many).slicers).toHaveLength(MAX_SLICERS);
  });

  it("setSlicers makes exactly these the slicers, in this order", () => {
    const start = run([
      { type: "addSlicer", hierarchy: city },
      { type: "setMembers", hierarchy: CITY, members: ["[Geography].[City].&[2]"], captions: ["Rey"] },
      { type: "addSlicer", hierarchy: year },
      { type: "addHierarchy", hierarchy: service, area: "rows" },
    ]);
    const layout = layoutReducer(start, { type: "setSlicers", hierarchies: [service, city] });
    expect(slicerNames(layout)).toEqual(["Service Type", "City"]);
    expect(layout.rows).toEqual([]);
    // City keeps its chosen item; Year is gone.
    expect(layout.slicers?.[1].members).toEqual(["[Geography].[City].&[2]"]);
    expect(layoutReducer(layout, { type: "setSlicers", hierarchies: [service, city] })).toBe(layout);
    expect("slicers" in layoutReducer(layout, { type: "setSlicers", hierarchies: [] })).toBe(false);
  });
});

describe("one place per field", () => {
  const sliced = () =>
    run([
      { type: "addSlicer", hierarchy: city },
      { type: "setMembers", hierarchy: CITY, members: ["[Geography].[City].&[1]"], captions: ["Tehran"] },
    ]);

  it("ticking a slicer's field in the field list moves it to Rows, keeping its items", () => {
    const layout = layoutReducer(sliced(), { type: "addHierarchy", hierarchy: city });
    expect("slicers" in layout).toBe(false);
    expect(layout.rows.map((f) => f.caption)).toEqual(["City"]);
    expect(layout.rows[0].members).toEqual(["[Geography].[City].&[1]"]);
  });

  it("dropping it on an area, or moving it there, puts it there", () => {
    expect(layoutReducer(sliced(), { type: "addHierarchy", hierarchy: city, area: "filters", index: 0 }).filters[0].hierarchy).toBe(CITY);
    const moved = layoutReducer(sliced(), { type: "move", name: CITY, area: "columns" });
    expect(moved.columns[0].hierarchy).toBe(CITY);
    expect(moved.slicers).toBeUndefined();
    expect(layoutReducer(sliced(), { type: "move", name: CITY, area: "values" }).slicers).toHaveLength(1);
  });

  it("a named set of the field replaces its slicer", () => {
    const layout = layoutReducer(sliced(), {
      type: "addSet",
      set: { uniqueName: "[Top Cities]", caption: "Top Cities", hierarchy: CITY, allMember: null },
    });
    expect(layout.slicers).toBeUndefined();
    expect(layout.rows[0].set).toBe("[Top Cities]");
  });

  it("a slicer is not in an area: locate and findField do not see it", () => {
    expect(locate(sliced(), CITY)).toBeNull();
    expect(findField(sliced(), CITY)).toBeNull();
  });
});

describe("choosing a slicer's items", () => {
  it("setMembers and clearFilters work on a slicer as on a page filter", () => {
    const layout = run([
      { type: "addSlicer", hierarchy: city },
      { type: "setMembers", hierarchy: CITY, members: ["a", "b"], captions: ["A", "B"] },
    ]);
    expect(layout.slicers?.[0]).toMatchObject({ members: ["a", "b"], memberCaptions: ["A", "B"] });
    const same = layoutReducer(layout, { type: "setMembers", hierarchy: CITY, members: ["a", "b"], captions: ["A", "B"] });
    expect(same).toBe(layout);
    expect(layoutReducer(layout, { type: "clearFilters", hierarchy: CITY }).slicers?.[0].members).toEqual([]);
  });
});

describe("clicking a slicer's buttons", () => {
  const items: SlicerItem[] = [
    { uniqueName: "t", caption: "Tehran" },
    { uniqueName: "s", caption: "Shiraz" },
    { uniqueName: "m", caption: "Mashhad" },
  ];
  const field = (members: string[]) => ({ members, memberCaptions: members.map((m) => m.toUpperCase()) });

  it("shows every button selected while nothing is filtered", () => {
    expect(items.every((i) => isSlicerSelected(field([]), i.uniqueName))).toBe(true);
    expect(isSlicerSelected(field(["t"]), "s")).toBe(false);
  });

  it("a plain click selects that item only", () => {
    expect(slicerClick(field([]), items, false, items[1], false)).toEqual({ members: ["s"], captions: ["Shiraz"] });
    expect(slicerClick(field(["t", "m"]), items, false, items[1], false)).toEqual({ members: ["s"], captions: ["Shiraz"] });
    expect(slicerClick(field(["s"]), items, false, items[1], false)).toEqual({ unchanged: true });
  });

  it("Multi-Select with nothing filtered takes the clicked item out, as Excel does", () => {
    expect(slicerClick(field([]), items, false, items[0], true)).toEqual({ members: ["s", "m"], captions: ["Shiraz", "Mashhad"] });
  });

  it("Multi-Select adds and removes, in the list's order", () => {
    expect(slicerClick(field(["m"]), items, false, items[0], true)).toEqual({ members: ["t", "m"], captions: ["Tehran", "Mashhad"] });
    expect(slicerClick(field(["t", "m"]), items, false, items[2], true)).toEqual({ members: ["t"], captions: ["Tehran"] });
  });

  it("selecting every item again is no filter", () => {
    expect(slicerClick(field(["t", "s"]), items, false, items[2], true)).toEqual({ members: [], captions: [] });
  });

  it("keeps the last selected item, and will not leave one out of a list it has not all of", () => {
    expect(slicerClick(field(["t"]), items, false, items[0], true)).toEqual({ blocked: "last" });
    expect(slicerClick(field([]), items, true, items[0], true)).toEqual({ blocked: "truncated" });
    // With part of a long list, adding is fine and never turns into "all".
    expect(slicerClick(field(["t", "s"]), items, true, items[2], true)).toEqual({
      members: ["t", "s", "m"],
      captions: ["Tehran", "Shiraz", "Mashhad"],
    });
  });

  it("keeps items chosen elsewhere on a multi-select click, drops them on a plain one", () => {
    const deep = { members: ["jan", "t"], memberCaptions: ["Jan 2025", "Tehran"] };
    expect(slicerExtras(deep, items)).toEqual([{ uniqueName: "jan", caption: "Jan 2025" }]);
    expect(slicerClick(deep, items, false, items[1], true)).toEqual({
      members: ["t", "s", "jan"],
      captions: ["Tehran", "Shiraz", "Jan 2025"],
    });
    expect(slicerClick(deep, items, false, items[1], false)).toEqual({ members: ["s"], captions: ["Shiraz"] });
  });
});

describe("reading slicers back", () => {
  it("reads slicers, keeping what a slicer uses", () => {
    const layout = parseLayout({
      slicers: [
        { hierarchy: CITY, caption: "City", allMember: city.allMember, members: ["t"], memberCaptions: ["Tehran"], expanded: ["x"], sort: { by: "caption", order: "asc" } },
        { hierarchy: year.uniqueName, caption: "Year", allMember: null, members: [], memberCaptions: [] },
      ],
    });
    expect(layout.slicers).toEqual([
      { hierarchy: CITY, caption: "City", allMember: city.allMember, expanded: [], members: ["t"], memberCaptions: ["Tehran"] },
      { hierarchy: year.uniqueName, caption: "Year", allMember: null, expanded: [], members: [], memberCaptions: [] },
    ]);
  });

  it("has no slicers key for a report without slicers", () => {
    expect("slicers" in parseLayout({ rows: [] })).toBe(false);
    expect("slicers" in parseLayout({ slicers: "nope" })).toBe(false);
    expect("slicers" in parseLayout({ slicers: [{ caption: "no hierarchy" }, 3, null] })).toBe(false);
  });

  it("drops a slicer whose field is in an area, a duplicate, a named set, and more than 16", () => {
    const many = Array.from({ length: 20 }, (_, i) => ({ hierarchy: `[D].[H${i}]`, caption: `H${i}` }));
    const layout = parseLayout({
      rows: [{ hierarchy: CITY, caption: "City" }],
      slicers: [
        { hierarchy: CITY, caption: "City" },
        { hierarchy: "[Set].[H]", caption: "Set", set: "[Some Set]" },
        { hierarchy: "[D].[H0]", caption: "H0" },
        ...many,
      ],
    });
    expect(layout.rows.map((f) => f.hierarchy)).toEqual([CITY]);
    expect(layout.slicers).toHaveLength(MAX_SLICERS);
    expect(layout.slicers?.map((s) => s.hierarchy).slice(0, 3)).toEqual(["[D].[H0]", "[D].[H1]", "[D].[H2]"]);
  });
});

describe("the Sample cube filters by slicers like page filters", () => {
  const api = createMockApi({ delayMs: 0 });
  const total = (grid: Grid) => grid.values[grid.values.length - 1][0];
  const base = run([
    { type: "addHierarchy", hierarchy: service },
    { type: "addMeasure", measure: rides },
  ]);
  const query = (layout: Layout) => api.query("demo", { catalog: "Sample", cube: "Rides", layout });

  it("gives the numbers a page filter gives, without a filter row", async () => {
    const asFilter = run(
      [
        { type: "addHierarchy", hierarchy: city, area: "filters" },
        { type: "setMembers", hierarchy: CITY, members: ["[Geography].[City].&[1]"], captions: ["Tehran"] },
      ],
      base,
    );
    const asSlicer = layoutReducer(asFilter, { type: "addSlicer", hierarchy: city });
    expect(asSlicer.filters).toEqual([]);
    const [byFilter, bySlicer, unfiltered] = await Promise.all([query(asFilter), query(asSlicer), query(base)]);
    expect(bySlicer.filterRows).toEqual([]);
    expect(bySlicer.values).toEqual(byFilter.values);
    expect(total(bySlicer)).toBeLessThan(total(unfiltered) as number);
  });

  it("takes several items, and none means all", async () => {
    const members = (names: string[]) =>
      run([{ type: "addSlicer", hierarchy: city }, { type: "setMembers", hierarchy: CITY, members: names, captions: names }], base);
    const [one, two, none, unfiltered] = await Promise.all([
      query(members(["[Geography].[City].&[1]"])),
      query(members(["[Geography].[City].&[1]", "[Geography].[City].&[2]"])),
      query(members([])),
      query(base),
    ]);
    expect(total(two)).toBeGreaterThan(total(one) as number);
    expect(none.values).toEqual(unfiltered.values);
  });

  it("lists a slicer's items with the members route, the top level", async () => {
    const response = await api.members("demo", { catalog: "Sample", cube: "Rides", hierarchy: CITY });
    expect(response.members.length).toBeGreaterThan(1);
  });
});
