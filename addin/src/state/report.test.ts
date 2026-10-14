import { describe, expect, it } from "vitest";
import type { ReportDef } from "../sheet/host";
import { emptyLayout, isEmptyLayout } from "./layout";
import {
  currentReducer,
  decidePlace,
  placeLabel,
  newReportDef,
  newReportId,
  parseReportDef,
  serializeReportDef,
  type Current,
} from "./report";

const rides = { uniqueName: "[Measures].[Rides]", caption: "Rides" };
const calendar = {
  uniqueName: "[Date].[Calendar]",
  caption: "Calendar",
  allMember: "[Date].[Calendar].[All]",
};

function open(def: ReportDef = newReportDef("demo", "Sample", "Rides"), persisted = false): Current {
  return currentReducer(null, { type: "open", def, persisted }) as Current;
}

describe("report ids", () => {
  it("are ten letters or digits, safe inside a workbook name", () => {
    for (let i = 0; i < 50; i++) expect(newReportId()).toMatch(/^[a-z0-9]{10}$/);
  });
  it("differ from one another", () => {
    expect(newReportId()).not.toBe(newReportId());
  });
});

describe("current report reducer", () => {
  it("opens a report at revision 0 with no pending query", () => {
    const state = open();
    expect(state.revision).toBe(0);
    expect(state.persisted).toBe(false);
    expect(state.anchor).toBeNull();
  });

  it("bumps the revision on a layout change, and not on a no-op", () => {
    const start = open();
    const once = currentReducer(start, { type: "layout", action: { type: "addMeasure", measure: rides } })!;
    expect(once.revision).toBe(1);
    expect(once.urgent).toBe(false);
    expect(once.def.layout.values).toHaveLength(1);
    const twice = currentReducer(once, { type: "layout", action: { type: "addMeasure", measure: rides } });
    expect(twice).toBe(once);
  });

  it("marks drill actions as urgent", () => {
    const withRow = currentReducer(open(), {
      type: "layout",
      action: { type: "addHierarchy", hierarchy: calendar },
    })!;
    const expanded = currentReducer(withRow, {
      type: "layout",
      action: { type: "expand", hierarchy: calendar.uniqueName, member: "[Date].[Calendar].[2025]" },
      urgent: true,
    })!;
    expect(expanded.urgent).toBe(true);
    expect(expanded.revision).toBe(2);
  });

  it("drops the layout when the cube changes, and the cube too when the database changes", () => {
    const filled = currentReducer(open(), {
      type: "layout",
      action: { type: "addMeasure", measure: rides },
    })!;
    const otherCube = currentReducer(filled, { type: "setCube", cube: "Other" })!;
    expect(otherCube.def.cube).toBe("Other");
    expect(isEmptyLayout(otherCube.def.layout)).toBe(true);
    expect(otherCube.revision).toBe(filled.revision + 1);

    const otherDb = currentReducer(filled, { type: "setCatalog", catalog: "Finance" })!;
    expect(otherDb.def).toMatchObject({ catalog: "Finance", cube: "" });
    expect(isEmptyLayout(otherDb.def.layout)).toBe(true);
    expect(currentReducer(filled, { type: "setCube", cube: "Rides" })).toBe(filled);
  });

  it("rebinding keeps database, cube and layout", () => {
    const filled = currentReducer(open(), {
      type: "layout",
      action: { type: "addMeasure", measure: rides },
    })!;
    const rebound = currentReducer(filled, { type: "rebind", connectionId: "other" })!;
    expect(rebound.def).toMatchObject({ connectionId: "other", catalog: "Sample", cube: "Rides" });
    expect(rebound.def.layout).toBe(filled.def.layout);
    expect(rebound.revision).toBe(filled.revision + 1);
  });

  it("fills in a missing database and cube without asking for a query", () => {
    const start = open(newReportDef("demo"));
    const filled = currentReducer(start, { type: "defaults", catalog: "Sample", cube: "Rides" })!;
    expect(filled.def).toMatchObject({ catalog: "Sample", cube: "Rides" });
    expect(filled.revision).toBe(0);
    // Never overrides a choice already made.
    const again = currentReducer(filled, { type: "defaults", catalog: "X", cube: "Y" });
    expect(again).toBe(filled);
  });

  it("pins an anchor, marks saved, touches and closes", () => {
    const anchor = { sheetId: "s", sheetName: "Sheet1", row: 2, col: 3 };
    const pinned = currentReducer(open(), { type: "pin", anchor })!;
    expect(pinned.anchor).toEqual(anchor);
    const saved = currentReducer(pinned, { type: "saved" })!;
    expect(saved.persisted).toBe(true);
    expect(currentReducer(saved, { type: "saved" })).toBe(saved);
    // A result for some other report does not mark this one as saved.
    expect(currentReducer(pinned, { type: "saved", id: "someone-else" })).toBe(pinned);
    expect(currentReducer(pinned, { type: "saved", id: pinned.def.id })!.persisted).toBe(true);
    const touched = currentReducer(saved, { type: "touch" })!;
    expect(touched.revision).toBe(1);
    expect(touched.def).toBe(saved.def);
    expect(currentReducer(touched, { type: "close" })).toBeNull();
    expect(currentReducer(null, { type: "touch" })).toBeNull();
  });
});

describe("reading a stored definition", () => {
  it("round-trips through JSON", () => {
    const def = newReportDef("demo", "Sample", "Rides");
    def.layout.values.push({ measure: rides.uniqueName, caption: "Rides" });
    def.layout.rows.push({
      hierarchy: calendar.uniqueName,
      caption: "Calendar",
      allMember: calendar.allMember,
      expanded: ["[Date].[Calendar].[2025]"],
      members: ["[Date].[Calendar].[2025]"],
      memberCaptions: ["2025"],
    });
    expect(parseReportDef(serializeReportDef(def), def.id)).toEqual(def);
  });

  it("returns null for junk instead of throwing", () => {
    expect(parseReportDef("not json")).toBeNull();
    expect(parseReportDef("null")).toBeNull();
    expect(parseReportDef("[1,2]")).toBeNull();
    expect(parseReportDef(42)).toBeNull();
    expect(parseReportDef(JSON.stringify({ connectionId: "demo" }))).toBeNull(); // no id
    expect(parseReportDef(JSON.stringify({ id: "bad id!" }))).toBeNull();
  });

  it("rejects a definition stored under another report's key", () => {
    expect(parseReportDef(JSON.stringify({ id: "abc" }), "xyz")).toBeNull();
  });

  it("takes the id from the key when the body has none", () => {
    expect(parseReportDef(JSON.stringify({ connectionId: "demo" }), "abc123")?.id).toBe("abc123");
  });

  it("repairs a partial definition", () => {
    const def = parseReportDef(
      JSON.stringify({
        id: "abc",
        connectionId: "gone",
        layout: {
          rows: [
            { hierarchy: "[A].[B]", members: ["m1", "m2"], memberCaptions: ["One"] },
            { hierarchy: 5 },
            "junk",
          ],
          columns: [{ hierarchy: "[A].[B]", caption: "Duplicate" }],
          values: [{ measure: "[Measures].[X]" }, { measure: "[Measures].[X]" }, {}],
          options: { subtotals: true, nonEmpty: "yes" },
        },
      }),
    );
    expect(def).toEqual({
      id: "abc",
      connectionId: "gone",
      catalog: "",
      cube: "",
      layout: {
        rows: [
          {
            hierarchy: "[A].[B]",
            caption: "[A].[B]",
            allMember: null,
            expanded: [],
            members: ["m1", "m2"],
            memberCaptions: ["One", "m2"],
          },
        ],
        columns: [],
        filters: [],
        values: [{ measure: "[Measures].[X]", caption: "[Measures].[X]" }],
        options: { grandTotals: true, subtotals: true, nonEmpty: true },
      },
    });
  });

  it("gives an empty layout when there is none", () => {
    expect(parseReportDef(JSON.stringify({ id: "abc" }))?.layout).toEqual(emptyLayout());
  });
});

describe("what the place button does", () => {
  const extent = { sheetId: "s1", sheetName: "Sheet1", row: 0, col: 0, rows: 4, cols: 2 };
  const outside = { report: null };
  const inside = (id: string) => ({ report: { id, row: 1, col: 0, rows: 4, cols: 2 } });

  it("places the report being edited when it has no cells", () => {
    expect(decidePlace("r1", null, outside)).toEqual({ kind: "place" });
  });

  it("starts another report when the one being edited is on a sheet", () => {
    expect(decidePlace("r1", extent, outside)).toEqual({ kind: "new" });
  });

  it("goes by the location it is handed, so a report whose sheet was deleted is placed, not replaced", () => {
    // What the pane remembered said "on Sheet1"; the workbook, read just now, says no cells.
    const remembered = extent;
    const live = null;
    expect(decidePlace("r1", remembered, outside).kind).toBe("new");
    expect(decidePlace("r1", live, outside).kind).toBe("place");
  });

  it("does nothing inside a report, whichever it is", () => {
    expect(decidePlace("r1", extent, inside("r1"))).toEqual({ kind: "blocked", inside: "this" });
    expect(decidePlace("r1", extent, inside("r2"))).toEqual({ kind: "blocked", inside: "other" });
    expect(decidePlace("r1", null, inside("r2"))).toEqual({ kind: "blocked", inside: "other" });
  });

  it("words the button for the same two cases", () => {
    expect(placeLabel(null)).toBe("Place here");
    expect(placeLabel(extent)).toBe("New report here");
  });
});

describe("reading sorts, filters, show-as and formats back", () => {
  const base = {
    hierarchy: "[A].[B]",
    caption: "B",
    allMember: "[A].[B].[All]",
    expanded: [],
    members: [],
    memberCaptions: [],
  };

  it("round-trips every new setting", () => {
    const def = newReportDef("demo", "Sample", "Rides");
    def.layout.rows.push({
      ...base,
      sort: { by: "measure", measure: "[Measures].[X]", order: "desc" },
      labelFilter: { operator: "between", value: "a", value2: "m" },
      valueFilter: { operator: "top", measure: "[Measures].[X]", value: 10, topBy: "percent" },
    });
    def.layout.values.push({ measure: "[Measures].[X]", caption: "X", showAs: "percentOfColumnTotal", numberFormat: "0.0%" });
    expect(parseReportDef(serializeReportDef(def), def.id)).toEqual(def);
  });

  it("drops settings it cannot read instead of guessing", () => {
    const def = parseReportDef(
      JSON.stringify({
        id: "abc",
        layout: {
          rows: [
            {
              ...base,
              sort: { by: "measure", order: "asc" }, // no measure
              labelFilter: { operator: "between", value: "a" }, // no upper bound
              valueFilter: { operator: "greaterThan", measure: "[M]", value: "5" }, // not a number
            },
            { ...base, hierarchy: "[A].[C]", sort: { by: "caption", order: "sideways" }, labelFilter: { operator: "like", value: "x" } },
          ],
          values: [
            { measure: "[M]", showAs: "percentOfEverything", numberFormat: "" },
            { measure: "[N]", numberFormat: "x".repeat(101) },
          ],
        },
      }),
    )!;
    for (const field of def.layout.rows) {
      expect("sort" in field).toBe(false);
      expect("labelFilter" in field).toBe(false);
      expect("valueFilter" in field).toBe(false);
    }
    expect(def.layout.values).toEqual([
      { measure: "[M]", caption: "[M]" },
      { measure: "[N]", caption: "[N]" },
    ]);
  });

  it("gives a top filter its default unit and drops a stray upper bound", () => {
    const def = parseReportDef(
      JSON.stringify({
        id: "abc",
        layout: { rows: [{ ...base, valueFilter: { operator: "bottom", measure: "[M]", value: 3, value2: 9 } }] },
      }),
    )!;
    expect(def.layout.rows[0].valueFilter).toEqual({ operator: "bottom", measure: "[M]", value: 3, topBy: "items" });
  });

  it("keeps the same measure twice when it is shown two ways, but not an exact repeat", () => {
    const def = parseReportDef(
      JSON.stringify({
        id: "abc",
        layout: {
          values: [
            { measure: "[M]", caption: "M" },
            { measure: "[M]", caption: "M", showAs: "percentOfGrandTotal" },
            { measure: "[M]", caption: "M" },
          ],
        },
      }),
    )!;
    expect(def.layout.values).toEqual([
      { measure: "[M]", caption: "M" },
      { measure: "[M]", caption: "M", showAs: "percentOfGrandTotal" },
    ]);
  });
});

describe("reading the Windows options back", () => {
  it("round-trips grand totals per axis, repeat labels, empty-cell text and a base field", () => {
    const def = newReportDef("demo", "Sample", "Rides");
    def.layout.options = {
      grandTotals: true,
      grandTotalRow: false,
      grandTotalColumn: true,
      subtotals: false,
      nonEmpty: true,
      repeatLabels: true,
      emptyCellText: "-",
    };
    def.layout.rows.push({ hierarchy: "[A].[B]", caption: "B", allMember: null, expanded: [], members: [], memberCaptions: [] });
    def.layout.values.push({ measure: "[M]", caption: "M", showAs: "differenceFrom", baseField: "[A].[B]", baseItem: "previous" });
    def.layout.values.push({ measure: "[M]", caption: "M", showAs: "rankAscending", baseField: "[A].[B]" });
    expect(parseReportDef(serializeReportDef(def), def.id)).toEqual(def);
  });

  it("reads an incomplete show-as as No Calculation, and drops a base it does not need", () => {
    const def = parseReportDef(
      JSON.stringify({
        id: "abc",
        layout: {
          values: [
            { measure: "[A]", showAs: "percentOf", baseField: "[X].[Y]" }, // no base item
            { measure: "[B]", showAs: "runningTotalIn" }, // no base field
            { measure: "[C]", showAs: "index", baseField: "[X].[Y]", baseItem: "next" },
            { measure: "[D]", showAs: "percentOf", baseField: "[X].[Y]", baseItem: "[X].[Y].&[1]" },
          ],
          options: { grandTotalRow: "yes", repeatLabels: "true", emptyCellText: "x".repeat(25) },
        },
      }),
    )!;
    expect(def.layout.values).toEqual([
      { measure: "[A]", caption: "[A]" },
      { measure: "[B]", caption: "[B]" },
      { measure: "[C]", caption: "[C]", showAs: "index" },
      { measure: "[D]", caption: "[D]", showAs: "percentOf", baseField: "[X].[Y]", baseItem: "[X].[Y].&[1]" },
    ]);
    expect(def.layout.options).toEqual({ grandTotals: true, subtotals: false, nonEmpty: true, emptyCellText: "x".repeat(20) });
  });
});
