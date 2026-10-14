// A report's charts in its stored definition, and in the report being edited.

import { describe, expect, it } from "vitest";
import type { ReportChart } from "../sheet/charts";
import { currentReducer, newReportDef, parseReportDef, serializeReportDef, type Current } from "./report";

const pie: ReportChart = { id: "c1", type: "Pie", sheetId: "{S1}", title: "Rides" };
const column: ReportChart = { id: "c2", type: "ColumnClustered", sheetId: "{S2}" };

describe("charts in a stored definition", () => {
  it("round-trip through JSON", () => {
    const def = { ...newReportDef("demo", "Sample", "Rides"), charts: [pie, column] };
    expect(parseReportDef(serializeReportDef(def), def.id)).toEqual(def);
  });

  it("are left out of the JSON when there are none", () => {
    const def = newReportDef("demo", "Sample", "Rides");
    expect(JSON.parse(serializeReportDef(def))).not.toHaveProperty("charts");
    expect(JSON.parse(serializeReportDef({ ...def, charts: [] }))).not.toHaveProperty("charts");
    expect(parseReportDef(serializeReportDef({ ...def, charts: [] }), def.id)).toEqual(def);
  });

  it("read back without them from definitions written before charts existed", () => {
    const def = parseReportDef(JSON.stringify({ id: "abc", connectionId: "demo" }));
    expect(def).not.toBeNull();
    expect(def).not.toHaveProperty("charts");
  });

  it("drop what cannot be read and keep the rest of the definition", () => {
    const stored = JSON.stringify({
      id: "abc",
      connectionId: "demo",
      catalog: "Sample",
      cube: "Rides",
      charts: [pie, { id: "x", type: "Radar", sheetId: "s" }, "nonsense", { ...column, id: "c1" }],
    });
    const def = parseReportDef(stored, "abc");
    expect(def?.charts).toEqual([pie]);
    expect(def?.cube).toBe("Rides");
    expect(parseReportDef(JSON.stringify({ id: "abc", charts: "all of them" }))).not.toHaveProperty("charts");
  });
});

describe("charts of the report being edited", () => {
  function open(): Current {
    return currentReducer(null, { type: "open", def: newReportDef("demo", "Sample", "Rides"), persisted: true })!;
  }

  it("change without asking for a new query", () => {
    const state = open();
    const next = currentReducer(state, { type: "charts", id: state.def.id, charts: [pie] })!;
    expect(next.def.charts).toEqual([pie]);
    expect(next.revision).toBe(state.revision);
    expect(next.def.layout).toBe(state.def.layout);
  });

  it("ignore the same list, and lists for another report", () => {
    const state = currentReducer(open(), { type: "charts", id: open().def.id, charts: [] })!;
    const withPie = currentReducer(state, { type: "charts", id: state.def.id, charts: [pie] })!;
    expect(currentReducer(withPie, { type: "charts", id: state.def.id, charts: [{ ...pie }] })).toBe(withPie);
    expect(currentReducer(withPie, { type: "charts", id: "other", charts: [] })).toBe(withPie);
    expect(currentReducer(state, { type: "charts", id: state.def.id, charts: [] })).toBe(state);
  });

  it("take the field away when the last chart goes", () => {
    const state = open();
    const withPie = currentReducer(state, { type: "charts", id: state.def.id, charts: [pie] })!;
    const none = currentReducer(withPie, { type: "charts", id: state.def.id, charts: [] })!;
    expect(none.def).not.toHaveProperty("charts");
  });
});
