import { describe, expect, it } from "vitest";
import type { ReportDef } from "../sheet/host";
import { emptyLayout, layoutReducer } from "./layout";
import { newReportDef } from "./report";
import { definitionKey, syncReports } from "./sync";

function def(id: string): ReportDef {
  return { ...newReportDef("demo", "Sample", "Rides"), id };
}
/** The same report with one setting changed, as a command or the pane would. */
function edited(base: ReportDef, form: "compact" | "outline" | "tabular" = "outline"): ReportDef {
  return { ...base, layout: layoutReducer(base.layout, { type: "setForm", form }) };
}

describe("definitionKey", () => {
  it("is the same for equal definitions whatever their key order", () => {
    const a = def("r1");
    const reordered: ReportDef = {
      layout: { options: { ...a.layout.options }, filters: [], values: [], columns: [], rows: [] },
      cube: a.cube,
      catalog: a.catalog,
      connectionId: a.connectionId,
      id: a.id,
    };
    expect(definitionKey(reordered)).toBe(definitionKey(a));
    expect(definitionKey(edited(a))).not.toBe(definitionKey(a));
  });
});

describe("syncReports", () => {
  const a = def("a1");
  const b = def("b1");

  it("changes nothing when the workbook holds what the pane knows", () => {
    const known = [a, b];
    const result = syncReports({ loaded: [b, a], known, current: a, pending: false });
    expect(result).toEqual({ reports: known, changed: [], reopen: null });
    expect(result.reports).toBe(known);
  });

  it("takes a definition changed outside the pane, and reopens the report being edited", () => {
    const expanded = edited(a);
    const result = syncReports({ loaded: [expanded, b], known: [a, b], current: a, pending: false });
    expect(result.reports).toEqual([expanded, b]);
    expect(result.changed).toEqual([a.id]);
    expect(result.reopen).toEqual(expanded);
  });

  it("updates another report without reopening the one being edited", () => {
    const changedB = edited(b);
    const result = syncReports({ loaded: [a, changedB], known: [a, b], current: a, pending: false });
    expect(result.reports).toEqual([a, changedB]);
    expect(result.changed).toEqual([b.id]);
    expect(result.reopen).toBeNull();
  });

  it("does not reopen while the pane has a change of its own waiting", () => {
    const expanded = edited(a);
    const mine = edited(a, "compact");
    const result = syncReports({ loaded: [expanded], known: [a], current: mine, pending: true });
    expect(result.reopen).toBeNull();
    // The list still holds what the workbook holds; the pane's own run will save over it.
    expect(result.reports).toEqual([expanded]);
    expect(result.changed).toEqual([a.id]);
  });

  it("does not count the pane's own save, which its list has not caught up with", () => {
    const mine = edited(a);
    const result = syncReports({ loaded: [mine], known: [a], current: mine, pending: false });
    expect(result.reports).toEqual([mine]);
    expect(result.changed).toEqual([]);
    expect(result.reopen).toBeNull();
  });

  it("keeps reports the read did not see, and never adds one the pane does not know", () => {
    const stranger = def("zz9");
    const result = syncReports({ loaded: [stranger], known: [a, b], current: null, pending: false });
    expect(result.reports).toEqual([a, b]);
    expect(result.changed).toEqual([]);
  });

  it("does not reopen a new report that is not saved yet", () => {
    const expanded = edited(a);
    const result = syncReports({ loaded: [expanded], known: [a], current: null, pending: false });
    expect(result.changed).toEqual([a.id]);
    expect(result.reopen).toBeNull();
  });

  it("starts from an empty layout too", () => {
    const blank = { ...def("c1"), layout: emptyLayout() };
    const result = syncReports({ loaded: [blank], known: [blank], current: blank, pending: false });
    expect(result.changed).toEqual([]);
  });
});
