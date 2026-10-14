import { beforeEach, describe, expect, it } from "vitest";
import type { Api } from "../api";
import { ApiFailure, HelperDown } from "../apiErrors";
import { createMockApi } from "../dev/mockApi";
import type { Extent, ReportDef } from "../sheet/host";
import { FILTER_FORMAT, PLACEHOLDER_TEXT, guessFieldCell } from "../sheet/layout";
import { PreviewHost } from "../sheet/preview";
import type { Fields, Layout } from "../types";
import { emptyLayout, layoutReducer, type LayoutAction } from "./layout";
import { newReportDef } from "./report";
import { ReportRunner, type RunOutcome } from "./runner";

const CATALOG = "Sample";
const CUBE = "Rides";

let api: Api;
let host: PreviewHost;
let fields: Fields;
let asked: Extent[];
let answer: boolean;
let cancels: number;
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
function layoutOf(actions: LayoutAction[], start: Layout = emptyLayout()): Layout {
  return actions.reduce(layoutReducer, start);
}
function report(layout: Layout, id?: string): ReportDef {
  const def = newReportDef("demo", CATALOG, CUBE);
  return { ...def, id: id ?? def.id, layout };
}
function basic(): Layout {
  return layoutOf([
    { type: "addMeasure", measure: measure("Rides") },
    { type: "addHierarchy", hierarchy: hierarchy("Calendar") },
  ]);
}
function expectWritten(outcome: RunOutcome) {
  if (outcome.kind !== "written") throw new Error("expected written, got " + JSON.stringify(outcome));
  return outcome;
}
function cellText(row: number, col: number) {
  const cell = host.snapshot().cells.get(row + "," + col);
  return cell ? cell.value : undefined;
}

beforeEach(async () => {
  const mock = createMockApi({ delayMs: 0 });
  cancels = 0;
  api = {
    ...mock,
    cancel: async (id) => {
      cancels++;
      await mock.cancel(id);
    },
  };
  host = new PreviewHost({ persist: false });
  fields = await api.fields("demo", CATALOG, CUBE);
  asked = [];
  answer = true;
  runner = new ReportRunner({
    api,
    host,
    confirmOverwrite: async (_def, target) => {
      asked.push(target);
      return answer;
    },
  });
});

describe("first run of a new report", () => {
  it("writes at the active cell, names the block and saves the definition", async () => {
    host.select(3, 2);
    const def = report(basic());
    const outcome = expectWritten(await runner.run(def, { anchor: null, useSelection: true }));

    // Header row, 2025, 2026, Grand Total: 4 rows; Calendar + Rides: 2 columns.
    expect(outcome.extent).toMatchObject({ row: 3, col: 2, rows: 4, cols: 2 });
    expect(await host.getReportExtent(def.id)).toMatchObject({ row: 3, col: 2, rows: 4, cols: 2 });
    expect(cellText(3, 2)).toBe("Calendar");
    expect(cellText(3, 3)).toBe("Rides");
    expect(cellText(4, 2)).toBe("2025");
    expect(cellText(6, 2)).toBe("Grand Total");
    expect(typeof cellText(6, 3)).toBe("number");
    expect(await host.loadReports()).toEqual([def]);
    expect(asked).toEqual([]);
  });

  it("uses the pinned anchor rather than the active cell", async () => {
    host.select(0, 0);
    const def = report(basic());
    const anchor = { sheetId: "preview-sheet-1", sheetName: "Sheet1", row: 10, col: 5 };
    const outcome = expectWritten(await runner.run(def, { anchor, useSelection: true }));
    expect(outcome.extent).toMatchObject({ row: 10, col: 5 });
    expect(cellText(0, 0)).toBeUndefined();
  });

  it("writes nothing and saves nothing for an empty layout that has no place yet", async () => {
    // Not even at the selected cell: a placeholder only goes where the report was put.
    host.select(2, 2);
    const def = report(emptyLayout());
    expect(await runner.run(def, { anchor: null, useSelection: true })).toEqual({
      kind: "empty",
      extent: null,
    });
    expect(host.snapshot().cells.size).toBe(0);
    expect(await host.loadReports()).toEqual([]);
    expect(await host.getReportExtent(def.id)).toBeNull();
  });

  it("does not count a filter alone as something to write", async () => {
    const def = report(
      layoutOf([{ type: "addHierarchy", hierarchy: hierarchy("City"), area: "filters" }]),
    );
    expect((await runner.run(def, { anchor: null, useSelection: true })).kind).toBe("empty");
    expect(host.snapshot().cells.size).toBe(0);
  });

  it("asks for a place when there is none to use", async () => {
    const def = report(basic());
    expect(await runner.run(def, { anchor: null, useSelection: false })).toEqual({
      kind: "needs-anchor",
    });

    // The active cell is inside another report: not a place for a new one.
    const other = report(basic());
    host.select(1, 1);
    expectWritten(await runner.run(other, { anchor: null, useSelection: true }));
    host.select(2, 1);
    expect(await runner.run(def, { anchor: null, useSelection: true })).toEqual({
      kind: "needs-anchor",
    });
  });
});

describe("refresh", () => {
  it("rewrites in place and clears what the smaller result no longer covers", async () => {
    host.select(1, 1);
    const calendar = hierarchy("Calendar");
    const year2025 = "[Date].[Calendar].[Year].&[2025]";
    const wide = layoutOf([{ type: "expand", hierarchy: calendar.uniqueName, member: year2025 }], basic());
    const def = report(wide);
    const first = expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    expect(first.extent.rows).toBe(4 + 4); // four quarters under 2025
    expect(cellText(3, 1)).toBe("Q1 2025");

    const second = expectWritten(
      await runner.run({ ...def, layout: basic() }, { anchor: null, useSelection: false }),
    );
    expect(second.extent).toMatchObject({ row: 1, col: 1, rows: 4, cols: 2 });
    expect(cellText(4, 1)).toBe("Grand Total");
    expect(cellText(5, 1)).toBeUndefined();
    expect(cellText(8, 1)).toBeUndefined();
    expect(asked).toEqual([]);
  });

  it("follows the block when rows are inserted above it", async () => {
    host.select(2, 0);
    const def = report(basic());
    expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    host.insertRows(0, 5);
    host.select(0, 0);
    const again = expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    expect(again.extent).toMatchObject({ row: 7, col: 0 });
    expect(cellText(7, 0)).toBe("Calendar");
    expect(cellText(0, 0)).toBeUndefined();
  });

  it("leaves a placeholder at its place when every field is removed", async () => {
    host.select(4, 4);
    const def = report(basic());
    expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    const emptied = { ...def, layout: emptyLayout() };
    const outcome = await runner.run(emptied, { anchor: null, useSelection: false });
    expect(outcome).toMatchObject({ kind: "empty", extent: { row: 4, col: 4, rows: 1, cols: 1 } });
    // The old block is gone; one quiet cell marks where the report is.
    expect(host.snapshot().cells.size).toBe(1);
    expect(cellText(4, 4)).toBe(PLACEHOLDER_TEXT);
    expect(host.snapshot().cells.get("4,4")?.muted).toBe(true);
    expect(await host.getReportExtent(def.id)).toMatchObject({ row: 4, col: 4, rows: 1, cols: 1 });
    expect(await host.loadReports()).toEqual([emptied]);
    expect(runner.runtimes.has(def.id)).toBe(false);

    // Adding a field again writes at the same place.
    const back = expectWritten(await runner.run(def, { anchor: null, useSelection: false }));
    expect(back.extent).toMatchObject({ row: 4, col: 4, rows: 4, cols: 2 });
  });
});

describe("overwriting other cells", () => {
  it("asks before covering a non-empty cell, and leaves everything alone on No", async () => {
    host.select(0, 0);
    host.type(2, 1, "mine");
    answer = false;
    const def = report(basic());
    expect(await runner.run(def, { anchor: null, useSelection: true })).toEqual({ kind: "declined" });
    expect(asked).toHaveLength(1);
    expect(asked[0]).toMatchObject({ row: 0, col: 0, rows: 4, cols: 2 });
    expect(cellText(2, 1)).toBe("mine");
    expect(cellText(0, 0)).toBeUndefined();
    expect(await host.loadReports()).toEqual([]);
  });

  it("writes over it on Yes", async () => {
    host.select(0, 0);
    host.type(2, 1, "mine");
    const def = report(basic());
    expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    expect(asked).toHaveLength(1);
    expect(typeof cellText(2, 1)).toBe("number");
  });

  it("does not ask about cells next to the block", async () => {
    host.select(0, 0);
    host.type(0, 2, "beside");
    host.type(4, 0, "below");
    expectWritten(await runner.run(report(basic()), { anchor: null, useSelection: true }));
    expect(asked).toEqual([]);
    expect(cellText(0, 2)).toBe("beside");
    expect(cellText(4, 0)).toBe("below");
  });

  it("asks when a refresh grows onto cells outside the old extent, not for its own cells", async () => {
    host.select(0, 0);
    const def = report(basic());
    expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    expectWritten(await runner.run(def, { anchor: null, useSelection: false }));
    expect(asked).toEqual([]);

    host.type(5, 0, "note");
    const grown = layoutOf(
      [{ type: "expand", hierarchy: hierarchy("Calendar").uniqueName, member: "[Date].[Calendar].[Year].&[2025]" }],
      def.layout,
    );
    answer = false;
    expect((await runner.run({ ...def, layout: grown }, { anchor: null, useSelection: false })).kind).toBe(
      "declined",
    );
    expect(asked).toHaveLength(1);
    expect(cellText(5, 0)).toBe("note");
    // The old result is still there, untouched.
    expect(cellText(3, 0)).toBe("Grand Total");
  });
});

describe("which request wins", () => {
  it("drops the older response when a newer run starts, and stops the older query", async () => {
    host.select(0, 0);
    const def = report(basic());
    const newer = { ...def, layout: layoutOf([{ type: "addMeasure", measure: measure("Fare") }], def.layout) };
    const first = runner.run(def, { anchor: null, useSelection: true });
    const second = runner.run(newer, { anchor: null, useSelection: true });
    expect(await first).toEqual({ kind: "superseded" });
    const written = expectWritten(await second);
    expect(written.extent.cols).toBe(3);
    expect(cancels).toBe(1);
    expect(cellText(0, 2)).toBe("Fare");
    expect(await host.loadReports()).toEqual([newer]);
  });

  it("keeps only the last of several quick runs", async () => {
    host.select(0, 0);
    const def = report(basic());
    const outcomes = await Promise.all([
      runner.run(def, { anchor: null, useSelection: true }),
      runner.run(def, { anchor: null, useSelection: true }),
      runner.run(def, { anchor: null, useSelection: true }),
    ]);
    expect(outcomes.map((o) => o.kind)).toEqual(["superseded", "superseded", "written"]);
  });

  it("cancel stops the query, writes nothing and is not an error", async () => {
    host.select(0, 0);
    const pending = runner.run(report(basic()), { anchor: null, useSelection: true });
    await Promise.resolve();
    expect(runner.busy).toBe(true);
    runner.cancel();
    expect(await pending).toEqual({ kind: "cancelled" });
    expect(cancels).toBe(1);
    expect(host.snapshot().cells.size).toBe(0);
    expect(runner.busy).toBe(false);
  });

  it("runs normally again after a cancel", async () => {
    host.select(0, 0);
    const def = report(basic());
    const pending = runner.run(def, { anchor: null, useSelection: true });
    runner.cancel();
    await pending;
    expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
  });

  it("asks once more when the helper reports a cancel nobody asked for", async () => {
    let calls = 0;
    const flaky: Api = {
      ...api,
      query: async (id, request, signal) => {
        calls++;
        if (calls === 1) throw new ApiFailure({ category: "cancelled", message: "Cancelled." }, 499);
        return api.query(id, request, signal);
      },
    };
    const r = new ReportRunner({ api: flaky, host, confirmOverwrite: async () => true });
    host.select(0, 0);
    expectWritten(await r.run(report(basic()), { anchor: null, useSelection: true }));
    expect(calls).toBe(2);
  });

  it("takes the overwrite question down when the run is cancelled while it is open", async () => {
    host.select(0, 0);
    host.type(1, 0, "mine");
    let aborted = false;
    const r = new ReportRunner({
      api,
      host,
      confirmOverwrite: (_def, _target, signal) =>
        new Promise<boolean>(() => {
          signal.addEventListener("abort", () => (aborted = true));
          setTimeout(() => r.cancel(), 0);
        }),
    });
    expect(await r.run(report(basic()), { anchor: null, useSelection: true })).toEqual({
      kind: "cancelled",
    });
    expect(aborted).toBe(true);
    expect(cellText(1, 0)).toBe("mine");
  });
});

describe("errors", () => {
  it("reports the helper's own message for an API error", async () => {
    const def = { ...report(basic()), cube: "Nope" };
    const outcome = await runner.run(def, { anchor: null, useSelection: true });
    expect(outcome.kind).toBe("error");
    if (outcome.kind === "error") {
      expect(outcome.error.message).toBe('There is no cube named "Nope".');
      expect(outcome.error.category).toBe("not-found");
      expect(outcome.error.helperDown).toBeUndefined();
    }
    expect(host.snapshot().cells.size).toBe(0);
  });

  it("flags a helper that cannot be reached", async () => {
    const down: Api = {
      ...api,
      query: async () => {
        throw new HelperDown("Failed to fetch");
      },
    };
    const r = new ReportRunner({ api: down, host, confirmOverwrite: async () => true });
    const outcome = await r.run(report(basic()), { anchor: null, useSelection: true });
    expect(outcome).toMatchObject({
      kind: "error",
      error: { message: "The helper is not running.", helperDown: true },
    });
  });

  it("refuses a block that would run off the sheet", async () => {
    const def = report(basic());
    const anchor = { sheetId: "preview-sheet-1", sheetName: "Sheet1", row: 1048574, col: 0 };
    const outcome = await runner.run(def, { anchor, useSelection: false });
    expect(outcome.kind).toBe("error");
    expect(host.snapshot().cells.size).toBe(0);
  });
});

describe("drill", () => {
  const year2025 = "[Date].[Calendar].[Year].&[2025]";
  const q1 = "[Date].[Calendar].[Quarter].&[2025]&[1]";

  it("remembers the member behind each header cell of the last result", async () => {
    host.select(0, 0);
    const def = report(basic());
    expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    const hit = runner.memberAt(def.id, 1, 0);
    expect(hit).toMatchObject({
      axis: "rows",
      member: { uniqueName: year2025, caption: "2025", hasChildren: true, expanded: false },
    });
    expect(runner.memberAt(def.id, 0, 0)).toBeNull(); // corner caption
    expect(runner.memberAt(def.id, 3, 0)).toBeNull(); // Grand Total
    expect(runner.memberAt("unknown", 1, 0)).toBeNull();
  });

  it("expands under the member and collapses with its expanded descendants", async () => {
    host.select(0, 0);
    const calendar = hierarchy("Calendar").uniqueName;
    let def = report(basic());
    expectWritten(await runner.run(def, { anchor: null, useSelection: true }));

    def = { ...def, layout: layoutOf([
      { type: "expand", hierarchy: calendar, member: year2025 },
      { type: "expand", hierarchy: calendar, member: q1 },
    ], def.layout) };
    expectWritten(await runner.run(def, { anchor: null, useSelection: false }));
    expect(cellText(2, 0)).toBe("Q1 2025");
    expect(cellText(3, 0)).toBe("Jan 2025");
    const hit = runner.memberAt(def.id, 1, 0);
    expect(hit?.member.expanded).toBe(true);
    expect(host.snapshot().cells.get("3,0")?.indent).toBe(2);

    const descendants = runner.descendants(def.id, "rows", hit!.member);
    expect(descendants).toEqual([q1]);
    def = { ...def, layout: layoutOf([
      { type: "collapse", hierarchy: calendar, member: year2025, descendants },
    ], def.layout) };
    expect(def.layout.rows[0].expanded).toEqual([]);
    const collapsed = expectWritten(await runner.run(def, { anchor: null, useSelection: false }));
    expect(collapsed.extent.rows).toBe(4);
  });

  it("reports no data with the header rows still written", async () => {
    host.select(0, 0);
    // Box is not offered in Rey: this selection has no facts.
    const layout = layoutOf([
      { type: "addMeasure", measure: measure("Rides") },
      { type: "addHierarchy", hierarchy: hierarchy("Calendar") },
      { type: "addHierarchy", hierarchy: hierarchy("City"), area: "filters" },
      { type: "setMembers", hierarchy: hierarchy("City").uniqueName, members: ["[Geography].[City].&[2]"], captions: ["Rey"] },
      { type: "addHierarchy", hierarchy: hierarchy("Service Type"), area: "filters" },
      { type: "setMembers", hierarchy: hierarchy("Service Type").uniqueName, members: ["[Service].[Service Type].&[3]"], captions: ["Box"] },
    ]);
    const outcome = expectWritten(await runner.run(report(layout), { anchor: null, useSelection: true }));
    expect(outcome.runtime.dataRows).toBe(0);
    // Two filter rows, a blank row, then the header row.
    expect(outcome.extent.rows).toBe(4);
    expect(cellText(0, 0)).toBe("City");
    expect(cellText(0, 1)).toBe("Rey");
    expect(cellText(3, 0)).toBe("Calendar");
  });
});

describe("placeholder of a report with no fields", () => {
  const at = (row: number, col: number) => ({ sheetId: "preview-sheet-1", sheetName: "Sheet1", row, col });

  it("is written where a new report is put, tracked by its name, and saved", async () => {
    const def = report(emptyLayout());
    const outcome = await runner.run(def, { anchor: at(6, 2), useSelection: false });
    expect(outcome).toMatchObject({ kind: "empty", extent: { row: 6, col: 2, rows: 1, cols: 1 } });
    expect(cellText(6, 2)).toBe(PLACEHOLDER_TEXT);
    expect(host.snapshot().cells.get("6,2")).toMatchObject({ muted: true, format: "@" });
    expect(await host.getReportExtent(def.id)).toMatchObject({ row: 6, col: 2, rows: 1, cols: 1 });
    expect(await host.loadReports()).toEqual([def]);
    // Selecting it finds the report.
    host.select(6, 2);
    expect((await host.getSelection()).report).toMatchObject({ id: def.id, row: 0, col: 0 });
  });

  it("is replaced by the first real result, at the same cell", async () => {
    const def = report(emptyLayout());
    await runner.run(def, { anchor: at(6, 2), useSelection: false });
    const filled = { ...def, layout: basic() };
    const written = expectWritten(await runner.run(filled, { anchor: null, useSelection: false }));
    expect(written.extent).toMatchObject({ row: 6, col: 2, rows: 4, cols: 2 });
    expect(cellText(6, 2)).toBe("Calendar");
    expect(host.snapshot().cells.get("6,2")?.muted).toBeUndefined();
    expect(asked).toEqual([]);
  });

  it("counts a layout with only a filter as having no fields", async () => {
    const def = report(layoutOf([{ type: "addHierarchy", hierarchy: hierarchy("City"), area: "filters" }]));
    const outcome = await runner.run(def, { anchor: at(0, 0), useSelection: false });
    expect(outcome.kind).toBe("empty");
    expect(cellText(0, 0)).toBe(PLACEHOLDER_TEXT);
    expect(host.snapshot().cells.size).toBe(1);
  });

  it("asks before covering a cell that holds something, and leaves it on No", async () => {
    host.type(6, 2, "mine");
    answer = false;
    const def = report(emptyLayout());
    expect(await runner.run(def, { anchor: at(6, 2), useSelection: false })).toEqual({ kind: "declined" });
    expect(asked).toHaveLength(1);
    expect(asked[0]).toMatchObject({ row: 6, col: 2, rows: 1, cols: 1 });
    expect(cellText(6, 2)).toBe("mine");
    expect(await host.loadReports()).toEqual([]);

    answer = true;
    expect((await runner.run(def, { anchor: at(6, 2), useSelection: false })).kind).toBe("empty");
    expect(cellText(6, 2)).toBe(PLACEHOLDER_TEXT);
  });

  it("is not rewritten over itself with a question", async () => {
    const def = report(emptyLayout());
    await runner.run(def, { anchor: at(1, 1), useSelection: false });
    expect((await runner.run(def, { anchor: null, useSelection: false })).kind).toBe("empty");
    expect(asked).toEqual([]);
    expect(host.snapshot().cells.size).toBe(1);
  });
});

describe("a report that lost its cells, and moving a report", () => {
  it("is written again, with its fields, at the cell it is given", async () => {
    host.select(1, 1);
    const def = report(basic());
    expectWritten(await runner.run(def, { anchor: null, useSelection: true }));

    // The sheet is deleted: the definition stays, the cells and the name's target go.
    host.deleteSheet();
    expect(await host.getReportExtent(def.id)).toBeNull();
    expect(await host.loadReports()).toEqual([def]);
    expect(await runner.run(def, { anchor: null, useSelection: false })).toEqual({ kind: "needs-anchor" });

    const selected = (await host.getSelection()).cell;
    expect(selected).toMatchObject({ sheetName: "Sheet2", row: 0, col: 0 });
    const again = expectWritten(await runner.run(def, { anchor: selected, useSelection: false }));
    expect(again.extent).toMatchObject({ sheetName: "Sheet2", row: 0, col: 0, rows: 4, cols: 2 });
    expect(cellText(0, 0)).toBe("Calendar");
    expect(cellText(3, 0)).toBe("Grand Total");
    expect(await host.loadReports()).toEqual([def]);
  });

  it("relocate moves the block: old cells cleared, name at the new place", async () => {
    host.select(0, 0);
    const def = report(basic());
    expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    const target = { sheetId: "preview-sheet-1", sheetName: "Sheet1", row: 10, col: 3 };
    const moved = expectWritten(await runner.run(def, { anchor: target, useSelection: false, relocate: true }));
    expect(moved.extent).toMatchObject({ row: 10, col: 3, rows: 4, cols: 2 });
    expect(cellText(0, 0)).toBeUndefined();
    expect(cellText(10, 3)).toBe("Calendar");
    expect(host.snapshot().cells.size).toBe(8);
    expect(await host.getReportExtent(def.id)).toMatchObject({ row: 10, col: 3 });
    expect(asked).toEqual([]);
  });

  it("without relocate an offered anchor does not move a report that has cells", async () => {
    host.select(0, 0);
    const def = report(basic());
    expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    const target = { sheetId: "preview-sheet-1", sheetName: "Sheet1", row: 10, col: 3 };
    const same = expectWritten(await runner.run(def, { anchor: target, useSelection: false }));
    expect(same.extent).toMatchObject({ row: 0, col: 0 });
  });

  it("relocate moves a placeholder too", async () => {
    const def = report(emptyLayout());
    const from = { sheetId: "preview-sheet-1", sheetName: "Sheet1", row: 2, col: 2 };
    const to = { sheetId: "preview-sheet-1", sheetName: "Sheet1", row: 8, col: 1 };
    await runner.run(def, { anchor: from, useSelection: false });
    const outcome = await runner.run(def, { anchor: to, useSelection: false, relocate: true });
    expect(outcome).toMatchObject({ kind: "empty", extent: { row: 8, col: 1 } });
    expect(cellText(2, 2)).toBeUndefined();
    expect(cellText(8, 1)).toBe(PLACEHOLDER_TEXT);
  });
});

describe("filter cells of a written report", () => {
  // Filter: Year. Columns: City. Rows: Calendar. Values: Rides.
  // Row 0: filter. Row 1: blank. Row 2: City header. Row 3: measures + corner. Row 4+: data.
  function full(): Layout {
    return layoutOf(
      [
        { type: "addHierarchy", hierarchy: hierarchy("Year"), area: "filters" },
        { type: "addHierarchy", hierarchy: hierarchy("City"), area: "columns" },
      ],
      basic(),
    );
  }
  const size = (extent: Extent) => ({ rows: extent.rows, cols: extent.cols });

  it("writes the arrow format on the three kinds of filter cell, and the column field's caption", async () => {
    host.select(0, 0);
    expectWritten(await runner.run(report(full()), { anchor: null, useSelection: true }));
    const cell = (row: number, col: number) => host.snapshot().cells.get(row + "," + col);
    expect(cell(0, 0)).toMatchObject({ value: "Year", format: "@" });
    expect(cell(0, 1)).toMatchObject({ value: "All", format: FILTER_FORMAT });
    expect(cell(2, 0)).toMatchObject({ value: "City", format: FILTER_FORMAT });
    expect(cell(3, 0)).toMatchObject({ value: "Calendar", format: FILTER_FORMAT });
    // Member and measure headers are plain text.
    expect(cell(2, 1)).toMatchObject({ value: "Tehran", format: "@" });
    expect(cell(3, 1)).toMatchObject({ value: "Rides", format: "@" });
  });

  it("knows the field behind each button, and that other cells are not buttons", async () => {
    host.select(0, 0);
    const def = report(full());
    const written = expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    const at = (row: number, col: number) => ({ row, col, ...size(written.extent) });
    expect(runner.buttonAt(def.id, at(0, 0))?.caption).toBe("Year");
    expect(runner.buttonAt(def.id, at(0, 1))?.caption).toBe("Year");
    expect(runner.buttonAt(def.id, at(2, 0))?.caption).toBe("City");
    expect(runner.buttonAt(def.id, at(3, 0))?.caption).toBe("Calendar");
    expect(runner.buttonAt(def.id, at(2, 1))).toBeNull(); // a city's header
    expect(runner.buttonAt(def.id, at(3, 1))).toBeNull(); // a measure caption
    expect(runner.buttonAt(def.id, at(4, 0))).toBeNull(); // a member
    expect(runner.buttonAt(def.id, at(4, 1))).toBeNull(); // a number
    expect(runner.buttonAt(def.id, at(1, 0))).toBeNull(); // the blank row
  });

  it("also offers the column field for the cells of its header row", async () => {
    host.select(0, 0);
    const def = report(full());
    const written = expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    const at = (row: number, col: number) => ({ row, col, ...size(written.extent) });
    expect(runner.fieldAt(def.id, at(2, 1))?.caption).toBe("City");
    expect(runner.fieldAt(def.id, at(2, written.extent.cols - 1))?.caption).toBe("City"); // Grand Total
    expect(runner.fieldAt(def.id, at(0, 1))?.caption).toBe("Year");
    expect(runner.fieldAt(def.id, at(3, 1))).toBeNull();
    expect(runner.fieldAt(def.id, at(4, 0))).toBeNull();
  });

  it("cannot tell without a result, or once the block has changed size", async () => {
    host.select(0, 0);
    const def = report(full());
    expect(runner.buttonAt(def.id, { row: 0, col: 1, rows: 7, cols: 10 })).toBeUndefined();
    const written = expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    const grown = { rows: written.extent.rows + 1, cols: written.extent.cols };
    expect(runner.buttonAt(def.id, { row: 0, col: 1, ...grown })).toBeUndefined();
    expect(runner.fieldAt(def.id, { row: 0, col: 1, ...grown })).toBeUndefined();
    runner.forget(def.id);
    expect(runner.buttonAt(def.id, { row: 0, col: 1, ...size(written.extent) })).toBeUndefined();
  });

  it("agrees with the guess made from the layout alone, for every cell of the block", async () => {
    host.select(0, 0);
    const layouts = [
      full(),
      basic(),
      layoutOf([{ type: "addMeasure", measure: measure("Rides") }]),
      layoutOf([
        { type: "addMeasure", measure: measure("Rides") },
        { type: "addHierarchy", hierarchy: hierarchy("City"), area: "columns" },
        { type: "addHierarchy", hierarchy: hierarchy("Year"), area: "columns" },
        { type: "addHierarchy", hierarchy: hierarchy("Service Type"), area: "filters" },
      ]),
    ];
    for (const layout of layouts) {
      const def = report(layout);
      const written = expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
      for (let row = 0; row < written.extent.rows; row++) {
        for (let col = 0; col < written.extent.cols; col++) {
          const guess = guessFieldCell(layout, row, col);
          const guessed = guess ? layout[guess.area][guess.index] : null;
          expect(runner.buttonAt(def.id, { row, col, ...size(written.extent) }) ?? null).toEqual(guessed ?? null);
        }
      }
      host.reset();
      host.select(0, 0);
    }
  });

  it("has no buttons on a placeholder", async () => {
    const def = report(emptyLayout());
    await runner.run(def, { anchor: { sheetId: "preview-sheet-1", sheetName: "Sheet1", row: 0, col: 0 }, useSelection: false });
    expect(runner.buttonAt(def.id, { row: 0, col: 0, rows: 1, cols: 1 })).toBeUndefined();
    expect(host.snapshot().cells.get("0,0")?.format).toBe("@");
  });
});

describe("sorting, filtering, show-as and number formats through the sample cube", () => {
  const city = () => hierarchy("City").uniqueName;
  const rides = () => measure("Rides").uniqueName;
  function cityRows(extra: LayoutAction[] = []): Layout {
    return layoutOf([
      { type: "addMeasure", measure: measure("Rides") },
      { type: "addHierarchy", hierarchy: hierarchy("City") },
      ...extra,
    ]);
  }
  /** Row captions of the written block, from row 1 down, without the Grand Total line. */
  function captions(): string[] {
    const out: string[] = [];
    for (let row = 1; ; row++) {
      const text = cellText(row, 0);
      if (text === undefined || text === "Grand Total") return out;
      out.push(String(text));
    }
  }
  function column(col: number, rows: number): number[] {
    return Array.from({ length: rows }, (_, i) => Number(cellText(1 + i, col)));
  }

  it("sorts items by caption and by a measure, and back to the cube's order", async () => {
    host.select(0, 0);
    const def = report(cityRows());
    expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    const cubeOrder = captions();
    expect(cubeOrder.slice(0, 3)).toEqual(["Tehran", "Rey", "Shiraz"]);

    const az = report(cityRows([{ type: "setSort", hierarchy: city(), sort: { by: "caption", order: "asc" } }]), def.id);
    expectWritten(await runner.run(az, { anchor: null, useSelection: false }));
    expect(captions()).toEqual(cubeOrder.slice().sort());

    const byRides = report(
      cityRows([{ type: "setSort", hierarchy: city(), sort: { by: "measure", measure: rides(), order: "desc" } }]),
      def.id,
    );
    expectWritten(await runner.run(byRides, { anchor: null, useSelection: false }));
    const values = column(1, cubeOrder.length);
    expect(values).toEqual(values.slice().sort((a, b) => b - a));
  });

  it("applies label filters, and totals count only what is left", async () => {
    host.select(0, 0);
    const def = report(
      cityRows([{ type: "setLabelFilter", hierarchy: city(), filter: { operator: "beginsWith", value: "ma" } }]),
    );
    const written = expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    expect(captions()).toEqual(["Marvdasht", "Mashhad"]);
    const total = Number(cellText(written.extent.rows - 1, 1));
    expect(total).toBe(column(1, 2).reduce((a, b) => a + b, 0));

    const contains = report(
      cityRows([{ type: "setLabelFilter", hierarchy: city(), filter: { operator: "contains", value: "SH" } }]),
      def.id,
    );
    expectWritten(await runner.run(contains, { anchor: null, useSelection: false }));
    expect(captions()).toEqual(["Shiraz", "Marvdasht", "Kashan", "Mashhad", "Neyshabur"]);
  });

  it("keeps the top N items by a measure, and their total", async () => {
    host.select(0, 0);
    const all = report(cityRows([{ type: "setSort", hierarchy: city(), sort: { by: "measure", measure: rides(), order: "desc" } }]));
    expectWritten(await runner.run(all, { anchor: null, useSelection: true }));
    const top3 = captions().slice(0, 3);

    const def = report(
      cityRows([{ type: "setValueFilter", hierarchy: city(), filter: { operator: "top", measure: rides(), value: 3, topBy: "items" } }]),
      all.id,
    );
    const written = expectWritten(await runner.run(def, { anchor: null, useSelection: false }));
    expect(captions().slice().sort()).toEqual(top3.slice().sort());
    expect(Number(cellText(written.extent.rows - 1, 1))).toBe(column(1, 3).reduce((a, b) => a + b, 0));
  });

  it("writes nothing for a filter nothing passes, rather than pretending", async () => {
    host.select(0, 0);
    const def = report(
      cityRows([{ type: "setLabelFilter", hierarchy: city(), filter: { operator: "equals", value: "Atlantis" } }]),
    );
    const written = expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    expect(written.runtime.dataRows).toBe(0);
  });

  it("shows values as a percent of the grand total that add up to 100%", async () => {
    host.select(0, 0);
    const def = report(
      cityRows([{ type: "setShowAs", measure: rides(), showAs: "percentOfGrandTotal" }]),
    );
    const written = expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    expect(cellText(0, 1)).toBe("Rides (% of Grand Total)");
    const shares = column(1, written.runtime.dataRows - 1);
    expect(shares.reduce((a, b) => a + b, 0)).toBeCloseTo(1, 9);
    expect(Number(cellText(written.extent.rows - 1, 1))).toBeCloseTo(1, 9);
    expect(host.snapshot().cells.get("1,1")?.format).toBe("0.00%");
  });

  it("applies the user's number format to that value field's cells only", async () => {
    host.select(0, 0);
    const def = report(
      cityRows([
        { type: "addMeasure", measure: measure("Fare") },
        { type: "setNumberFormat", measure: measure("Fare").uniqueName, format: "$#,##0.00" },
      ]),
    );
    expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    expect(host.snapshot().cells.get("1,1")?.format).toBe("#,##0");
    expect(host.snapshot().cells.get("1,2")?.format).toBe("$#,##0.00");
  });
});

describe("Show Details", () => {
  function layout(): Layout {
    return layoutOf([
      { type: "addMeasure", measure: measure("Rides") },
      { type: "addMeasure", measure: measure("Fare") },
      { type: "addHierarchy", hierarchy: hierarchy("City") },
      { type: "addHierarchy", hierarchy: hierarchy("Year"), area: "columns" },
    ]);
  }
  // Row 0: Year header. Row 1: measures + corner. Rows 2+: cities. Columns: 2025 x (Rides, Fare), 2026 x …, Grand Total x …

  it("finds the members and the measure behind a value cell", async () => {
    host.select(0, 0);
    const def = report(layout());
    const written = expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    const at = (row: number, col: number) => ({ row, col, rows: written.extent.rows, cols: written.extent.cols });
    expect(runner.detailsFor(def.id, at(2, 1))).toEqual({
      members: ["[Geography].[City].&[1]", "[Date].[Year].&[2025]"],
      measure: measure("Rides").uniqueName,
      layout: def.layout,
    });
    expect(runner.detailsFor(def.id, at(2, 2))?.measure).toBe(measure("Fare").uniqueName);
    // Grand Total column: the year adds nothing.
    expect(runner.detailsFor(def.id, at(2, 5))?.members).toEqual(["[Geography].[City].&[1]"]);
    // Grand Total line and column: nothing at all.
    expect(runner.detailsFor(def.id, at(written.extent.rows - 1, 5))?.members).toEqual([]);
  });

  it("is nothing for headers, and unknown without a result", async () => {
    host.select(0, 0);
    const def = report(layout());
    expect(runner.detailsFor(def.id, { row: 2, col: 1, rows: 12, cols: 7 })).toBeUndefined();
    const written = expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    const at = (row: number, col: number) => ({ row, col, rows: written.extent.rows, cols: written.extent.cols });
    expect(runner.detailsFor(def.id, at(0, 1))).toBeNull();
    expect(runner.detailsFor(def.id, at(1, 1))).toBeNull();
    expect(runner.detailsFor(def.id, at(2, 0))).toBeNull();
  });

  it("is nothing without value fields: the cell's measure is unknown", async () => {
    host.select(0, 0);
    const def = report(layoutOf([{ type: "addHierarchy", hierarchy: hierarchy("City") }]));
    const written = expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    expect(runner.detailsFor(def.id, { row: 1, col: 1, rows: written.extent.rows, cols: written.extent.cols })).toBeNull();
  });

  it("returns the fact rows behind a cell from the sample cube", async () => {
    host.select(0, 0);
    const def = report(layout());
    const written = expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    const target = runner.detailsFor(def.id, { row: 2, col: 1, rows: written.extent.rows, cols: written.extent.cols })!;
    const response = await api.details("demo", { catalog: CATALOG, cube: CUBE, ...target });
    expect(response.columns[response.columns.length - 1]).toBe("Rides");
    expect(response.rows.length).toBeGreaterThan(0);
    expect(response.truncated).toBe(false);
    expect(response.rows.every((row) => row[2] === "Tehran" && String(row[0]).endsWith("2025"))).toBe(true);
    const sum = response.rows.reduce((total, row) => total + Number(row[row.length - 1]), 0);
    expect(sum).toBe(Number(cellText(2, 1)));

    const few = await api.details("demo", { catalog: CATALOG, cube: CUBE, ...target, maxRows: 3 });
    expect(few.rows).toHaveLength(3);
    expect(few.truncated).toBe(true);
  });

  it("refuses when a field filtered to several items is not pinned by the cell", async () => {
    const filtered = layoutOf(
      [
        { type: "addHierarchy", hierarchy: hierarchy("Service Type"), area: "filters" },
        {
          type: "setMembers",
          hierarchy: hierarchy("Service Type").uniqueName,
          members: ["[Service].[Service Type].&[1]", "[Service].[Service Type].&[2]"],
          captions: ["Ride", "Bike"],
        },
      ],
      layout(),
    );
    await expect(
      api.details("demo", {
        catalog: CATALOG,
        cube: CUBE,
        layout: filtered,
        members: ["[Geography].[City].&[1]"],
        measure: measure("Rides").uniqueName,
      }),
    ).rejects.toMatchObject({ error: { category: "bad-request", message: expect.stringContaining("filtered to several items") } });
  });
});
