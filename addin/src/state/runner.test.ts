import { beforeEach, describe, expect, it } from "vitest";
import type { Api } from "../api";
import { ApiFailure, HelperDown } from "../apiErrors";
import { createMockApi } from "../dev/mockApi";
import type { Extent, ReportDef } from "../sheet/host";
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

  it("writes nothing and saves nothing while the layout is empty", async () => {
    const def = report(emptyLayout());
    expect(await runner.run(def, { anchor: null, useSelection: true })).toEqual({
      kind: "empty",
      cleared: false,
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

  it("empties the block but keeps its place when every field is removed", async () => {
    host.select(4, 4);
    const def = report(basic());
    expectWritten(await runner.run(def, { anchor: null, useSelection: true }));
    const emptied = { ...def, layout: emptyLayout() };
    expect(await runner.run(emptied, { anchor: null, useSelection: false })).toEqual({
      kind: "empty",
      cleared: true,
    });
    expect(host.snapshot().cells.size).toBe(0);
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
