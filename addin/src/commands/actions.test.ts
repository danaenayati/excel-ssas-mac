import { beforeEach, describe, expect, it } from "vitest";
import type { Api } from "../api";
import { HelperDown } from "../apiErrors";
import { createMockApi } from "../dev/mockApi";
import type { Rect } from "../sheet/geometry";
import type { ReportDef } from "../sheet/host";
import { PreviewHost } from "../sheet/preview";
import { emptyLayout, layoutReducer, type LayoutAction } from "../state/layout";
import { newReportDef } from "../state/report";
import { ReportRunner } from "../state/runner";
import type { Fields, Layout } from "../types";
import { COMMANDS, MESSAGES, labelCross, runCommand, type CommandDeps, type CommandName } from "./actions";

const CATALOG = "Sample";
const CUBE = "Rides";
const YEAR_2025 = "[Date].[Calendar].[Year].&[2025]";
const YEAR_2026 = "[Date].[Calendar].[Year].&[2026]";

let api: Api;
let host: PreviewHost;
let fields: Fields;
let deps: CommandDeps;

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
/** Calendar on rows, Rides as the value: header, 2025, 2026, Grand Total. */
function basic(): Layout {
  return layoutOf([
    { type: "addMeasure", measure: measure("Rides") },
    { type: "addHierarchy", hierarchy: hierarchy("Calendar") },
  ]);
}
function cellText(row: number, col: number) {
  const cell = host.snapshot().cells.get(row + "," + col);
  return cell ? cell.value : undefined;
}
/** Put a report on the sheet the way the pane does, with its top-left at (row, col). */
async function place(layout: Layout, row = 0, col = 0, connectionId = "demo"): Promise<ReportDef> {
  const def: ReportDef = { ...newReportDef("demo", CATALOG, CUBE), layout };
  const runner = new ReportRunner({ api, host, confirmOverwrite: async () => true });
  const anchor = { sheetId: "preview-sheet-1", sheetName: "Sheet1", row, col };
  const outcome = await runner.run(def, { anchor, useSelection: false });
  if (outcome.kind !== "written" && outcome.kind !== "empty") throw new Error("not written: " + outcome.kind);
  if (connectionId !== def.connectionId) {
    const moved = { ...def, connectionId };
    await host.saveReport(moved);
    return moved;
  }
  return def;
}
async function saved(id: string): Promise<ReportDef> {
  const found = (await host.loadReports()).find((d) => d.id === id);
  if (!found) throw new Error("not saved: " + id);
  return found;
}
async function run(name: CommandName, row: number, col: number) {
  host.select(row, col);
  return runCommand(name, deps);
}

beforeEach(async () => {
  api = createMockApi({ delayMs: 0 });
  host = new PreviewHost({ persist: false });
  fields = await api.fields("demo", CATALOG, CUBE);
  deps = {
    api,
    host,
    // What the preview sheet holds, as Excel's range.values would give it.
    readValues: async (sheetId: string, rect: Rect) => {
      const snapshot = host.snapshot();
      if (snapshot.activeId !== sheetId) throw new Error("not the active sheet");
      const out: unknown[][] = [];
      for (let r = 0; r < rect.rows; r++) {
        const line: unknown[] = [];
        for (let c = 0; c < rect.cols; c++) {
          const cell = snapshot.cells.get(rect.row + r + "," + (rect.col + c));
          line.push(cell ? cell.value : "");
        }
        out.push(line);
      }
      return out;
    },
  };
});

describe("where the commands work", () => {
  it("says so for a cell outside every report", async () => {
    await place(basic(), 0, 0);
    for (const name of COMMANDS.filter((c) => c !== "refreshAll")) {
      expect(await run(name, 10, 10)).toBe(MESSAGES.notInReport);
    }
  });

  it("asks for a connection when the report's is not on this Mac", async () => {
    await place(basic(), 0, 0, "gone");
    expect(await run("refresh", 1, 1)).toBe(MESSAGES.noConnection);
    expect(await run("expand", 1, 0)).toBe(MESSAGES.noConnection);
  });

  it("explains that a report without fields has nothing to drill, and refreshes it quietly", async () => {
    await place(emptyLayout(), 2, 2);
    expect(await run("expand", 2, 2)).toBe(MESSAGES.noFields);
    expect(await run("showDetails", 2, 2)).toBe(MESSAGES.noFields);
    expect(await run("refresh", 2, 2)).toBeNull();
  });

  it("passes on a failure as one sentence", async () => {
    await place(basic(), 0, 0);
    deps = { ...deps, api: { ...api, query: async () => Promise.reject(new HelperDown("connection refused")) } };
    expect(await run("expand", 1, 0)).toBe(MESSAGES.helperDown);
  });
});

describe("Refresh", () => {
  it("writes the report again from the cube", async () => {
    const def = await place(basic(), 0, 0);
    const value = cellText(1, 1);
    host.type(1, 1, 12345);
    expect(await run("refresh", 1, 1)).toBeNull();
    expect(cellText(1, 1)).toBe(value);
    expect(await saved(def.id)).toEqual(def);
  });
});

describe("Expand and Collapse", () => {
  it("expands the item in the cell, saves the definition and writes the report", async () => {
    const def = await place(basic(), 0, 0);
    expect(await run("expand", 1, 0)).toBeNull();
    expect((await saved(def.id)).layout.rows[0].expanded).toEqual([YEAR_2025]);
    expect(cellText(1, 0)).toBe("2025");
    expect(cellText(2, 0)).toBe("Q1 2025");
  });

  it("collapses it again, with what was expanded under it", async () => {
    const calendar = hierarchy("Calendar").uniqueName;
    const q1 = "[Date].[Calendar].[Quarter].&[2025]&[1]";
    const def = await place(
      layoutOf(
        [
          { type: "expand", hierarchy: calendar, member: YEAR_2025 },
          { type: "expand", hierarchy: calendar, member: q1 },
        ],
        basic(),
      ),
      0,
      0,
    );
    expect(await run("collapse", 1, 0)).toBeNull();
    expect((await saved(def.id)).layout.rows[0].expanded).toEqual([]);
    expect(await host.getReportExtent(def.id)).toMatchObject({ rows: 4 });
  });

  it("says why when there is nothing to do", async () => {
    await place(basic(), 0, 0);
    expect(await run("collapse", 1, 0)).toBe("“2025” is not expanded.");
    expect(await run("expand", 1, 1)).toBe(MESSAGES.expandWhere); // a number
    expect(await run("collapse", 3, 0)).toBe(MESSAGES.collapseWhere); // Grand Total
    expect(await run("expand", 1, 0)).toBeNull();
    expect(await run("expand", 1, 0)).toBe("“2025” is already expanded.");

    const flat = layoutOf([
      { type: "addMeasure", measure: measure("Rides") },
      { type: "addHierarchy", hierarchy: hierarchy("City") },
    ]);
    await place(flat, 20, 0);
    expect(await run("expand", 21, 0)).toMatch(/has no items under it\.$/);
  });

  it("leaves the report as it is when it would cover other cells", async () => {
    const def = await place(basic(), 0, 0);
    host.type(5, 0, "keep me");
    expect(await run("expand", 1, 0)).toBe(MESSAGES.declined);
    expect(cellText(5, 0)).toBe("keep me");
    expect(await saved(def.id)).toEqual(def);
  });
});

describe("Expand and Collapse Entire Field", () => {
  it("works from an item or from the field's name", async () => {
    const def = await place(basic(), 0, 0);
    expect(await run("expandField", 2, 0)).toBeNull();
    expect((await saved(def.id)).layout.rows[0].expanded.slice().sort()).toEqual([YEAR_2025, YEAR_2026]);
    // One level deeper each time, as in Excel: from the caption cell now.
    expect(await run("expandField", 0, 0)).toBeNull();
    expect((await saved(def.id)).layout.rows[0].expanded.length).toBeGreaterThan(2);
    expect(await run("collapseField", 0, 0)).toBeNull();
    expect((await saved(def.id)).layout.rows[0].expanded).toEqual([]);
    expect(await host.getReportExtent(def.id)).toMatchObject({ rows: 4 });
  });

  it("says why when there is nothing to do", async () => {
    await place(basic(), 0, 0);
    expect(await run("collapseField", 1, 0)).toBe("Nothing in “Calendar” is expanded.");
    expect(await run("expandField", 1, 1)).toBe(MESSAGES.fieldWhere);

    const flat = layoutOf([
      { type: "addMeasure", measure: measure("Rides") },
      { type: "addHierarchy", hierarchy: hierarchy("City") },
    ]);
    await place(flat, 20, 0);
    expect(await run("expandField", 21, 0)).toMatch(/have no items under them\.$/);
  });

  it("does not take a page filter for a row or column field", async () => {
    await place(
      layoutOf([{ type: "addHierarchy", hierarchy: hierarchy("City"), area: "filters" }], basic()),
      0,
      0,
    );
    expect(await run("expandField", 0, 0)).toBe(MESSAGES.fieldWhere);
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

  it("puts the rows behind a value cell on a new sheet", async () => {
    const def = await place(layout(), 0, 0);
    const value = Number(cellText(2, 1));
    expect(await run("showDetails", 2, 1)).toBeNull();
    const snapshot = host.snapshot();
    expect(snapshot.sheetName).toBe("Details");
    expect(cellText(0, 0)).not.toBe("");
    // The Rides of the rows add up to the number in the cell.
    let sum = 0;
    let rows = 0;
    const last = Array.from(snapshot.cells.keys()).reduce((m, k) => Math.max(m, Number(k.split(",")[1])), 0);
    for (let r = 1; snapshot.cells.has(r + ",0"); r++) {
      rows++;
      sum += Number(cellText(r, last));
    }
    expect(rows).toBeGreaterThan(0);
    expect(sum).toBe(value);
    // The report itself is untouched.
    expect(await saved(def.id)).toEqual(def);
  });

  it("does not work on labels", async () => {
    await place(layout(), 0, 0);
    expect(await run("showDetails", 2, 0)).toBe(MESSAGES.notValue);
    expect(await run("showDetails", 0, 1)).toBe(MESSAGES.notValue);
    expect(host.snapshot().sheetName).toBe("Sheet1");
  });
});

describe("a report that changed since it was last refreshed", () => {
  it("is not drilled when a label on the cell's row differs from the cube's", async () => {
    const def = await place(basic(), 0, 0);
    host.type(1, 0, "1999");
    expect(await run("showDetails", 1, 1)).toBe(MESSAGES.stale);
    expect(await run("expand", 1, 0)).toBe(MESSAGES.stale);
    expect(await saved(def.id)).toEqual(def);
    // A changed number is not a changed report.
    await run("refresh", 1, 1);
    host.type(2, 1, 1);
    expect(await run("expand", 1, 0)).toBeNull();
  });

  it("is not drilled when the cube's result has another size", async () => {
    const def = await place(basic(), 0, 0);
    host.insertRows(2, 1); // inside the block: the report's name grows with it
    expect(await run("expand", 1, 0)).toBe(MESSAGES.stale);
    expect(await saved(def.id)).toEqual(def);
  });
});

describe("Refresh All", () => {
  it("refreshes every report quietly", async () => {
    await place(basic(), 0, 0);
    await place(basic(), 0, 5);
    host.type(1, 1, 1);
    host.type(1, 6, 1);
    expect(await run("refreshAll", 30, 30)).toBeNull();
    expect(cellText(1, 1)).not.toBe(1);
    expect(cellText(1, 6)).not.toBe(1);
  });

  it("says which reports it could not refresh", async () => {
    await place(basic(), 0, 0);
    await place(basic(), 0, 5, "gone");
    expect(await run("refreshAll", 30, 30)).toBe(
      "1 of 2 reports refreshed. 1 needs a connection (open the Cube Explorer panel).",
    );
  });
});

describe("labelCross", () => {
  const shape = { filterRows: 0, headerRows: 1, headerCols: 2, dataRows: 5, dataCols: 3, headerTop: 0, dataTop: 1 };
  it("takes the row labels and the column headers of a value cell", () => {
    expect(labelCross(shape, 5, 3, 3)).toEqual([
      { row: 3, col: 0, rows: 1, cols: 2 },
      { row: 0, col: 3, rows: 1, cols: 1 },
    ]);
  });
  it("takes a label column down to the cell, and a whole header row", () => {
    expect(labelCross(shape, 5, 4, 1)).toEqual([
      { row: 4, col: 0, rows: 1, cols: 2 },
      { row: 0, col: 1, rows: 5, cols: 1 },
    ]);
    expect(labelCross(shape, 5, 0, 4)).toEqual([
      { row: 0, col: 0, rows: 1, cols: 5 },
      { row: 0, col: 4, rows: 1, cols: 1 },
    ]);
  });
});
