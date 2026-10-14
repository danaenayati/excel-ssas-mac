// Charts that follow a report, end to end: the sample cube, the runner and the preview
// host. Every write of the report (Refresh, Refresh All, a layout change) updates its
// charts; a chart deleted in Excel drops out of the definition quietly.

import { beforeEach, describe, expect, it } from "vitest";
import type { Api } from "../api";
import { createMockApi } from "../dev/mockApi";
import type { ChartTable, ReportChart } from "../sheet/charts";
import { SheetError, type ChartsUpdate, type ReportDef } from "../sheet/host";
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
/** Rides by City. */
function byCity(): Layout {
  return layoutOf([
    { type: "addMeasure", measure: measure("Rides") },
    { type: "addHierarchy", hierarchy: hierarchy("City") },
  ]);
}
function report(layout: Layout): ReportDef {
  return { ...newReportDef("demo", CATALOG, CUBE), layout };
}
function written(outcome: RunOutcome) {
  if (outcome.kind !== "written") throw new Error("expected written, got " + JSON.stringify(outcome));
  return outcome;
}
async function write(def: ReportDef, at = { row: 0, col: 0 }) {
  return written(
    await runner.run(def, {
      anchor: { sheetId: "preview-sheet-1", sheetName: "Sheet1", ...at },
      useSelection: false,
    }),
  );
}
async function insert(def: ReportDef, type: ReportChart["type"] = "ColumnClustered") {
  const result = await runner.insertChart(def, type);
  if ("error" in result) throw new Error(result.error.message);
  return result;
}
function shownTable(chartId: string): ChartTable {
  const chart = host.allCharts().find((c) => c.id === chartId);
  if (!chart) throw new Error("no chart " + chartId);
  return chart.table;
}
async function savedCharts(id: string) {
  const def = (await host.loadReports()).find((d) => d.id === id);
  return def ? def.charts : "no definition";
}

beforeEach(async () => {
  api = createMockApi({ delayMs: 0 });
  host = new PreviewHost({ persist: false });
  fields = await api.fields("demo", CATALOG, CUBE);
  runner = new ReportRunner({ api, host, confirmOverwrite: async () => true });
});

describe("Insert Chart", () => {
  it("needs the report written first", async () => {
    const result = await runner.insertChart(report(byCity()), "Pie");
    expect(result).toEqual({ error: { message: "Refresh the report first, then insert the chart." } });
    expect(host.allCharts()).toEqual([]);
  });

  it("adds a chart of the report's last result and saves it in the definition", async () => {
    const def = report(byCity());
    const outcome = await write(def);
    expect(outcome.def).toEqual(def); // no charts yet: no `charts` field either
    expect(outcome.runtime.chart?.title).toBe("Rides");

    const { def: saved, chart } = await insert(def, "Pie");
    expect(chart).toMatchObject({ type: "Pie", sheetId: "preview-sheet-1", title: "Rides" });
    expect(saved.charts).toEqual([chart]);
    expect(await savedCharts(def.id)).toEqual([chart]);
    expect(runner.chartsOf(def)).toEqual([chart]);

    const table = shownTable(chart.id);
    expect(table.series).toEqual(["Total"]);
    expect(table.categories.length).toBeGreaterThan(1);
    expect(table.categories).not.toContain("Grand Total");
  });

  it("is refused when a write queued before it leaves the report without fields", async () => {
    const def = report(byCity());
    await write(def);
    const emptying = runner.run({ ...def, layout: emptyLayout() }, { anchor: null, useSelection: false });
    const result = await runner.insertChart(def, "Pie");
    expect((await emptying).kind).toBe("empty");
    expect(result).toEqual({ error: { message: "Refresh the report first, then insert the chart." } });
    expect(host.allCharts()).toEqual([]);
  });

  it("refuses a report with more columns than a chart can show", async () => {
    const def = report(
      layoutOf([
        { type: "addMeasure", measure: measure("Rides") },
        { type: "addHierarchy", hierarchy: hierarchy("Driver"), area: "columns" },
      ]),
    );
    await write(def);
    const result = await runner.insertChart(def, "Line");
    expect("error" in result && result.error.message).toMatch(/at most 255 series/);
    expect(host.allCharts()).toEqual([]);
  });

  it("puts a second chart a little aside from the first", async () => {
    const def = report(byCity());
    await write(def, { row: 1, col: 1 });
    const first = await insert(def);
    const second = await insert(def, "Line");
    const at = new Map(host.allCharts().map((c) => [c.id, c.at]));
    expect(at.get(first.chart.id)).toEqual({ row: 1, col: 4 });
    expect(at.get(second.chart.id)).toEqual({ row: 3, col: 5 });
    expect(await savedCharts(def.id)).toEqual([first.chart, second.chart]);
  });
});

describe("a chart follows its report", () => {
  it("on Refresh, even when the pane's copy of the definition has no charts yet", async () => {
    const def = report(byCity());
    await write(def);
    const { chart } = await insert(def);
    // The pane may still hold the definition from before Insert Chart.
    const outcome = await write(def);
    expect(outcome.def.charts).toEqual([chart]);
    expect(await savedCharts(def.id)).toEqual([chart]);
  });

  it("on a layout change: new columns become new series, and the title follows", async () => {
    const def = report(byCity());
    await write(def);
    const { chart } = await insert(def);
    const changed = {
      ...def,
      layout: layoutOf(
        [
          { type: "addHierarchy", hierarchy: hierarchy("Year"), area: "columns" },
          { type: "addMeasure", measure: measure("Fare") },
        ],
        def.layout,
      ),
    };
    const outcome = await write(changed);
    const table = shownTable(chart.id);
    expect(table.series).toEqual(["2025 / Rides", "2025 / Fare", "2026 / Rides", "2026 / Fare"]);
    expect(table.title).toBe("Rides, Fare");
    expect(outcome.def.charts).toEqual([{ ...chart, title: "Rides, Fare" }]);
    expect(host.allCharts()[0].title).toBe("Rides, Fare");
  });

  it("keeps a title the user typed", async () => {
    const def = report(byCity());
    await write(def);
    const { chart } = await insert(def);
    host.retitleChart(chart.id, "Rides per city");
    await write({ ...def, layout: layoutOf([{ type: "addMeasure", measure: measure("Fare") }], def.layout) });
    expect(host.allCharts()[0].title).toBe("Rides per city");
    expect(shownTable(chart.id).series).toEqual(["Rides", "Fare"]);
  });

  it("on Refresh All: each report's charts follow their own report", async () => {
    const one = report(byCity());
    const two = report(
      layoutOf([
        { type: "addMeasure", measure: measure("Rides") },
        { type: "addHierarchy", hierarchy: hierarchy("Service Type") },
      ]),
    );
    await write(one, { row: 0, col: 0 });
    await write(two, { row: 0, col: 10 });
    const a = (await insert(one)).chart;
    const b = (await insert(two)).chart;
    // As Refresh All does: each report in turn, where it is.
    const tables = [];
    for (const def of [one, two]) {
      tables.push(written(await runner.run(def, { anchor: null, useSelection: false })).runtime.chart);
    }
    expect(shownTable(a.id)).toEqual(tables[0]);
    expect(shownTable(b.id)).toEqual(tables[1]);
    expect(shownTable(a.id).categories).not.toEqual(shownTable(b.id).categories);
    expect(await savedCharts(one.id)).toEqual([a]);
    expect(await savedCharts(two.id)).toEqual([b]);
  });

  it("drops a chart the user deleted in Excel, quietly", async () => {
    const def = report(byCity());
    await write(def);
    const kept = (await insert(def)).chart;
    const gone = (await insert(def, "Pie")).chart;
    host.removeChart(gone.id);
    const outcome = await write(def);
    expect(outcome.chartError).toBeUndefined();
    expect(outcome.def.charts).toEqual([kept]);
    expect(await savedCharts(def.id)).toEqual([kept]);

    host.removeChart(kept.id);
    const last = await write(def);
    expect(last.def).not.toHaveProperty("charts");
    expect(await savedCharts(def.id)).toBeUndefined();
  });

  it("is not lost by a run that started before it was inserted", async () => {
    const def = report(byCity());
    await write(def);
    const running = runner.run(def, { anchor: null, useSelection: false });
    const inserting = runner.insertChart(def, "Area");
    const [outcome, result] = await Promise.all([running, inserting]);
    if ("error" in result) throw new Error(result.error.message);
    written(outcome);
    expect(await savedCharts(def.id)).toEqual([result.chart]);
  });

  it("after the workbook is reopened, from the saved definition", async () => {
    const def = report(byCity());
    await write(def);
    const { chart } = await insert(def);
    // A new pane: a new runner, the definition read back from the workbook.
    runner = new ReportRunner({ api, host, confirmOverwrite: async () => true });
    const [loaded] = await host.loadReports();
    expect(runner.chartsOf(loaded)).toEqual([chart]);
    await write({ ...loaded, layout: layoutOf([{ type: "addMeasure", measure: measure("Fare") }], loaded.layout) });
    expect(shownTable(chart.id).series).toEqual(["Rides", "Fare"]);
  });

  it("keeps its last picture while the report has no fields", async () => {
    const def = report(byCity());
    await write(def);
    const { chart } = await insert(def);
    const before = shownTable(chart.id);
    const outcome = await runner.run({ ...def, layout: emptyLayout() }, { anchor: null, useSelection: false });
    expect(outcome.kind).toBe("empty");
    expect(shownTable(chart.id)).toEqual(before);
    expect(await savedCharts(def.id)).toEqual([chart]);
  });
});

describe("when a chart cannot follow", () => {
  class BusyHost extends PreviewHost {
    override async updateCharts(charts: ReportChart[], table: ChartTable): Promise<ChartsUpdate> {
      void table;
      throw new SheetError("Excel is busy.", `${charts.length} charts`);
    }
  }

  it("the report is still written, the chart kept, and the pane told", async () => {
    host = new BusyHost({ persist: false });
    runner = new ReportRunner({ api, host, confirmOverwrite: async () => true });
    const def = report(byCity());
    await write(def);
    const { chart } = await insert(def);
    const outcome = await write(def);
    expect(outcome.chartError?.message).toBe("The report was updated, but its chart could not be. Excel is busy.");
    expect(outcome.def.charts).toEqual([chart]);
    expect(await savedCharts(def.id)).toEqual([chart]);
  });

  it("nor when the report grows past what a chart can show", async () => {
    const def = report(byCity());
    await write(def);
    const { chart } = await insert(def);
    const before = shownTable(chart.id);
    const outcome = await write({
      ...def,
      layout: layoutOf([{ type: "addHierarchy", hierarchy: hierarchy("Driver"), area: "columns" }], def.layout),
    });
    expect(outcome.chartError?.message).toMatch(/^The chart was not updated\. A chart can show at most 255 series/);
    expect(shownTable(chart.id)).toEqual(before);
    expect(outcome.def.charts).toEqual([chart]);
  });
});

describe("Delete Chart, and deleting the report", () => {
  it("deletes the chart and takes it out of the definition", async () => {
    const def = report(byCity());
    await write(def);
    const one = (await insert(def)).chart;
    const two = (await insert(def, "Line")).chart;
    const result = await runner.deleteChart(def, one.id);
    if ("error" in result) throw new Error(result.error.message);
    expect(result.def.charts).toEqual([two]);
    expect(host.allCharts().map((c) => c.id)).toEqual([two.id]);
    expect(await savedCharts(def.id)).toEqual([two]);
    const last = await runner.deleteChart(result.def, two.id);
    if ("error" in last) throw new Error(last.error.message);
    expect(last.def).not.toHaveProperty("charts");
  });

  it("leaves the charts of a deleted report as they are, no longer updated", async () => {
    const def = report(byCity());
    await write(def);
    const { chart } = await insert(def);
    const before = shownTable(chart.id);
    // What the pane does for Delete Report and Convert to Formulas.
    await host.deleteReport(def.id);
    runner.forget(def.id);
    expect(host.allCharts().map((c) => c.id)).toEqual([chart.id]);
    // Another report written over the same cells does not take the chart over.
    await write(report(layoutOf([{ type: "addMeasure", measure: measure("Fare") }])));
    expect(shownTable(chart.id)).toEqual(before);
  });
});
