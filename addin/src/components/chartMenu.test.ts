// The report menu's Insert Chart and Delete Chart, against the runner and the preview.

import { beforeEach, describe, expect, it } from "vitest";
import { createMockApi } from "../dev/mockApi";
import type { ReportDef } from "../sheet/host";
import { PreviewHost } from "../sheet/preview";
import { layoutReducer, newLayout, type LayoutAction } from "../state/layout";
import { newReportDef } from "../state/report";
import { ReportRunner, type UiError } from "../state/runner";
import type { Api } from "../api";
import { chartMenuItems, deleteChart, insertChart, type ChartActions } from "./chartMenu";
import type { MenuEntry } from "./Menu";

let api: Api;
let host: PreviewHost;
let runner: ReportRunner;
let def: ReportDef;
let log: { notes: (string | null)[]; errors: (UiError | null)[]; saved: ReportDef[]; refreshed: number };
let actions: ChartActions;

const run = () =>
  runner.run(def, { anchor: { sheetId: "preview-sheet-1", sheetName: "Sheet1", row: 0, col: 0 }, useSelection: false });

beforeEach(async () => {
  api = createMockApi({ delayMs: 0 });
  host = new PreviewHost({ persist: false });
  runner = new ReportRunner({ api, host, confirmOverwrite: async () => true });
  const fields = await api.fields("demo", "Sample", "Rides");
  const city = fields.dimensions.flatMap((d) => d.hierarchies).find((h) => h.caption === "City")!;
  const layout = [
    { type: "addMeasure", measure: fields.measures[0] },
    { type: "addHierarchy", hierarchy: city },
  ].reduce((l, a) => layoutReducer(l, a as LayoutAction), newLayout());
  def = { ...newReportDef("demo", "Sample", "Rides"), layout };
  log = { notes: [], errors: [], saved: [], refreshed: 0 };
  actions = {
    runner,
    refresh: async () => {
      log.refreshed++;
      return run();
    },
    setBusy: () => undefined,
    setError: (error) => log.errors.push(error),
    setNote: (note) => log.notes.push(note),
    saved: (saved) => log.saved.push(saved),
  };
});

function labels(items: MenuEntry[]): string[] {
  return items.map((item) => (typeof item === "string" ? item : "label" in item ? item.label : item.note));
}
function entry(items: MenuEntry[], label: string) {
  const found = items.find((item) => typeof item !== "string" && "label" in item && item.label === label);
  if (!found || typeof found === "string" || !("label" in found)) throw new Error("no entry " + label);
  return found;
}

describe("the report menu", () => {
  it("offers Windows' PivotChart types, and no Delete Chart without a chart", () => {
    const items = chartMenuItems(actions, def, { busy: false, canInsert: true });
    expect(labels(items)).toEqual(["Insert Chart"]);
    const insert = entry(items, "Insert Chart");
    expect(insert.disabled).toBe(false);
    expect("submenu" in insert && labels(insert.submenu)).toEqual([
      "Clustered Column",
      "Stacked Column",
      "Clustered Bar",
      "Line",
      "Pie",
      "Area",
    ]);
    expect(entry(chartMenuItems(actions, def, { busy: true, canInsert: true }), "Insert Chart").disabled).toBe(true);
    expect(entry(chartMenuItems(actions, def, { busy: false, canInsert: false }), "Insert Chart").disabled).toBe(true);
  });

  it("offers Delete Chart once there is a chart, and a list when there are several", async () => {
    await run();
    await insertChart(actions, def, "Pie");
    const one = entry(chartMenuItems(actions, def, { busy: false, canInsert: true }), "Delete Chart");
    expect("run" in one).toBe(true);
    await insertChart(actions, def, "Line");
    const several = entry(chartMenuItems(actions, def, { busy: false, canInsert: true }), "Delete Chart");
    expect("submenu" in several && labels(several.submenu)).toEqual(["Chart 1: Pie", "Chart 2: Line"]);
  });
});

describe("Insert Chart and Delete Chart", () => {
  it("refreshes a report not yet written in this session, then adds the chart", async () => {
    await insertChart(actions, def, "ColumnClustered");
    expect(log.refreshed).toBe(1);
    expect(host.allCharts()).toHaveLength(1);
    expect(log.saved[0].charts).toHaveLength(1);
    expect(log.notes[log.notes.length - 1]).toBe(
      "Clustered Column chart added beside the report. It updates whenever the report does.",
    );
    expect(log.errors.filter(Boolean)).toEqual([]);
  });

  it("says so when the report cannot be asked", async () => {
    actions.refresh = async () => null;
    await insertChart(actions, def, "Pie");
    expect(log.notes[log.notes.length - 1]).toBe("Choose a connection for this report first.");
    expect(host.allCharts()).toEqual([]);
  });

  it("shows why a chart was refused", async () => {
    await run();
    host.deleteSheet(); // the report loses its cells, but its last result is still known
    await insertChart(actions, def, "Pie");
    expect(log.refreshed).toBe(0);
    expect(log.errors[log.errors.length - 1]?.message).toBe("This report is no longer on a sheet.");
  });

  it("deletes a chart and saves the definition without it", async () => {
    await run();
    await insertChart(actions, def, "Area");
    const [chart] = log.saved[0].charts ?? [];
    await deleteChart(actions, log.saved[0], chart);
    expect(host.allCharts()).toEqual([]);
    expect(log.saved[1]).not.toHaveProperty("charts");
    expect(log.notes[log.notes.length - 1]).toBe("Chart deleted.");
  });
});
