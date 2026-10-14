// The preview's stand-in for charts: it records them and what they show.

import { beforeEach, describe, expect, it } from "vitest";
import type { ChartTable } from "./charts";
import { SheetError } from "./host";
import { placeholderBlock } from "./layout";
import { PreviewHost } from "./preview";

const table = (title: string, first = 1): ChartTable => ({
  title,
  series: ["2025", "2026"],
  categories: ["Tehran", "Shiraz"],
  values: [
    [first, 2],
    [3, 4],
  ],
  numberFormats: [
    ["General", "General"],
    ["General", "General"],
  ],
});

let host: PreviewHost;

beforeEach(async () => {
  host = new PreviewHost({ persist: false });
  await host.writeReport("r1", { sheetId: "preview-sheet-1", sheetName: "Sheet1", row: 2, col: 1 }, placeholderBlock());
});

describe("PreviewHost charts", () => {
  it("adds a chart beside the report and shows it on the report's sheet", async () => {
    const chart = await host.addChart("r1", "c1", "ColumnClustered", table("Rides"), 0);
    expect(chart).toEqual({ id: "c1", type: "ColumnClustered", sheetId: "preview-sheet-1", title: "Rides" });
    const shown = host.snapshot().charts;
    expect(shown).toHaveLength(1);
    expect(shown[0]).toMatchObject({ id: "c1", reportId: "r1", type: "ColumnClustered", title: "Rides", at: { row: 2, col: 3 } });
    expect(shown[0].table.categories).toEqual(["Tehran", "Shiraz"]);
  });

  it("refuses a chart for a report that is not on a sheet", async () => {
    await expect(host.addChart("nowhere", "c1", "Pie", table("Rides"), 0)).rejects.toBeInstanceOf(SheetError);
    expect(host.allCharts()).toEqual([]);
  });

  it("updates what a chart shows, and its title unless the user changed it", async () => {
    const one = await host.addChart("r1", "c1", "Line", table("Rides"), 0);
    const two = await host.addChart("r1", "c2", "Pie", table("Rides"), 1);
    host.retitleChart("c2", "My own title");
    const update = await host.updateCharts([one, two], table("Fare", 99));
    expect(update.error).toBeUndefined();
    expect(update.charts).toEqual([
      { ...one, title: "Fare" },
      { ...two, title: "Rides" }, // the last title written stays on record
    ]);
    const byId = new Map(host.allCharts().map((c) => [c.id, c]));
    expect(byId.get("c1")?.title).toBe("Fare");
    expect(byId.get("c2")?.title).toBe("My own title");
    expect(byId.get("c2")?.table.values[0][0]).toBe(99);
  });

  it("drops a chart the user deleted", async () => {
    const one = await host.addChart("r1", "c1", "Area", table("Rides"), 0);
    const two = await host.addChart("r1", "c2", "BarClustered", table("Rides"), 1);
    host.removeChart("c1");
    const update = await host.updateCharts([one, two], table("Rides"));
    expect(update.charts.map((c) => c.id)).toEqual(["c2"]);
  });

  it("loses the charts of a deleted sheet, as Excel does", async () => {
    const chart = await host.addChart("r1", "c1", "Pie", table("Rides"), 0);
    host.deleteSheet();
    expect(host.allCharts()).toEqual([]);
    expect((await host.updateCharts([chart], table("Rides"))).charts).toEqual([]);
  });

  it("deletes a chart, and nothing when it is already gone", async () => {
    const chart = await host.addChart("r1", "c1", "Pie", table("Rides"), 0);
    await host.deleteChart(chart);
    expect(host.allCharts()).toEqual([]);
    await host.deleteChart(chart);
  });

  it("shows only the active sheet's charts", async () => {
    await host.addChart("r1", "c1", "Pie", table("Rides"), 0);
    await host.addDetailsSheet(["A"], [[1]]); // becomes the active sheet
    expect(host.snapshot().charts).toEqual([]);
    host.activate("preview-sheet-1");
    expect(host.snapshot().charts.map((c) => c.id)).toEqual(["c1"]);
  });

  it("forgets every chart for a new workbook", async () => {
    await host.addChart("r1", "c1", "Pie", table("Rides"), 0);
    host.reset();
    expect(host.allCharts()).toEqual([]);
  });
});
