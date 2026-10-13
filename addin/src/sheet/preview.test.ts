import { beforeEach, describe, expect, it } from "vitest";
import type { SheetClick } from "./host";
import { FILTER_FORMAT, buildBlock, placeholderBlock } from "./layout";
import { PreviewHost, formatForPreview } from "./preview";
import type { Grid } from "../types";

const SHEET = { sheetId: "preview-sheet-1", sheetName: "Sheet1" };

/** One filter row (Year | All), a blank row, a header row and two data rows. */
function grid(): Grid {
  const cell = (text: string) => ({ text, indent: 0, kind: "measure" as const, member: null });
  return {
    filterRows: [{ caption: "Year", value: "All" }],
    rowFieldCaptions: ["City"],
    colHeaderRows: [[cell("Rides")]],
    rowHeaders: [[cell("Tehran")], [cell("Shiraz")]],
    values: [[10], [20]],
    formatPalette: ["General"],
    cellFormats: [[0], [0]],
    rowKinds: ["data", "data"],
    colKinds: ["data"],
    stats: { rows: 2, columns: 1, elapsedMs: 1, mdx: "" },
  };
}
const fields = { filters: 1, rows: 1, columns: [] };

const settle = () => new Promise((resolve) => setTimeout(resolve, 5));

let host: PreviewHost;
let clicks: SheetClick[];

beforeEach(() => {
  host = new PreviewHost({ persist: false });
  clicks = [];
});

describe("clicks", () => {
  it("reports a click inside a report with the cell's offsets, block size and plain value", async () => {
    host.onCellClicked((click) => clicks.push(click));
    await host.writeReport("r1", { ...SHEET, row: 2, col: 1 }, buildBlock(grid(), fields));
    host.click(2, 2);
    expect(clicks).toEqual([]); // delivered after the selection, as Excel does
    await settle();
    expect(clicks).toEqual([{ report: { id: "r1", row: 0, col: 1, rows: 5, cols: 2 }, value: "All" }]);
    // The value has no arrow, though the cell shows one.
    const cell = host.snapshot().cells.get("2,2");
    expect(cell).toMatchObject({ value: "All", format: FILTER_FORMAT });
  });

  it("moves the selection first, then reports the click", async () => {
    const order: string[] = [];
    host.onSelectionChanged((selection) => order.push(`select ${selection.cell.row},${selection.cell.col}`));
    host.onCellClicked((click) => order.push(`click ${click.report.row},${click.report.col}`));
    await host.writeReport("r1", { ...SHEET, row: 0, col: 0 }, buildBlock(grid(), fields));
    host.click(3, 0);
    await settle();
    expect(order).toEqual(["select 3,0", "click 3,0"]);
  });

  it("reports a click on the cell that is already selected, which changes no selection", async () => {
    let selections = 0;
    host.onSelectionChanged(() => selections++);
    host.onCellClicked((click) => clicks.push(click));
    await host.writeReport("r1", { ...SHEET, row: 0, col: 0 }, buildBlock(grid(), fields));
    host.click(0, 1);
    await settle();
    host.click(0, 1);
    await settle();
    expect(selections).toBe(1);
    expect(clicks).toHaveLength(2);
  });

  it("does not report keyboard movement as a click", async () => {
    host.onCellClicked((click) => clicks.push(click));
    await host.writeReport("r1", { ...SHEET, row: 0, col: 0 }, buildBlock(grid(), fields));
    host.select(0, 1);
    await settle();
    expect(clicks).toEqual([]);
    expect((await host.getSelection()).report).toMatchObject({ id: "r1", row: 0, col: 1 });
  });

  it("says nothing about clicks outside every report", async () => {
    host.onCellClicked((click) => clicks.push(click));
    await host.writeReport("r1", { ...SHEET, row: 0, col: 0 }, buildBlock(grid(), fields));
    host.click(20, 5);
    await settle();
    expect(clicks).toEqual([]);
  });

  it("reports a click on a placeholder as a click in that report", async () => {
    host.onCellClicked((click) => clicks.push(click));
    await host.writeReport("r1", { ...SHEET, row: 4, col: 4 }, placeholderBlock());
    host.click(4, 4);
    await settle();
    expect(clicks[0].report).toEqual({ id: "r1", row: 0, col: 0, rows: 1, cols: 1 });
  });

  it("stops reporting after unsubscribing", async () => {
    const stop = host.onCellClicked((click) => clicks.push(click));
    await host.writeReport("r1", { ...SHEET, row: 0, col: 0 }, buildBlock(grid(), fields));
    stop();
    host.click(0, 1);
    await settle();
    expect(clicks).toEqual([]);
  });
});

describe("how a filter cell is drawn", () => {
  it("shows the arrow after the text, from the format alone", () => {
    expect(formatForPreview("Years", FILTER_FORMAT)).toBe("Years ▾");
    expect(formatForPreview("(Multiple Items)", FILTER_FORMAT)).toBe("(Multiple Items) ▾");
    expect(formatForPreview("Years", "@")).toBe("Years");
    expect(formatForPreview("", FILTER_FORMAT)).toBe("");
  });

  it("leaves numbers and other formats as they were", () => {
    expect(formatForPreview(1234.5, "#,##0.00")).toBe("1,234.50");
    expect(formatForPreview(0.125, "0.0%")).toBe("12.5%");
    expect(formatForPreview(7, "General")).toBe("7");
    expect(formatForPreview(true, "General")).toBe("TRUE");
  });
});

describe("deleting the sheet", () => {
  it("leaves the definitions but no cells, and moves to a new sheet", async () => {
    const def = {
      id: "r1",
      connectionId: "demo",
      catalog: "Sample",
      cube: "Rides",
      layout: { rows: [], columns: [], values: [], filters: [], options: { grandTotals: true, subtotals: false, nonEmpty: true } },
    };
    await host.writeReport("r1", { ...SHEET, row: 0, col: 0 }, buildBlock(grid(), fields));
    await host.saveReport(def);
    const seen: string[] = [];
    host.onSelectionChanged((selection) => seen.push(selection.cell.sheetName));

    host.deleteSheet();
    expect(await host.getReportExtent("r1")).toBeNull();
    expect(await host.loadReports()).toEqual([def]);
    expect(host.snapshot().cells.size).toBe(0);
    expect(seen).toEqual(["Sheet2"]);
    const selection = await host.getSelection();
    expect(selection.cell).toMatchObject({ sheetId: "preview-sheet-2", sheetName: "Sheet2", row: 0, col: 0 });
    expect(selection.report).toBeNull();
  });

  it("does not deliver a click that was waiting when the sheet went", async () => {
    const clicksSeen: SheetClick[] = [];
    host.onCellClicked((click) => clicksSeen.push(click));
    await host.writeReport("r1", { ...SHEET, row: 0, col: 0 }, buildBlock(grid(), fields));
    host.click(0, 1);
    host.deleteSheet();
    await settle();
    expect(clicksSeen).toEqual([]);
  });
});
