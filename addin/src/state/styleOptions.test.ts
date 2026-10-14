import { describe, expect, it } from "vitest";
import { emptyLayout, layoutReducer, newLayout, type LayoutAction } from "./layout";
import { parseLayout } from "./report";
import type { Layout } from "../types";

const run = (actions: LayoutAction[], start: Layout = newLayout()): Layout => actions.reduce(layoutReducer, start);

describe("choosing a PivotTable style", () => {
  it("stores a style from the list, and nothing for the default", () => {
    const medium = run([{ type: "setStyle", style: "Medium 2" }]);
    expect(medium.options.style).toBe("Medium 2");
    const back = layoutReducer(medium, { type: "setStyle", style: "Light 16" });
    expect("style" in back.options).toBe(false);
    expect("style" in layoutReducer(medium, { type: "setStyle", style: null }).options).toBe(false);
  });

  it("changes nothing for the style already in use, or one it does not know", () => {
    const start = newLayout();
    expect(layoutReducer(start, { type: "setStyle", style: "Light 16" })).toBe(start);
    expect(layoutReducer(start, { type: "setStyle", style: "Medium 77" })).toBe(start);
    const dark = run([{ type: "setStyle", style: "Dark 2" }]);
    expect(layoutReducer(dark, { type: "setStyle", style: "Dark 2" })).toBe(dark);
  });

  it("stores a Style Option only while it differs from Excel's default", () => {
    const banded = run([
      { type: "setStyleOption", option: "bandedRows", value: true },
      { type: "setStyleOption", option: "rowHeaders", value: false },
    ]);
    expect(banded.options).toMatchObject({ bandedRows: true, rowHeaders: false });
    const back = run(
      [
        { type: "setStyleOption", option: "bandedRows", value: false },
        { type: "setStyleOption", option: "rowHeaders", value: true },
      ],
      banded,
    );
    expect(back.options).toEqual(newLayout().options);
  });

  it("changes nothing when an option already has that value", () => {
    const start = newLayout();
    expect(layoutReducer(start, { type: "setStyleOption", option: "columnHeaders", value: true })).toBe(start);
    expect(layoutReducer(start, { type: "setStyleOption", option: "bandedColumns", value: false })).toBe(start);
  });
});

describe("reading the style back", () => {
  it("leaves a report without a style as it was", () => {
    expect(parseLayout({ options: { grandTotals: true, subtotals: false, nonEmpty: true } }).options).toEqual(
      emptyLayout().options,
    );
  });

  it("keeps a known style and boolean options", () => {
    const layout = parseLayout({
      options: { style: "Medium 9", bandedRows: true, bandedColumns: false, rowHeaders: false, columnHeaders: true },
    });
    expect(layout.options).toMatchObject({
      style: "Medium 9",
      bandedRows: true,
      bandedColumns: false,
      rowHeaders: false,
      columnHeaders: true,
    });
  });

  it("drops a style it does not know and options that are not booleans", () => {
    const layout = parseLayout({ options: { style: "Fancy", bandedRows: "yes", rowHeaders: 0, columnHeaders: null } });
    for (const key of ["style", "bandedRows", "rowHeaders", "columnHeaders"]) expect(key in layout.options).toBe(false);
    expect("style" in parseLayout({ options: { style: 16 } }).options).toBe(false);
  });
});
