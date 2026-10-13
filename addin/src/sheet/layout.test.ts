import { describe, expect, it } from "vitest";
import type { Grid, HeaderCell, HeaderKind, LineKind, MemberRef } from "../types";
import {
  GENERAL_FORMAT,
  MAX_CHUNK_CELLS,
  TEXT_FORMAT,
  buildBlock,
  chunkBlock,
  excelText,
  excelValues,
  expandedDescendants,
  memberAt,
  sliceMatrix,
} from "./layout";

const CAL = "[Date].[Calendar]";
const CITY = "[Geography].[City]";

function ref(
  hierarchy: string,
  caption: string,
  level: number,
  flags: { hasChildren?: boolean; expanded?: boolean } = {},
): MemberRef {
  return {
    uniqueName: `${hierarchy}.[${caption}]`,
    hierarchy,
    caption,
    level,
    hasChildren: flags.hasChildren ?? false,
    expanded: flags.expanded ?? false,
  };
}

function member(m: MemberRef, indent = 0): HeaderCell {
  return { text: m.caption, indent, kind: "member", member: m };
}
function blank(m: MemberRef): HeaderCell {
  return { text: "", indent: 0, kind: "blank", member: m };
}
function label(text: string, kind: HeaderKind): HeaderCell {
  return { text, indent: 0, kind, member: null };
}

const y2025 = ref(CAL, "2025", 1, { hasChildren: true, expanded: true });
const q1 = ref(CAL, "Q1", 2, { hasChildren: true, expanded: true });
const jan = ref(CAL, "Jan", 3);
const feb = ref(CAL, "Feb", 3);
const q2 = ref(CAL, "Q2", 2, { hasChildren: true });
const y2026 = ref(CAL, "2026", 1, { hasChildren: true });
const tehran = ref(CITY, "Tehran", 1);
const shiraz = ref(CITY, "Shiraz", 1);

/**
 * Rows: Calendar (2025 expanded to quarters, Q1 expanded to months), 2026, Grand Total.
 * Columns: City (Tehran, Shiraz, Grand Total) x measures (Rides, Fare).
 * Two page filters.
 */
function sampleGrid(): Grid {
  const cities: [HeaderCell, LineKind][] = [
    [member(tehran), "data"],
    [member(shiraz), "data"],
    [label("Grand Total", "grandTotal"), "grandTotal"],
  ];
  const cityRow: HeaderCell[] = [];
  const measureRow: HeaderCell[] = [];
  const colKinds: LineKind[] = [];
  for (const [cell, kind] of cities) {
    cityRow.push(cell, cell.member ? blank(cell.member) : label("", "grandTotal"));
    measureRow.push(label("Rides", "measure"), label("Fare", "measure"));
    colKinds.push(kind, kind);
  }
  const rowHeaders: HeaderCell[][] = [
    [member(y2025, 0)],
    [member(q1, 1)],
    [member(jan, 2)],
    [member(feb, 2)],
    [member(q2, 1)],
    [member(y2026, 0)],
    [label("Grand Total", "grandTotal")],
  ];
  const rowKinds: LineKind[] = ["data", "data", "data", "data", "data", "data", "grandTotal"];
  const values = rowHeaders.map((_, r) => [r * 10 + 1, r + 0.5, r * 10 + 2, null, r * 20 + 3, r + 1.5]);
  const cellFormats = rowHeaders.map(() => [1, 2, 1, 2, 1, 2]);
  return {
    filterRows: [
      { caption: "Service Type", value: "Ride" },
      { caption: "Segment", value: "(Multiple Items)" },
    ],
    rowFieldCaptions: ["Calendar"],
    colHeaderRows: [cityRow, measureRow],
    rowHeaders,
    values,
    formatPalette: ["General", "#,##0", "#,##0.00"],
    cellFormats,
    rowKinds,
    colKinds,
    stats: { rows: rowHeaders.length, columns: colKinds.length, elapsedMs: 12, mdx: "SELECT ..." },
  };
}

/** A grid with R row-header columns, C value columns and N data rows; no filters. */
function plainGrid(R: number, C: number, N: number, H = 1): Grid {
  const header = (): HeaderCell[] =>
    Array.from({ length: C }, (_, c) => label(`M${c}`, "measure"));
  return {
    filterRows: [],
    rowFieldCaptions: Array.from({ length: R }, (_, r) => `Field ${r}`),
    colHeaderRows: Array.from({ length: H }, header),
    rowHeaders: Array.from({ length: N }, (_, n) =>
      Array.from({ length: R }, (_, r) => label(`r${n}c${r}`, "member")),
    ),
    values: Array.from({ length: N }, (_, n) => Array.from({ length: C }, (_, c) => n * C + c)),
    formatPalette: ["General"],
    cellFormats: Array.from({ length: N }, () => new Array<number>(C).fill(0)),
    rowKinds: new Array<LineKind>(N).fill("data"),
    colKinds: new Array<LineKind>(C).fill("data"),
    stats: { rows: N, columns: C, elapsedMs: 1, mdx: "" },
  };
}

describe("block dimensions", () => {
  it("is F + 1 + H + N rows by R + C columns when there are filters", () => {
    const grid = sampleGrid();
    const block = buildBlock(grid);
    // F = 2, blank = 1, H = 2, N = 7; R = 1, C = 6.
    expect(block.rows).toBe(2 + 1 + 2 + 7);
    expect(block.cols).toBe(1 + 6);
    expect(block.shape).toEqual({
      filterRows: 2,
      headerRows: 2,
      headerCols: 1,
      dataRows: 7,
      dataCols: 6,
      headerTop: 3,
      dataTop: 5,
    });
  });

  it("is H + N rows with no filters: no blank row", () => {
    const block = buildBlock(plainGrid(2, 3, 4));
    expect(block.rows).toBe(1 + 4);
    expect(block.cols).toBe(2 + 3);
    expect(block.shape.headerTop).toBe(0);
    expect(block.shape.dataTop).toBe(1);
  });

  it("every row of values and numberFormats has exactly `cols` cells", () => {
    for (const grid of [sampleGrid(), plainGrid(1, 1, 0), plainGrid(3, 2, 5, 3)]) {
      const block = buildBlock(grid);
      expect(block.values).toHaveLength(block.rows);
      expect(block.numberFormats).toHaveLength(block.rows);
      for (let r = 0; r < block.rows; r++) {
        expect(block.values[r]).toHaveLength(block.cols);
        expect(block.numberFormats[r]).toHaveLength(block.cols);
      }
    }
  });

  it("counts several header rows and several row-header columns", () => {
    const block = buildBlock(plainGrid(3, 2, 5, 3));
    expect(block.rows).toBe(3 + 5);
    expect(block.cols).toBe(3 + 2);
  });

  it("stays two columns wide for a filter row even if the grid is one column wide", () => {
    const grid = plainGrid(1, 0, 0);
    grid.filterRows = [{ caption: "City", value: "Tehran" }];
    const block = buildBlock(grid);
    expect(block.cols).toBe(2);
    expect(block.values[0]).toEqual(["City", "Tehran"]);
  });
});

describe("filter rows and the blank row", () => {
  const block = buildBlock(sampleGrid());

  it("writes caption and value in the first two cells, as text", () => {
    expect(block.values[0].slice(0, 3)).toEqual(["Service Type", "Ride", ""]);
    expect(block.values[1].slice(0, 3)).toEqual(["Segment", "(Multiple Items)", ""]);
    expect(block.numberFormats[0].slice(0, 3)).toEqual([TEXT_FORMAT, TEXT_FORMAT, GENERAL_FORMAT]);
  });

  it("leaves exactly one empty row between the filters and the header", () => {
    expect(block.values[2].every((v) => v === "")).toBe(true);
    expect(block.numberFormats[2].every((f) => f === GENERAL_FORMAT)).toBe(true);
    expect(block.values[3][1]).toBe("Tehran");
  });

  it("makes the filter captions bold and filled, not the values", () => {
    const captions = { row: 0, col: 0, rows: 2, cols: 1 };
    expect(block.bold).toContainEqual(captions);
    expect(block.fill).toContainEqual(captions);
  });
});

describe("corner captions and headers", () => {
  it("puts the row field captions in the last header row only", () => {
    const block = buildBlock(sampleGrid());
    expect(block.values[3][0]).toBe("");
    expect(block.values[4][0]).toBe("Calendar");
  });

  it("puts each row field caption above its own column", () => {
    const block = buildBlock(plainGrid(3, 2, 2, 2));
    expect(block.values[0].slice(0, 3)).toEqual(["", "", ""]);
    expect(block.values[1].slice(0, 3)).toEqual(["Field 0", "Field 1", "Field 2"]);
    expect(block.values[1].slice(3)).toEqual(["M0", "M1"]);
  });

  it("writes the column header rows to the right of the corner", () => {
    const block = buildBlock(sampleGrid());
    expect(block.values[3].slice(1)).toEqual(["Tehran", "", "Shiraz", "", "Grand Total", ""]);
    expect(block.values[4].slice(1)).toEqual(["Rides", "Fare", "Rides", "Fare", "Rides", "Fare"]);
  });

  it("formats every header and corner cell as text", () => {
    const block = buildBlock(sampleGrid());
    for (const r of [3, 4]) {
      expect(block.numberFormats[r].every((f) => f === TEXT_FORMAT)).toBe(true);
    }
  });

  it("makes the header rows bold and filled across the whole width", () => {
    const block = buildBlock(sampleGrid());
    const header = { row: 3, col: 0, rows: 2, cols: 7 };
    expect(block.bold).toContainEqual(header);
    expect(block.fill).toContainEqual(header);
  });
});

describe("data rows", () => {
  const grid = sampleGrid();
  const block = buildBlock(grid);

  it("writes row headers then values", () => {
    expect(block.values[5]).toEqual(["2025", 1, 0.5, 2, "", 3, 1.5]);
    expect(block.values[11][0]).toBe("Grand Total");
    expect(block.values[11].slice(1)).toEqual([61, 6.5, 62, "", 123, 7.5]);
  });

  it("turns null into an empty cell that keeps its number format", () => {
    expect(block.values[5][4]).toBe("");
    expect(block.numberFormats[5][4]).toBe("#,##0.00");
  });

  it("maps cell formats through the palette", () => {
    expect(block.numberFormats[5]).toEqual([
      TEXT_FORMAT,
      "#,##0",
      "#,##0.00",
      "#,##0",
      "#,##0.00",
      "#,##0",
      "#,##0.00",
    ]);
  });

  it("falls back to General for a palette index that does not exist", () => {
    const odd = plainGrid(1, 2, 1);
    odd.cellFormats = [[7, -1]];
    expect(buildBlock(odd).numberFormats[1].slice(1)).toEqual([GENERAL_FORMAT, GENERAL_FORMAT]);
  });

  it("writes string and boolean values as they are, strings as text", () => {
    const odd = plainGrid(1, 3, 1);
    odd.values = [["n/a", true, 3]];
    const b = buildBlock(odd);
    expect(b.values[1].slice(1)).toEqual(["n/a", true, 3]);
    expect(b.numberFormats[1].slice(1)).toEqual([TEXT_FORMAT, GENERAL_FORMAT, GENERAL_FORMAT]);
  });

  it("makes total rows bold across the whole width, merging neighbours", () => {
    expect(block.bold).toContainEqual({ row: 11, col: 0, rows: 1, cols: 7 });

    const totals = plainGrid(1, 2, 6);
    totals.rowKinds = ["data", "subtotal", "grandTotal", "data", "subtotal", "data"];
    const b = buildBlock(totals);
    expect(b.bold).toEqual([
      { row: 0, col: 0, rows: 1, cols: 3 },
      { row: 2, col: 0, rows: 2, cols: 3 },
      { row: 5, col: 0, rows: 1, cols: 3 },
    ]);
  });

  it("does not fill data rows", () => {
    expect(block.fill).toHaveLength(2);
  });

  it("tolerates ragged rows from a faulty grid instead of throwing", () => {
    const odd = plainGrid(2, 2, 2);
    odd.values = [[1]];
    odd.cellFormats = [];
    odd.rowHeaders[1] = [];
    const b = buildBlock(odd);
    expect(b.values[1]).toEqual(["r0c0", "r0c1", 1, ""]);
    expect(b.values[2]).toEqual(["", "", "", ""]);
  });
});

describe("indent descriptors", () => {
  it("groups indented row headers by level and merges vertical runs", () => {
    const block = buildBlock(sampleGrid());
    // Data rows start at 5: 2025(0) Q1(1) Jan(2) Feb(2) Q2(1) 2026(0) Total(0).
    expect(block.indents).toEqual([
      {
        level: 1,
        rects: [
          { row: 6, col: 0, rows: 1, cols: 1 },
          { row: 9, col: 0, rows: 1, cols: 1 },
        ],
      },
      { level: 2, rects: [{ row: 7, col: 0, rows: 2, cols: 1 }] },
    ]);
  });

  it("has no indent groups when nothing is indented", () => {
    expect(buildBlock(plainGrid(2, 2, 3)).indents).toEqual([]);
  });

  it("covers indented column headers with horizontal runs", () => {
    const grid = plainGrid(1, 4, 1);
    grid.colHeaderRows = [
      [member(y2025, 0), member(q1, 1), member(q2, 1), member(y2026, 0)],
    ];
    expect(buildBlock(grid).indents).toEqual([
      { level: 1, rects: [{ row: 0, col: 2, rows: 1, cols: 2 }] },
    ]);
  });
});

describe("position to member lookup", () => {
  const block = buildBlock(sampleGrid());

  it("finds the member behind a row header", () => {
    expect(memberAt(block, 5, 0)).toEqual({ member: y2025, axis: "rows" });
    expect(memberAt(block, 7, 0)).toEqual({ member: jan, axis: "rows" });
  });

  it("finds the member behind a column header, including a blanked repeat", () => {
    expect(memberAt(block, 3, 1)).toEqual({ member: tehran, axis: "columns" });
    expect(memberAt(block, 3, 2)).toEqual({ member: tehran, axis: "columns" });
    expect(memberAt(block, 3, 3)).toEqual({ member: shiraz, axis: "columns" });
  });

  it("has nothing for totals, measures, corner, filters, values and cells outside", () => {
    expect(memberAt(block, 11, 0)).toBeNull(); // Grand Total row
    expect(memberAt(block, 3, 5)).toBeNull(); // Grand Total column
    expect(memberAt(block, 4, 1)).toBeNull(); // measure caption
    expect(memberAt(block, 4, 0)).toBeNull(); // corner
    expect(memberAt(block, 0, 0)).toBeNull(); // filter caption
    expect(memberAt(block, 5, 1)).toBeNull(); // a value
    expect(memberAt(block, 99, 99)).toBeNull();
  });

  it("holds one entry per header cell that has a member", () => {
    // Rows: 6 members. Columns: Tehran and Shiraz, each with one blanked repeat.
    expect(block.members.size).toBe(6 + 4);
  });
});

describe("no data (N = 0)", () => {
  const grid = plainGrid(2, 3, 0, 2);
  const block = buildBlock(grid);

  it("still writes the header rows", () => {
    expect(block.rows).toBe(2);
    expect(block.cols).toBe(5);
    expect(block.shape.dataRows).toBe(0);
    expect(block.values[1]).toEqual(["Field 0", "Field 1", "M0", "M1", "M2"]);
  });

  it("has header formatting and nothing else", () => {
    expect(block.bold).toEqual([{ row: 0, col: 0, rows: 2, cols: 5 }]);
    expect(block.fill).toEqual([{ row: 0, col: 0, rows: 2, cols: 5 }]);
    expect(block.indents).toEqual([]);
    expect(block.members.size).toBe(0);
  });

  it("keeps the filters and blank row above the header", () => {
    const filtered = plainGrid(1, 1, 0);
    filtered.filterRows = [{ caption: "City", value: "Tehran" }];
    const b = buildBlock(filtered);
    expect(b.rows).toBe(1 + 1 + 1);
    expect(b.values).toEqual([
      ["City", "Tehran"],
      ["", ""],
      ["Field 0", "M0"],
    ]);
  });
});

describe("chunk boundaries", () => {
  const cells = (rects: { rows: number; cols: number }[]) =>
    rects.reduce((sum, r) => sum + r.rows * r.cols, 0);

  it("uses 20,000 cells as the default limit", () => {
    expect(MAX_CHUNK_CELLS).toBe(20000);
  });

  it("is a single chunk at exactly the limit", () => {
    expect(chunkBlock({ rows: 200, cols: 100 })).toEqual([{ row: 0, col: 0, rows: 200, cols: 100 }]);
  });

  it("splits one row over the limit into a second chunk", () => {
    expect(chunkBlock({ rows: 201, cols: 100 })).toEqual([
      { row: 0, col: 0, rows: 200, cols: 100 },
      { row: 200, col: 0, rows: 1, cols: 100 },
    ]);
  });

  it("rounds the band down when the width does not divide the limit", () => {
    const chunks = chunkBlock({ rows: 10000, cols: 7 });
    // floor(20000 / 7) = 2857 rows = 19,999 cells per chunk.
    expect(chunks[0]).toEqual({ row: 0, col: 0, rows: 2857, cols: 7 });
    expect(chunks).toHaveLength(4);
    expect(chunks[3]).toEqual({ row: 8571, col: 0, rows: 1429, cols: 7 });
  });

  it("never exceeds the limit, and covers every row once, in order", () => {
    for (const [rows, cols] of [
      [1, 1],
      [300000, 1],
      [4321, 37],
      [15, 20000],
      [3, 19999],
      [1000, 300],
    ]) {
      const chunks = chunkBlock({ rows, cols });
      let next = 0;
      for (const chunk of chunks) {
        expect(chunk.rows * chunk.cols).toBeLessThanOrEqual(MAX_CHUNK_CELLS);
        expect(chunk.row).toBe(next);
        expect(chunk.col).toBe(0);
        expect(chunk.cols).toBe(cols);
        next += chunk.rows;
      }
      expect(next).toBe(rows);
      expect(cells(chunks)).toBe(rows * cols);
    }
  });

  it("cuts by columns only when a single row is over the limit", () => {
    const chunks = chunkBlock({ rows: 2, cols: 25 }, 10);
    expect(chunks).toEqual([
      { row: 0, col: 0, rows: 1, cols: 10 },
      { row: 0, col: 10, rows: 1, cols: 10 },
      { row: 0, col: 20, rows: 1, cols: 5 },
      { row: 1, col: 0, rows: 1, cols: 10 },
      { row: 1, col: 10, rows: 1, cols: 10 },
      { row: 1, col: 20, rows: 1, cols: 5 },
    ]);
  });

  it("returns nothing for an empty block", () => {
    expect(chunkBlock({ rows: 0, cols: 5 })).toEqual([]);
  });

  it("slices matching pieces out of the values matrix", () => {
    const block = buildBlock(plainGrid(1, 2, 5));
    const chunks = chunkBlock(block, 6); // 3 columns wide -> 2 rows per chunk
    expect(chunks.map((c) => c.rows)).toEqual([2, 2, 2]);
    const pieces = chunks.map((c) => sliceMatrix(block.values, c));
    expect(pieces[0]).toEqual([
      ["Field 0", "M0", "M1"],
      ["r0c0", 0, 1],
    ]);
    expect(pieces[2]).toEqual([
      ["r3c0", 6, 7],
      ["r4c0", 8, 9],
    ]);
    expect(([] as unknown[]).concat(...pieces)).toEqual(block.values);
  });
});

describe("values as Excel receives them", () => {
  it("guards text that Excel would read as a formula", () => {
    expect(excelText("=SUM(A1)")).toBe("'=SUM(A1)");
    expect(excelText("+1")).toBe("'+1");
    expect(excelText("-North")).toBe("'-North");
    expect(excelText("@home")).toBe("'@home");
    expect(excelText("'quoted")).toBe("''quoted");
    expect(excelText("Tehran")).toBe("Tehran");
    expect(excelText("2025")).toBe("2025");
    expect(excelText("")).toBe("");
  });

  it("guards header text but leaves numbers alone", () => {
    const grid = plainGrid(1, 1, 2);
    grid.rowHeaders = [[label("=cmd", "member")], [label("-5", "member")]];
    grid.values = [[-5], [7]];
    const block = buildBlock(grid);
    const out = excelValues(block, false);
    expect(out[1]).toEqual(["'=cmd", -5]);
    expect(out[2]).toEqual(["'-5", 7]);
    expect(block.values[1][0]).toBe("=cmd");
  });

  it("indents with spaces only when asked to", () => {
    const block = buildBlock(sampleGrid());
    expect(excelValues(block, false)[7][0]).toBe("Jan");
    const spaced = excelValues(block, true);
    expect(spaced[5][0]).toBe("2025");
    expect(spaced[6][0]).toBe("   Q1");
    expect(spaced[7][0]).toBe("      Jan");
    expect(spaced[11][0]).toBe("Grand Total");
  });
});

describe("expanded descendants", () => {
  const headers = () => {
    const grid = sampleGrid();
    return { rowHeaders: grid.rowHeaders, colHeaderRows: grid.colHeaderRows };
  };

  it("finds the expanded members shown under a member", () => {
    expect(expandedDescendants(headers(), "rows", y2025)).toEqual([q1.uniqueName]);
  });

  it("finds none under a leaf or an unexpanded member", () => {
    expect(expandedDescendants(headers(), "rows", q2)).toEqual([]);
    expect(expandedDescendants(headers(), "rows", y2026)).toEqual([]);
  });

  it("does not report siblings or members after the subtree ends", () => {
    const h = headers();
    const y2026open = { ...y2026, expanded: true };
    h.rowHeaders[5] = [member(y2026open)];
    expect(expandedDescendants(h, "rows", y2025)).toEqual([q1.uniqueName]);
    expect(expandedDescendants(h, "rows", q1)).toEqual([]);
  });

  it("follows a member repeated under several outer members", () => {
    const outerA = ref(CITY, "Tehran", 1);
    const outerB = ref(CITY, "Shiraz", 1);
    const q3 = ref(CAL, "Q3", 2, { hasChildren: true, expanded: true });
    const rowHeaders: HeaderCell[][] = [
      [member(outerA), member(y2025)],
      [blank(outerA), member(q1, 1)],
      [member(outerB), member(y2025)],
      [blank(outerB), member(q3, 1)],
      [label("Grand Total", "grandTotal"), label("", "grandTotal")],
    ];
    const found = expandedDescendants({ rowHeaders, colHeaderRows: [] }, "rows", y2025);
    expect(found.sort()).toEqual([q1.uniqueName, q3.uniqueName].sort());
  });

  it("works along a column header row", () => {
    const colHeaderRows: HeaderCell[][] = [
      [member(y2025), blank(y2025), member(q1, 1), member(jan, 2), member(y2026)],
    ];
    expect(expandedDescendants({ rowHeaders: [], colHeaderRows }, "columns", y2025)).toEqual([
      q1.uniqueName,
    ]);
  });
});
