import { describe, expect, it } from "vitest";
import type { Grid, HeaderCell, HeaderKind, LineKind, MemberRef } from "../types";
import {
  FILTERED_FORMAT,
  FILTER_FORMAT,
  GENERAL_FORMAT,
  MAX_CHUNK_CELLS,
  PLACEHOLDER_TEXT,
  TEXT_FORMAT,
  blockFields,
  buildBlock,
  chunkBlock,
  columnHeaderAt,
  detailsCells,
  detailsMembers,
  expandableMembers,
  excelText,
  excelValues,
  expandedDescendants,
  fieldCellAt,
  guessFieldCell,
  isTextFormat,
  memberAt,
  placeholderBlock,
  sliceMatrix,
  valueFieldIndex,
  type BlockFields,
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
    colValueIndex: colKinds.map((_, c) => c % 2),
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
    colValueIndex: new Array<number>(C).fill(0),
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

describe("filter cells: page filters", () => {
  // sampleGrid has two filter rows; the layout that asked for it has two filters.
  const fields: BlockFields = { filters: 2, rows: 1, columns: ["City"] };
  const block = buildBlock(sampleGrid(), fields);

  it("maps the caption and the value cell of each filter row to that filter", () => {
    expect(fieldCellAt(block, 0, 0)).toEqual({ area: "filters", index: 0 });
    expect(fieldCellAt(block, 0, 1)).toEqual({ area: "filters", index: 0 });
    expect(fieldCellAt(block, 1, 0)).toEqual({ area: "filters", index: 1 });
    expect(fieldCellAt(block, 1, 1)).toEqual({ area: "filters", index: 1 });
    expect(fieldCellAt(block, 0, 2)).toBeNull();
    expect(fieldCellAt(block, 2, 0)).toBeNull(); // the blank row
  });

  it("shows the arrow on the value cell only, and leaves both values as they were", () => {
    expect(block.numberFormats[0].slice(0, 3)).toEqual([TEXT_FORMAT, FILTER_FORMAT, GENERAL_FORMAT]);
    expect(block.numberFormats[1].slice(0, 2)).toEqual([TEXT_FORMAT, FILTER_FORMAT]);
    expect(block.values[0].slice(0, 2)).toEqual(["Service Type", "Ride"]);
    expect(block.values[1].slice(0, 2)).toEqual(["Segment", "(Multiple Items)"]);
  });

  it("marks none when the grid's filter rows do not match the layout's filters", () => {
    const odd = buildBlock(sampleGrid(), { filters: 1, rows: 1, columns: ["City"] });
    expect(fieldCellAt(odd, 0, 1)).toBeNull();
    expect(odd.numberFormats[0][1]).toBe(TEXT_FORMAT);
    // The other kinds do not depend on the filters.
    expect(fieldCellAt(odd, 4, 0)).toEqual({ area: "rows", index: 0 });
  });

  it("has no filter cells at all when no layout is given", () => {
    const bare = buildBlock(sampleGrid());
    expect(bare.fieldCells.size).toBe(0);
    expect(bare.numberFormats[0][1]).toBe(TEXT_FORMAT);
  });
});

describe("filter cells: row-field captions", () => {
  it("maps each corner cell of the last header row to its row field, with the arrow", () => {
    const block = buildBlock(plainGrid(3, 2, 2, 2), { filters: 0, rows: 3, columns: ["City"] });
    for (let r = 0; r < 3; r++) {
      expect(fieldCellAt(block, 1, r)).toEqual({ area: "rows", index: r });
      expect(block.numberFormats[1][r]).toBe(FILTER_FORMAT);
    }
    expect(block.values[1].slice(0, 3)).toEqual(["Field 0", "Field 1", "Field 2"]);
    expect(fieldCellAt(block, 1, 3)).toBeNull(); // a measure caption
  });

  it("marks no corner cell when there are no row fields (the grid still has one corner column)", () => {
    const grid = plainGrid(1, 2, 1);
    grid.rowFieldCaptions = [""];
    const block = buildBlock(grid, { filters: 0, rows: 0, columns: [] });
    expect(fieldCellAt(block, 0, 0)).toBeNull();
    expect(block.numberFormats[0][0]).toBe(TEXT_FORMAT);
    expect(block.fieldCells.size).toBe(0);
  });

  it("does not mark the data rows' own header cells", () => {
    const block = buildBlock(sampleGrid(), { filters: 2, rows: 1, columns: ["City"] });
    expect(fieldCellAt(block, 5, 0)).toBeNull();
    expect(block.numberFormats[5][0]).toBe(TEXT_FORMAT);
  });
});

describe("filter cells: column-field captions", () => {
  it("writes the field's caption in the corner cell of its header row, with the arrow", () => {
    // Header rows: 3 = City, 4 = measures (whose corner is the row field's caption).
    const block = buildBlock(sampleGrid(), { filters: 2, rows: 1, columns: ["City"] });
    expect(block.values[3][0]).toBe("City");
    expect(block.numberFormats[3][0]).toBe(FILTER_FORMAT);
    expect(fieldCellAt(block, 3, 0)).toEqual({ area: "columns", index: 0 });
    // The member headers next to it are not filter cells.
    expect(fieldCellAt(block, 3, 1)).toBeNull();
    expect(block.values[4][0]).toBe("Calendar");
    expect(fieldCellAt(block, 4, 0)).toEqual({ area: "rows", index: 0 });
  });

  it("uses the corner column nearest the data when there are several row fields", () => {
    const grid = plainGrid(3, 2, 2, 3); // two column fields, then the measures
    const block = buildBlock(grid, { filters: 0, rows: 3, columns: ["Year", "City"] });
    expect(block.values[0].slice(0, 3)).toEqual(["", "", "Year"]);
    expect(block.values[1].slice(0, 3)).toEqual(["", "", "City"]);
    expect(fieldCellAt(block, 0, 2)).toEqual({ area: "columns", index: 0 });
    expect(fieldCellAt(block, 1, 2)).toEqual({ area: "columns", index: 1 });
    expect(fieldCellAt(block, 0, 0)).toBeNull();
    expect(block.numberFormats[0].slice(0, 3)).toEqual([TEXT_FORMAT, TEXT_FORMAT, FILTER_FORMAT]);
    // The row captions keep the last header row.
    expect(block.values[2].slice(0, 3)).toEqual(["Field 0", "Field 1", "Field 2"]);
  });

  it("puts it in the only corner column when there are no row fields", () => {
    const grid = plainGrid(1, 2, 1, 2);
    grid.rowFieldCaptions = [""];
    const block = buildBlock(grid, { filters: 0, rows: 0, columns: ["City"] });
    expect(block.values[0][0]).toBe("City");
    expect(fieldCellAt(block, 0, 0)).toEqual({ area: "columns", index: 0 });
    expect(fieldCellAt(block, 1, 0)).toBeNull(); // the empty row-caption corner
    expect(block.fieldCells.size).toBe(1);
  });

  it("adds nothing when there are no column fields", () => {
    const block = buildBlock(plainGrid(2, 2, 2), { filters: 0, rows: 2, columns: [] });
    expect(Array.from(block.fieldCells.values()).every((cell) => cell.area === "rows")).toBe(true);
    expect(block.fieldCells.size).toBe(2);
  });

  it("gives way to the row captions if a grid has no separate measures row", () => {
    // One header row only, though the layout has a column field: the corner is the
    // row field's. The column field is still reachable from the pane.
    const block = buildBlock(plainGrid(1, 2, 1, 1), { filters: 0, rows: 1, columns: ["City"] });
    expect(block.values[0][0]).toBe("Field 0");
    expect(fieldCellAt(block, 0, 0)).toEqual({ area: "rows", index: 0 });
  });

  it("is still there with no data rows", () => {
    const block = buildBlock(plainGrid(2, 3, 0, 2), { filters: 0, rows: 2, columns: ["City"] });
    expect(block.rows).toBe(2);
    expect(block.values[0].slice(0, 2)).toEqual(["", "City"]);
    expect(fieldCellAt(block, 0, 1)).toEqual({ area: "columns", index: 0 });
    expect(fieldCellAt(block, 1, 0)).toEqual({ area: "rows", index: 0 });
    expect(fieldCellAt(block, 1, 1)).toEqual({ area: "rows", index: 1 });
    expect(block.fieldCells.size).toBe(3);
  });

  it("keeps the header bold and filled across the caption cells", () => {
    const block = buildBlock(sampleGrid(), { filters: 2, rows: 1, columns: ["City"] });
    expect(block.bold).toContainEqual({ row: 3, col: 0, rows: 2, cols: 7 });
    expect(block.fill).toContainEqual({ row: 3, col: 0, rows: 2, cols: 7 });
  });
});

describe("filter cells as Excel receives them", () => {
  it("forces every filter cell to text, so that 2025 keeps its arrow", () => {
    const grid = plainGrid(1, 1, 1, 2);
    grid.filterRows = [{ caption: "Years", value: "2025" }];
    grid.rowFieldCaptions = ["2024"];
    const block = buildBlock(grid, { filters: 1, rows: 1, columns: ["=Risky"] });
    const out = excelValues(block, false);
    expect(out[0]).toEqual(["Years", "'2025"]); // caption cell is plain text; value cell is forced
    expect(out[2][0]).toBe("'=Risky");
    expect(out[3][0]).toBe("'2024");
    // The block itself keeps the real values.
    expect(block.values[0][1]).toBe("2025");
  });

  it("leaves ordinary header text and numbers as before", () => {
    const block = buildBlock(sampleGrid(), { filters: 2, rows: 1, columns: ["City"] });
    const out = excelValues(block, false);
    expect(out[5][0]).toBe("2025"); // a member caption in a Text cell
    expect(out[5][1]).toBe(1);
    expect(out[3][1]).toBe("Tehran");
  });

  it("counts both text formats as text, for the fallback that drops custom formats", () => {
    expect(isTextFormat(TEXT_FORMAT)).toBe(true);
    expect(isTextFormat(FILTER_FORMAT)).toBe(true);
    expect(isTextFormat(GENERAL_FORMAT)).toBe(false);
    expect(isTextFormat("#,##0")).toBe(false);
    expect(FILTER_FORMAT.startsWith("@")).toBe(true);
  });
});

describe("blockFields", () => {
  it("takes the counts and the column captions from a layout", () => {
    const field = (caption: string) => ({
      hierarchy: `[D].[${caption}]`,
      caption,
      allMember: null,
      expanded: [],
      members: [],
      memberCaptions: [],
    });
    expect(
      blockFields({ filters: [field("Years")], rows: [field("A"), field("B")], columns: [field("City")] }),
    ).toEqual({
      filters: 1,
      rows: 2,
      columns: ["City"],
      rowsFiltered: [false, false],
      columnsFiltered: [false],
      valueFormats: [],
      valueIcons: [],
      repeatLabels: false,
      emptyCellText: "",
    });
    expect(blockFields({ filters: [], rows: [], columns: [] })).toMatchObject({ filters: 0, rows: 0, columns: [] });
  });

  it("marks a KPI's Status and Trend for icons, and nothing else", () => {
    const values = [
      { measure: "[Measures].[S]", caption: "S Value", kpi: { name: "S", part: "value" as const } },
      { measure: "[Measures].[S]", caption: "S Status", kpi: { name: "S", part: "status" as const } },
      { measure: "[Measures].[R]", caption: "R" },
    ];
    expect(blockFields({ filters: [], rows: [], columns: [], values }).valueIcons).toEqual([null, "status", null]);
  });
});

describe("KPI icons", () => {
  it("covers every data line of each Status or Trend column", () => {
    // Columns alternate value fields 0 and 1; field 1 is a KPI's Status.
    const block = buildBlock(sampleGrid(), { filters: 2, rows: 1, columns: ["City"], valueIcons: [null, "status"] });
    const { dataTop, dataRows, headerCols } = block.shape;
    expect(block.icons).toEqual([1, 3, 5].map((c) => ({
      kind: "status",
      rect: { row: dataTop, col: headerCols + c, rows: dataRows, cols: 1 },
    })));
    expect(buildBlock(sampleGrid(), { filters: 2, rows: 1, columns: ["City"] }).icons).toEqual([]);
  });
});

describe("placeholder block", () => {
  const block = placeholderBlock();

  it("is one quiet text cell", () => {
    expect(block.rows).toBe(1);
    expect(block.cols).toBe(1);
    expect(block.values).toEqual([[PLACEHOLDER_TEXT]]);
    expect(block.numberFormats).toEqual([[TEXT_FORMAT]]);
    expect(block.muted).toEqual([{ row: 0, col: 0, rows: 1, cols: 1 }]);
    expect(block.bold).toEqual([]);
    expect(block.fill).toEqual([]);
  });

  it("has no filter cells, no members, and leaves the column width alone", () => {
    expect(block.fieldCells.size).toBe(0);
    expect(block.members.size).toBe(0);
    expect(block.autofit).toBe(false);
  });

  it("says what to do, in words Excel will not read as a formula", () => {
    expect(PLACEHOLDER_TEXT).toMatch(/Cube Explorer/);
    expect(excelValues(block, false)).toEqual([[PLACEHOLDER_TEXT]]);
  });

  it("is a fresh object each time", () => {
    expect(placeholderBlock()).not.toBe(block);
    expect(placeholderBlock().values).not.toBe(block.values);
  });
});

describe("column header cells, for the selected-cell action", () => {
  // sampleGrid: 2 filter rows, blank row, header rows 3 (City) and 4 (measures); 1 corner column.
  const shape = buildBlock(sampleGrid()).shape;

  it("finds the column field for the member and total cells of its header row", () => {
    expect(columnHeaderAt(shape, 3, 1, 1)).toBe(0);
    expect(columnHeaderAt(shape, 3, 6, 1)).toBe(0); // Grand Total column
  });

  it("is null for the corner, the measures row, the data and outside", () => {
    expect(columnHeaderAt(shape, 3, 0, 1)).toBeNull();
    expect(columnHeaderAt(shape, 4, 1, 1)).toBeNull();
    expect(columnHeaderAt(shape, 5, 1, 1)).toBeNull();
    expect(columnHeaderAt(shape, 3, 7, 1)).toBeNull();
    expect(columnHeaderAt(shape, 0, 1, 1)).toBeNull();
    expect(columnHeaderAt(buildBlock(plainGrid(1, 2, 2)).shape, 0, 1, 0)).toBeNull();
  });
});

describe("guessing the filter cells from the layout alone", () => {
  const field = (caption: string) => ({
    hierarchy: `[D].[${caption}]`,
    caption,
    allMember: null,
    expanded: [],
    members: [],
    memberCaptions: [],
  });

  it("agrees with the block for the shapes the helper produces", () => {
    const cases: { layout: { filters: string[]; rows: string[]; columns: string[] }; grid: Grid }[] = [];
    const add = (filters: string[], rows: string[], columns: string[], N: number) => {
      const grid = plainGrid(Math.max(1, rows.length), 2, N, columns.length + 1);
      grid.filterRows = filters.map((caption) => ({ caption, value: "All" }));
      grid.rowFieldCaptions = rows.length > 0 ? rows : [""];
      cases.push({ layout: { filters, rows, columns }, grid });
    };
    add([], ["A"], [], 3);
    add(["F1", "F2"], ["A", "B"], ["C1"], 3);
    add(["F1"], [], ["C1", "C2"], 1);
    add([], [], [], 1);
    add(["F1"], ["A"], ["C1"], 0);

    for (const { layout, grid } of cases) {
      const full = {
        filters: layout.filters.map(field),
        rows: layout.rows.map(field),
        columns: layout.columns.map(field),
      };
      const block = buildBlock(grid, blockFields(full));
      for (let r = 0; r < block.rows + 1; r++) {
        for (let c = 0; c < block.cols + 1; c++) {
          expect(guessFieldCell(full, r, c)).toEqual(fieldCellAt(block, r, c));
        }
      }
    }
  });

  it("is null outside the block's top-left region and for negative positions", () => {
    const layout = { filters: [field("F")], rows: [field("A")], columns: [] };
    expect(guessFieldCell(layout, 0, 2)).toBeNull();
    expect(guessFieldCell(layout, 1, 0)).toBeNull(); // blank row
    expect(guessFieldCell(layout, 3, 0)).toBeNull(); // first data row
    expect(guessFieldCell(layout, -1, 0)).toBeNull();
  });
});

describe("a filtered field's caption cell", () => {
  it("gets the solid arrow when the row or column field is filtered", () => {
    const fields: BlockFields = {
      filters: 2,
      rows: 1,
      columns: ["City"],
      rowsFiltered: [true],
      columnsFiltered: [false],
    };
    const block = buildBlock(sampleGrid(), fields);
    expect(block.numberFormats[4][0]).toBe(FILTERED_FORMAT); // row field caption
    expect(block.numberFormats[3][0]).toBe(FILTER_FORMAT); // column field caption
    const both = buildBlock(sampleGrid(), { ...fields, columnsFiltered: [true] });
    expect(both.numberFormats[3][0]).toBe(FILTERED_FORMAT);
  });

  it("keeps the plain arrow on page filters, whose value cell says what is chosen", () => {
    const block = buildBlock(sampleGrid(), { filters: 2, rows: 1, columns: ["City"], rowsFiltered: [true] });
    expect(block.numberFormats[0][1]).toBe(FILTER_FORMAT);
  });

  it("is still a filter cell, forced to text, and plain text in the format fallback", () => {
    const block = buildBlock(sampleGrid(), { filters: 2, rows: 1, columns: ["City"], rowsFiltered: [true] });
    expect(fieldCellAt(block, 4, 0)).toEqual({ area: "rows", index: 0 });
    expect(excelValues(block, false)[4][0]).toBe("'Calendar");
    expect(isTextFormat(FILTERED_FORMAT)).toBe(true);
    expect(FILTERED_FORMAT).not.toBe(FILTER_FORMAT);
    expect(FILTERED_FORMAT.startsWith("@")).toBe(true);
  });
});

describe("which value field a column belongs to", () => {
  it("is what the helper reported, even after NON EMPTY dropped a column", () => {
    // (2025, Rides) (2025, Rides %) (2026, Rides %) (Total, Rides) (Total, Rides %)
    const reported = [0, 1, 1, 0, 1];
    expect([0, 1, 2, 3, 4].map((c) => valueFieldIndex(reported, c, 2))).toEqual([0, 1, 1, 0, 1]);
  });

  it("is unknown for the default measure, out of range, or without the list", () => {
    expect(valueFieldIndex([-1], 0, 0)).toBeNull();
    expect(valueFieldIndex([0, 1], 2, 2)).toBeNull();
    expect(valueFieldIndex([0, 1], -1, 2)).toBeNull();
    expect(valueFieldIndex([3], 0, 2)).toBeNull();
    expect(valueFieldIndex(undefined, 0, 2)).toBeNull();
  });
});

describe("a value field's own number format", () => {
  // sampleGrid: 6 value columns = 3 column tuples x (Rides, Fare).
  it("replaces the cube's format in every value cell of that field", () => {
    const block = buildBlock(sampleGrid(), {
      filters: 2,
      rows: 1,
      columns: ["City"],
      valueFormats: ["0.0", null],
    });
    for (let r = block.shape.dataTop; r < block.rows; r++) {
      expect(block.numberFormats[r].slice(1)).toEqual(["0.0", "#,##0.00", "0.0", "#,##0.00", "0.0", "#,##0.00"]);
    }
  });

  it("applies to the second field too, and to empty cells", () => {
    const block = buildBlock(sampleGrid(), {
      filters: 2,
      rows: 1,
      columns: ["City"],
      valueFormats: [null, "$#,##0.00"],
    });
    expect(block.numberFormats[5].slice(1)).toEqual(["#,##0", "$#,##0.00", "#,##0", "$#,##0.00", "#,##0", "$#,##0.00"]);
    expect(block.values[5][4]).toBe(""); // the null value keeps the field's format
  });

  it("leaves header cells, text values, and grids it cannot map alone", () => {
    const grid = sampleGrid();
    grid.values[0][0] = "n/a";
    const block = buildBlock(grid, { filters: 2, rows: 1, columns: ["City"], valueFormats: ["0.0", "0.0"] });
    expect(block.numberFormats[5][1]).toBe(TEXT_FORMAT);
    expect(block.numberFormats[3].every((f) => f === TEXT_FORMAT || f === FILTER_FORMAT)).toBe(true);
    // Columns the helper could not tie to a value field (the cube's default measure).
    const grid2 = sampleGrid();
    grid2.colValueIndex = grid2.colValueIndex.map(() => -1);
    const unmapped = buildBlock(grid2, { filters: 2, rows: 1, columns: ["City"], valueFormats: ["0", "0"] });
    expect(unmapped.numberFormats[5][1]).toBe("#,##0");
  });
});

describe("the members behind a value cell, for Show Details", () => {
  const year = ref(CAL, "2025", 1, { hasChildren: true });
  const q1 = ref(CAL, "Q1", 2);
  const headers = {
    rowHeaders: [
      [member(tehran), member(year)],
      [blank(tehran), member(q1, 1)],
      [label("Tehran Total", "subtotal"), label("", "subtotal")],
      [label("Grand Total", "grandTotal"), label("", "grandTotal")],
    ],
    colHeaderRows: [
      [member(shiraz), blank(shiraz), label("Grand Total", "grandTotal")],
      [label("Rides", "measure"), label("Fare", "measure"), label("Rides", "measure")],
    ],
  };

  it("takes every row member of the line and every column member above", () => {
    expect(detailsMembers(headers, 0, 0)).toEqual([tehran.uniqueName, year.uniqueName, shiraz.uniqueName]);
  });

  it("counts a blanked repeat as its member", () => {
    expect(detailsMembers(headers, 1, 1)).toEqual([tehran.uniqueName, q1.uniqueName, shiraz.uniqueName]);
  });

  it("skips totals and measure captions", () => {
    expect(detailsMembers(headers, 2, 0)).toEqual([shiraz.uniqueName]); // subtotal line: no member cells
    expect(detailsMembers(headers, 0, 2)).toEqual([tehran.uniqueName, year.uniqueName]); // Grand Total column
    expect(detailsMembers(headers, 3, 2)).toEqual([]); // grand total of everything
  });

  it("does not repeat a member and copes with positions past the end", () => {
    const twice = { rowHeaders: [[member(tehran), blank(tehran)]], colHeaderRows: [] };
    expect(detailsMembers(twice, 0, 0)).toEqual([tehran.uniqueName]);
    expect(detailsMembers(headers, 9, 9)).toEqual([]);
  });
});

describe("the cells of a details sheet", () => {
  it("is a text header row and the rows, text as text and numbers as General", () => {
    const out = detailsCells(["Month", "City", "Rides"], [["Jan 2025", "Tehran", 12], ["01", null, 3.5]]);
    expect(out.cols).toBe(3);
    expect(out.values).toEqual([
      ["Month", "City", "Rides"],
      ["Jan 2025", "Tehran", 12],
      ["01", "", 3.5],
    ]);
    expect(out.numberFormats).toEqual([
      [TEXT_FORMAT, TEXT_FORMAT, TEXT_FORMAT],
      [TEXT_FORMAT, TEXT_FORMAT, GENERAL_FORMAT],
      [TEXT_FORMAT, GENERAL_FORMAT, GENERAL_FORMAT],
    ]);
  });

  it("pads short rows and cuts long ones to the header", () => {
    const out = detailsCells(["A", "B"], [[1], [1, 2, 3]]);
    expect(out.values.slice(1)).toEqual([
      [1, ""],
      [1, 2],
    ]);
  });

  it("still has a header row with no rows, and one column with no columns", () => {
    expect(detailsCells(["A"], []).values).toEqual([["A"]]);
    expect(detailsCells([], []).cols).toBe(1);
  });
});

describe("Repeat All Item Labels", () => {
  // Outer field City, inner Calendar: the city is sent once, then as blank repeats.
  function nested(): Grid {
    const grid = plainGrid(2, 1, 4);
    grid.rowFieldCaptions = ["City", "Calendar"];
    const y2026 = ref(CAL, "2026", 1);
    grid.rowHeaders = [
      [member(tehran), member(y2025)],
      [blank(tehran), member(y2026)],
      [member(shiraz), member(y2025)],
      [blank(shiraz), member(y2026)],
    ];
    return grid;
  }

  it("leaves repeats empty by default", () => {
    const block = buildBlock(nested(), { filters: 0, rows: 2, columns: [] });
    expect(block.values.slice(1).map((line) => line[0])).toEqual(["Tehran", "", "Shiraz", ""]);
  });

  it("writes each repeat's caption when on, as text, and keeps the member behind it", () => {
    const block = buildBlock(nested(), { filters: 0, rows: 2, columns: [], repeatLabels: true });
    expect(block.values.slice(1).map((line) => line[0])).toEqual(["Tehran", "Tehran", "Shiraz", "Shiraz"]);
    expect(block.numberFormats[2][0]).toBe(TEXT_FORMAT);
    expect(memberAt(block, 2, 0)).toEqual({ member: tehran, axis: "rows" });
  });

  it("does not invent labels for totals", () => {
    const grid = nested();
    grid.rowHeaders.push([label("Grand Total", "grandTotal"), label("", "grandTotal")]);
    grid.values.push([1]);
    grid.cellFormats.push([0]);
    grid.rowKinds.push("grandTotal");
    const block = buildBlock(grid, { filters: 0, rows: 2, columns: [], repeatLabels: true });
    expect(block.values[5].slice(0, 2)).toEqual(["Grand Total", ""]);
  });
});

describe("For empty cells show", () => {
  it("writes the text, as text, where a value is missing, and nowhere else", () => {
    const block = buildBlock(sampleGrid(), { filters: 2, rows: 1, columns: ["City"], emptyCellText: "n/a" });
    expect(block.values[5][4]).toBe("n/a");
    expect(block.numberFormats[5][4]).toBe(TEXT_FORMAT);
    expect(block.values[5][3]).toBe(2);
    expect(block.values[2].every((v) => v === "")).toBe(true); // the blank row is not a value cell
  });

  it("is guarded like any text when it starts with =", () => {
    const block = buildBlock(sampleGrid(), { filters: 2, rows: 1, columns: ["City"], emptyCellText: "-" });
    expect(excelValues(block, false)[5][4]).toBe("'-");
  });

  it("leaves missing values empty by default", () => {
    const block = buildBlock(sampleGrid(), { filters: 2, rows: 1, columns: ["City"], emptyCellText: "" });
    expect(block.values[5][4]).toBe("");
  });
});

describe("Expand Entire Field", () => {
  const q2 = ref(CAL, "Q2", 2, { hasChildren: true });
  const opened = ref(CAL, "2025", 1, { hasChildren: true, expanded: true });
  const closed = ref(CAL, "2026", 1, { hasChildren: true });

  it("lists the field's shown members that have children and are not open yet", () => {
    const headers = {
      rowHeaders: [[member(opened)], [member(q1, 1)], [member(q2, 1)], [member(closed)], [blank(closed)]],
      colHeaderRows: [[member(tehran)]],
    };
    // q1 (from the fixtures above) is already expanded.
    expect(expandableMembers(headers, CAL)).toEqual({ members: [q2.uniqueName, closed.uniqueName], anyChildren: true });
  });

  it("finds members on the columns too", () => {
    const headers = { rowHeaders: [], colHeaderRows: [[member(closed), blank(closed)]] };
    expect(expandableMembers(headers, CAL)).toEqual({ members: [closed.uniqueName], anyChildren: true });
  });

  it("says when no item has children, and when all are already open", () => {
    expect(expandableMembers({ rowHeaders: [[member(tehran)], [member(shiraz)]], colHeaderRows: [] }, CITY)).toEqual({
      members: [],
      anyChildren: false,
    });
    expect(expandableMembers({ rowHeaders: [[member(opened)]], colHeaderRows: [] }, CAL)).toEqual({
      members: [],
      anyChildren: true,
    });
  });
});
