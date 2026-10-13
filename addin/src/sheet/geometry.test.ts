import { describe, expect, it } from "vitest";
import {
  MAX_COLS,
  MAX_ROWS,
  batchAddresses,
  cellA1,
  chunkRect,
  columnLetters,
  containsCell,
  fitsSheet,
  intersect,
  rectA1,
  runs,
  subtract,
  type Rect,
} from "./geometry";

const area = (rects: Rect[]) => rects.reduce((sum, r) => sum + r.rows * r.cols, 0);

describe("A1 addresses", () => {
  it("converts column indexes to letters", () => {
    expect(columnLetters(0)).toBe("A");
    expect(columnLetters(25)).toBe("Z");
    expect(columnLetters(26)).toBe("AA");
    expect(columnLetters(27)).toBe("AB");
    expect(columnLetters(701)).toBe("ZZ");
    expect(columnLetters(702)).toBe("AAA");
    expect(columnLetters(MAX_COLS - 1)).toBe("XFD");
  });

  it("builds cell and range addresses", () => {
    expect(cellA1(0, 0)).toBe("A1");
    expect(cellA1(9, 2)).toBe("C10");
    expect(rectA1({ row: 1, col: 1, rows: 1, cols: 1 })).toBe("B2");
    expect(rectA1({ row: 1, col: 1, rows: 3, cols: 2 })).toBe("B2:C4");
    expect(rectA1({ row: MAX_ROWS - 1, col: MAX_COLS - 1, rows: 1, cols: 1 })).toBe("XFD1048576");
  });
});

describe("rectangles", () => {
  const a: Rect = { row: 2, col: 2, rows: 4, cols: 4 }; // rows 2-5, cols 2-5

  it("tests containment at the edges", () => {
    expect(containsCell(a, 2, 2)).toBe(true);
    expect(containsCell(a, 5, 5)).toBe(true);
    expect(containsCell(a, 6, 5)).toBe(false);
    expect(containsCell(a, 1, 2)).toBe(false);
  });

  it("intersects", () => {
    expect(intersect(a, { row: 4, col: 4, rows: 10, cols: 10 })).toEqual({
      row: 4,
      col: 4,
      rows: 2,
      cols: 2,
    });
    expect(intersect(a, { row: 6, col: 2, rows: 1, cols: 1 })).toBeNull();
    expect(intersect(a, { row: 0, col: 0, rows: 2, cols: 9 })).toBeNull();
  });

  it("subtracts nothing when there is no overlap or no previous extent", () => {
    expect(subtract(a, null)).toEqual([a]);
    expect(subtract(a, { row: 20, col: 20, rows: 2, cols: 2 })).toEqual([a]);
  });

  it("subtracts everything when fully covered", () => {
    expect(subtract(a, { row: 0, col: 0, rows: 10, cols: 10 })).toEqual([]);
    expect(subtract(a, a)).toEqual([]);
  });

  it("returns the grown part when a block gets taller and wider", () => {
    const before: Rect = { row: 2, col: 2, rows: 2, cols: 2 };
    const parts = subtract(a, before);
    expect(area(parts)).toBe(16 - 4);
    for (const part of parts) expect(intersect(part, before)).toBeNull();
    expect(parts).toEqual([
      { row: 4, col: 2, rows: 2, cols: 4 },
      { row: 2, col: 4, rows: 2, cols: 2 },
    ]);
  });

  it("returns four pieces around a hole in the middle, without overlap", () => {
    const hole: Rect = { row: 3, col: 3, rows: 2, cols: 2 };
    const parts = subtract(a, hole);
    expect(parts).toHaveLength(4);
    expect(area(parts)).toBe(16 - 4);
    for (let i = 0; i < parts.length; i++) {
      expect(intersect(parts[i], hole)).toBeNull();
      for (let j = i + 1; j < parts.length; j++) expect(intersect(parts[i], parts[j])).toBeNull();
    }
  });

  it("knows the sheet limits", () => {
    expect(fitsSheet({ row: 0, col: 0, rows: MAX_ROWS, cols: MAX_COLS })).toBe(true);
    expect(fitsSheet({ row: 1, col: 0, rows: MAX_ROWS, cols: 1 })).toBe(false);
    expect(fitsSheet({ row: 0, col: MAX_COLS - 1, rows: 1, cols: 2 })).toBe(false);
  });
});

describe("chunkRect", () => {
  it("keeps the rectangle's own origin", () => {
    expect(chunkRect({ row: 5, col: 3, rows: 5, cols: 4 }, 8)).toEqual([
      { row: 5, col: 3, rows: 2, cols: 4 },
      { row: 7, col: 3, rows: 2, cols: 4 },
      { row: 9, col: 3, rows: 1, cols: 4 },
    ]);
  });
});

describe("batchAddresses", () => {
  it("joins up to the length limit", () => {
    expect(batchAddresses(["A1", "B2", "C3", "D4"], 5)).toEqual(["A1,B2", "C3,D4"]);
    expect(batchAddresses(["A1", "B2", "C3"], 100)).toEqual(["A1,B2,C3"]);
    expect(batchAddresses([], 10)).toEqual([]);
  });

  it("gives an over-long address a string of its own", () => {
    expect(batchAddresses(["A1", "AAA1000:ZZZ2000", "B2"], 6)).toEqual([
      "A1",
      "AAA1000:ZZZ2000",
      "B2",
    ]);
  });
});

describe("runs", () => {
  it("emits maximal runs of equal non-null keys", () => {
    const keys = [null, 1, 1, 2, null, 2, 2, 2];
    const out: [number, number, number][] = [];
    runs(
      keys.length,
      (i) => keys[i],
      (key, start, length) => out.push([key, start, length]),
    );
    expect(out).toEqual([
      [1, 1, 2],
      [2, 3, 1],
      [2, 5, 3],
    ]);
  });

  it("emits nothing for an empty list", () => {
    const out: number[] = [];
    runs(
      0,
      () => 1,
      (key) => out.push(key),
    );
    expect(out).toEqual([]);
  });
});
