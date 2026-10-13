import { describe, expect, it } from "vitest";
import type { Member } from "../types";
import {
  ROOT,
  checkState,
  describeSelection,
  initPicker,
  isAll,
  pickerResult,
  selectAll,
  selectNone,
  setChildren,
  toggle,
  toggleFlat,
  type PickerState,
} from "./picker";

const ALL = "[Date].[Calendar].[All]";
const m = (uniqueName: string, caption: string, level: number, hasChildren = false): Member => ({
  uniqueName,
  caption,
  level,
  hasChildren,
});

/**
 * The tree as the helper serves it: the top level is the years themselves (the All
 * member is never listed). 2025 has its quarters loaded; 2026 has not been opened.
 */
function tree(members: string[] = [], captions: string[] = []): PickerState {
  let state = initPicker({ members, memberCaptions: captions });
  state = setChildren(state, ROOT, [m("2025", "2025", 1, true), m("2026", "2026", 1, true)], false);
  state = setChildren(state, "2025", [m("Q1", "Q1", 2), m("Q2", "Q2", 2)], false);
  return state;
}

const sel = (state: PickerState) => Array.from(state.selected).sort();

describe("starting state", () => {
  it("has everything ticked when the field has no restriction", () => {
    const state = tree();
    expect(isAll(state)).toBe(true);
    for (const key of ["2025", "2026", "Q1", "Q2"]) expect(checkState(state, key)).toBe("on");
    expect(pickerResult(state, ALL)).toEqual({ members: [], captions: [] });
    expect(describeSelection(state, ALL)).toBe("All selected");
  });

  it("shows an existing restriction, with ancestors partly ticked", () => {
    const state = tree(["Q1"], ["Q1"]);
    expect(checkState(state, "Q1")).toBe("on");
    expect(checkState(state, "Q2")).toBe("off");
    expect(checkState(state, "2025")).toBe("partial");
    expect(checkState(state, "2026")).toBe("off");
  });

  it("keeps the caption of a selected member that is not loaded in the tree", () => {
    const state = tree(["[Date].[Calendar].[Month].&[2026]&[3]"], ["Mar 2026"]);
    expect(pickerResult(state, ALL)).toEqual({
      members: ["[Date].[Calendar].[Month].&[2026]&[3]"],
      captions: ["Mar 2026"],
    });
  });
});

describe("unticking", () => {
  it("unticking one leaf under 'all' keeps every other member", () => {
    const { state, blocked } = toggle(tree(), "Q1");
    expect(blocked).toBeUndefined();
    expect(sel(state)).toEqual(["2026", "Q2"]);
    expect(checkState(state, "Q1")).toBe("off");
    expect(checkState(state, "Q2")).toBe("on");
    expect(checkState(state, "2025")).toBe("partial");
    expect(checkState(state, "2026")).toBe("on");
    expect(isAll(state)).toBe(false);
    expect(pickerResult(state, ALL)).toEqual({ members: ["2026", "Q2"], captions: ["2026", "Q2"] });
  });

  it("unticking a year under 'all' keeps the other year", () => {
    const { state } = toggle(tree(), "2025");
    expect(sel(state)).toEqual(["2026"]);
    expect(checkState(state, "Q1")).toBe("off");
  });

  it("unticking an explicitly selected member just removes it", () => {
    const { state } = toggle(tree(["2025", "2026"]), "2026");
    expect(sel(state)).toEqual(["2025"]);
  });

  it("unticking every top-level member leaves nothing selected, which is not a valid filter", () => {
    const { state } = toggle(toggle(tree(), "2025").state, "2026");
    expect(sel(state)).toEqual([]);
    expect(pickerResult(state, ALL)).toBeNull();
    expect(describeSelection(state, ALL)).toBe("Nothing selected");
  });

  it("is blocked when the siblings are not all known", () => {
    let state = initPicker({ members: [], memberCaptions: [] });
    state = setChildren(state, ROOT, [m("a", "A", 0), m("b", "B", 0)], true);
    const result = toggle(state, "a");
    expect(result.blocked).toBe("truncated");
    expect(result.state).toBe(state);
  });

  it("is blocked when a list on the way down was cut off", () => {
    let state = tree();
    state = setChildren(state, "2025", [m("Q1", "Q1", 2), m("Q2", "Q2", 2)], true);
    expect(toggle(state, "Q1").blocked).toBe("truncated");
    // The year itself can still be unticked: its own siblings are complete.
    expect(sel(toggle(state, "2025").state)).toEqual(["2026"]);
  });
});

describe("ticking", () => {
  it("ticking the last missing child ticks the parent, all the way up to 'all'", () => {
    const start = toggle(tree(), "Q1").state; // {Q2, 2026}
    const { state } = toggle(start, "Q1");
    expect(isAll(state)).toBe(true);
    expect(sel(state)).toEqual([ROOT]);
  });

  it("ticking a parent replaces its ticked children", () => {
    const { state } = toggle(tree(["Q1"]), "2025");
    expect(sel(state)).toEqual(["2025"]);
    expect(checkState(state, "Q2")).toBe("on");
  });

  it("does not merge into a parent whose list was cut off", () => {
    let state = initPicker({ members: ["a"], memberCaptions: ["A"] });
    state = setChildren(state, ROOT, [m("a", "A", 0), m("b", "B", 0)], true);
    state = toggle(state, "b").state;
    expect(sel(state)).toEqual(["a", "b"]);
    expect(isAll(state)).toBe(false);
  });

  it("merges top-level members into 'all' for a hierarchy without an All member", () => {
    let state = initPicker({ members: ["new"], memberCaptions: ["New"] });
    state = setChildren(state, ROOT, [m("new", "New", 0), m("ret", "Returning", 0)], false);
    state = toggle(state, "ret").state;
    expect(isAll(state)).toBe(true);
    expect(pickerResult(state, null)).toEqual({ members: [], captions: [] });
  });
});

describe("select all and none", () => {
  it("select none then one item filters to that item", () => {
    let state = selectNone(tree());
    expect(checkState(state, "2025")).toBe("off");
    state = toggle(state, "2026").state;
    expect(pickerResult(state, ALL)).toEqual({ members: ["2026"], captions: ["2026"] });
    expect(describeSelection(state, ALL)).toBe("2026");
  });

  it("select all removes the restriction", () => {
    const state = selectAll(tree(["Q1"]));
    expect(pickerResult(state, ALL)).toEqual({ members: [], captions: [] });
  });

  it("treats a selection holding the All member as no restriction", () => {
    // The tree never offers it, but a workbook from an earlier version can carry it.
    const state = tree([ALL], ["All"]);
    expect(pickerResult(state, ALL)).toEqual({ members: [], captions: [] });
    expect(describeSelection(state, ALL)).toBe("All selected");
  });

  it("summarises a long selection by count", () => {
    let state = initPicker({ members: [], memberCaptions: [] });
    const items = ["a", "b", "c", "d", "e"].map((k) => m(k, k.toUpperCase(), 0));
    state = setChildren(state, ROOT, items, false);
    state = toggle(state, "a").state;
    expect(describeSelection(state, null)).toBe("4 selected");
    state = toggle(state, "b").state;
    expect(describeSelection(state, null)).toBe("C, D, E");
  });
});

describe("search results", () => {
  const tehran = m("[Geography].[City].&[1]", "Tehran", 1);
  const shiraz = m("[Geography].[City].&[3]", "Shiraz", 1);

  it("ticking a result while everything is selected means only that one", () => {
    const state = toggleFlat(tree(), tehran, ALL);
    expect(pickerResult(state, ALL)).toEqual({ members: [tehran.uniqueName], captions: ["Tehran"] });
  });

  it("does the same when the All member is what is selected", () => {
    const state = toggleFlat(tree([ALL], ["All"]), tehran, ALL);
    expect(sel(state)).toEqual([tehran.uniqueName]);
  });

  it("adds to and removes from an explicit selection", () => {
    let state = toggleFlat(tree(), tehran, ALL);
    state = toggleFlat(state, shiraz, ALL);
    expect(pickerResult(state, ALL)?.captions).toEqual(["Tehran", "Shiraz"]);
    state = toggleFlat(state, tehran, ALL);
    expect(pickerResult(state, ALL)).toEqual({ members: [shiraz.uniqueName], captions: ["Shiraz"] });
  });
});
