import { describe, expect, it } from "vitest";
import {
  RELATIVE_ITEMS,
  checkEmptyCellText,
  checkShowAs,
  showAsChoice,
  LABEL_OPERATORS,
  NUMBER_FORMAT_CHOICES,
  SHOW_AS_CHOICES,
  VALUE_OPERATORS,
  checkLabelFilter,
  checkNumberFormat,
  checkTopFilter,
  checkValueFilter,
  describeLabelFilter,
  describeSort,
  describeValueFilter,
  isListedFormat,
  parseNumber,
} from "./fieldMenu";

const RIDES = "[Measures].[Rides]";
const caption = (name: string) => (name === RIDES ? "Rides" : name);

describe("the wording", () => {
  it("offers Excel's 14 label filters, in Excel's order and words", () => {
    expect(LABEL_OPERATORS.map((c) => c.label)).toEqual([
      "Equals…",
      "Does Not Equal…",
      "Begins With…",
      "Does Not Begin With…",
      "Ends With…",
      "Does Not End With…",
      "Contains…",
      "Does Not Contain…",
      "Greater Than…",
      "Greater Than Or Equal To…",
      "Less Than…",
      "Less Than Or Equal To…",
      "Between…",
      "Not Between…",
    ]);
    expect(LABEL_OPERATORS.filter((c) => c.two).map((c) => c.operator)).toEqual(["between", "notBetween"]);
  });

  it("offers the value comparisons, with Top 10 kept apart", () => {
    expect(VALUE_OPERATORS.map((c) => c.operator)).toEqual([
      "equals",
      "notEquals",
      "greaterThan",
      "greaterOrEqual",
      "lessThan",
      "lessOrEqual",
      "between",
      "notBetween",
    ]);
  });

  it("lists the Show Values As choices and number formats asked for", () => {
    expect(SHOW_AS_CHOICES.map((c) => c.label)).toEqual([
      "No Calculation",
      "% of Grand Total",
      "% of Column Total",
      "% of Row Total",
      "% Of…",
      "% of Parent Row Total",
      "% of Parent Column Total",
      "% of Parent Total…",
      "Difference From…",
      "% Difference From…",
      "Running Total In…",
      "% Running Total In…",
      "Rank Smallest to Largest…",
      "Rank Largest to Smallest…",
      "Index",
    ]);
    // An item asks for something exactly when its wording ends with "…".
    for (const choice of SHOW_AS_CHOICES) {
      expect(choice.label.endsWith("…")).toBe(choice.needs !== "nothing");
    }
    expect(SHOW_AS_CHOICES.filter((c) => c.needs === "fieldAndItem").map((c) => c.showAs)).toEqual([
      "percentOf",
      "differenceFrom",
      "percentDifferenceFrom",
    ]);
    expect(NUMBER_FORMAT_CHOICES.map((c) => c.format)).toEqual([
      null,
      "General",
      "#,##0",
      "#,##0.00",
      "0.00%",
      "$#,##0.00",
    ]);
    expect(isListedFormat(null)).toBe(true);
    expect(isListedFormat(undefined)).toBe(true);
    expect(isListedFormat("#,##0")).toBe(true);
    expect(isListedFormat("0.0")).toBe(false);
  });
});

describe("label filter form", () => {
  it("accepts one text, trimmed", () => {
    expect(checkLabelFilter({ operator: "beginsWith", value: "  Te ", value2: "" })).toEqual({
      ok: true,
      value: { operator: "beginsWith", value: "Te" },
    });
  });

  it("refuses empty or over-long text", () => {
    expect(checkLabelFilter({ operator: "equals", value: "  ", value2: "" }).ok).toBe(false);
    expect(checkLabelFilter({ operator: "equals", value: "x".repeat(201), value2: "" }).ok).toBe(false);
    expect(checkLabelFilter({ operator: "equals", value: "x".repeat(200), value2: "" }).ok).toBe(true);
  });

  it("needs both bounds for between, in order, without regard to case", () => {
    expect(checkLabelFilter({ operator: "between", value: "A", value2: "" }).ok).toBe(false);
    expect(checkLabelFilter({ operator: "between", value: "M", value2: "c" }).ok).toBe(false);
    expect(checkLabelFilter({ operator: "notBetween", value: "a", value2: "M" })).toEqual({
      ok: true,
      value: { operator: "notBetween", value: "a", value2: "M" },
    });
    expect(checkLabelFilter({ operator: "between", value: "b", value2: "B" }).ok).toBe(true);
  });

  it("ignores a second text it does not need", () => {
    expect(checkLabelFilter({ operator: "contains", value: "an", value2: "zzz" })).toEqual({
      ok: true,
      value: { operator: "contains", value: "an" },
    });
  });
});

describe("numbers as typed", () => {
  it("reads plain, signed, decimal and comma-grouped numbers", () => {
    expect(parseNumber("1500")).toBe(1500);
    expect(parseNumber("1,500")).toBe(1500);
    expect(parseNumber("1,234,567.25")).toBe(1234567.25);
    expect(parseNumber("-2.5")).toBe(-2.5);
    expect(parseNumber(" +3 ")).toBe(3);
    expect(parseNumber(".5")).toBe(0.5);
  });

  it("refuses what is not a number", () => {
    for (const text of ["", "abc", "1,5", "12,34,567", "1.2.3", "10%", "1e5", "--1"]) {
      expect(parseNumber(text)).toBeNull();
    }
  });
});

describe("value filter form", () => {
  it("accepts a comparison", () => {
    expect(checkValueFilter({ operator: "greaterThan", measure: RIDES, value: "1,000", value2: "" })).toEqual({
      ok: true,
      value: { operator: "greaterThan", measure: RIDES, value: 1000 },
    });
  });

  it("needs a measure and a number", () => {
    expect(checkValueFilter({ operator: "equals", measure: "", value: "1", value2: "" }).ok).toBe(false);
    expect(checkValueFilter({ operator: "equals", measure: RIDES, value: "x", value2: "" }).ok).toBe(false);
  });

  it("needs both bounds for between, the first not above the second", () => {
    expect(checkValueFilter({ operator: "between", measure: RIDES, value: "1", value2: "" }).ok).toBe(false);
    expect(checkValueFilter({ operator: "between", measure: RIDES, value: "10", value2: "5" }).ok).toBe(false);
    expect(checkValueFilter({ operator: "notBetween", measure: RIDES, value: "5", value2: "5" })).toEqual({
      ok: true,
      value: { operator: "notBetween", measure: RIDES, value: 5, value2: 5 },
    });
  });
});

describe("Top 10 form", () => {
  it("defaults to what Excel offers and accepts it", () => {
    expect(checkTopFilter({ direction: "top", count: "10", topBy: "items", measure: RIDES })).toEqual({
      ok: true,
      value: { operator: "top", measure: RIDES, value: 10, topBy: "items" },
    });
  });

  it("needs a positive number", () => {
    for (const count of ["0", "-3", "", "x"]) {
      expect(checkTopFilter({ direction: "top", count, topBy: "items", measure: RIDES }).ok).toBe(false);
    }
  });

  it("needs whole items and a percent of at most 100", () => {
    expect(checkTopFilter({ direction: "bottom", count: "2.5", topBy: "items", measure: RIDES }).ok).toBe(false);
    expect(checkTopFilter({ direction: "top", count: "100", topBy: "percent", measure: RIDES }).ok).toBe(true);
    expect(checkTopFilter({ direction: "top", count: "100.5", topBy: "percent", measure: RIDES }).ok).toBe(false);
    expect(checkTopFilter({ direction: "top", count: "12.5", topBy: "percent", measure: RIDES }).ok).toBe(true);
    expect(checkTopFilter({ direction: "top", count: "2,500,000", topBy: "sum", measure: RIDES })).toMatchObject({
      ok: true,
      value: { value: 2500000, topBy: "sum" },
    });
  });

  it("needs a measure", () => {
    expect(checkTopFilter({ direction: "top", count: "10", topBy: "items", measure: "" }).ok).toBe(false);
  });
});

describe("custom number format", () => {
  it("accepts a format, trimmed", () => {
    expect(checkNumberFormat(" 0.0% ")).toEqual({ ok: true, value: "0.0%" });
  });

  it("refuses an empty one or one over 100 characters", () => {
    expect(checkNumberFormat("   ").ok).toBe(false);
    expect(checkNumberFormat("0".repeat(101)).ok).toBe(false);
    expect(checkNumberFormat("0".repeat(100)).ok).toBe(true);
  });
});

describe("describing what is set", () => {
  it("describes sorts", () => {
    expect(describeSort(null, caption)).toBe("Cube order");
    expect(describeSort({ by: "caption", order: "asc" }, caption)).toBe("A to Z");
    expect(describeSort({ by: "caption", order: "desc" }, caption)).toBe("Z to A");
    expect(describeSort({ by: "measure", measure: RIDES, order: "desc" }, caption)).toBe(
      "Rides, largest to smallest",
    );
  });

  it("describes label and value filters", () => {
    expect(describeLabelFilter({ operator: "beginsWith", value: "Te" })).toBe("Label begins with “Te”");
    expect(describeLabelFilter({ operator: "between", value: "a", value2: "m" })).toBe(
      "Label is between “a” and “m”",
    );
    expect(describeValueFilter({ operator: "top", measure: RIDES, value: 3, topBy: "items" }, caption)).toBe(
      "Top 3 items by Rides",
    );
    expect(describeValueFilter({ operator: "bottom", measure: RIDES, value: 20, topBy: "percent" }, caption)).toBe(
      "Bottom 20 percent by Rides",
    );
    expect(describeValueFilter({ operator: "greaterThan", measure: RIDES, value: 1500 }, caption)).toBe(
      "Rides is greater than 1,500",
    );
    expect(describeValueFilter({ operator: "between", measure: RIDES, value: 1, value2: 2 }, caption)).toBe(
      "Rides is between 1 and 2",
    );
  });
});

describe("Show Values As form", () => {
  const CITY = "[Geography].[City]";
  const YEAR = "[Date].[Year]";
  const axis = [CITY, YEAR];

  it("needs nothing for the kinds without …, and drops any base it is given", () => {
    for (const showAs of ["normal", "percentOfGrandTotal", "percentOfParentRowTotal", "index"] as const) {
      expect(checkShowAs({ showAs, baseField: CITY, baseItem: "previous", axisFields: axis })).toEqual({
        ok: true,
        value: { showAs, baseField: null, baseItem: null },
      });
    }
    expect(checkShowAs({ showAs: "percentOfColumnTotal", baseField: "", baseItem: "", axisFields: [] }).ok).toBe(true);
  });

  it("requires a base field in Rows or Columns for the field kinds", () => {
    for (const showAs of ["percentOfParentTotal", "runningTotalIn", "percentRunningTotalIn", "rankAscending", "rankDescending"] as const) {
      expect(checkShowAs({ showAs, baseField: "", baseItem: "", axisFields: axis })).toEqual({
        ok: false,
        message: "Choose a base field.",
      });
      expect(checkShowAs({ showAs, baseField: "[Service].[Type]", baseItem: "", axisFields: axis }).ok).toBe(false);
      expect(checkShowAs({ showAs, baseField: YEAR, baseItem: "next", axisFields: axis })).toEqual({
        ok: true,
        value: { showAs, baseField: YEAR, baseItem: null },
      });
    }
  });

  it("requires a base item too for % Of, Difference From and % Difference From", () => {
    for (const showAs of ["percentOf", "differenceFrom", "percentDifferenceFrom"] as const) {
      expect(checkShowAs({ showAs, baseField: CITY, baseItem: "", axisFields: axis })).toEqual({
        ok: false,
        message: "Choose a base item.",
      });
      expect(checkShowAs({ showAs, baseField: "", baseItem: "previous", axisFields: axis }).ok).toBe(false);
      expect(checkShowAs({ showAs, baseField: CITY, baseItem: "[Geography].[City].&[1]", axisFields: axis })).toEqual({
        ok: true,
        value: { showAs, baseField: CITY, baseItem: "[Geography].[City].&[1]" },
      });
    }
  });

  it("offers (previous) and (next) before the field's items", () => {
    expect(RELATIVE_ITEMS).toEqual([
      { item: "previous", label: "(previous)" },
      { item: "next", label: "(next)" },
    ]);
  });

  it("finds a choice by kind, and No Calculation for none", () => {
    expect(showAsChoice("runningTotalIn").label).toBe("Running Total In…");
    expect(showAsChoice(undefined).label).toBe("No Calculation");
  });
});

describe("For empty cells show", () => {
  it("accepts up to 20 characters, and nothing", () => {
    expect(checkEmptyCellText("")).toEqual({ ok: true, value: "" });
    expect(checkEmptyCellText("-")).toEqual({ ok: true, value: "-" });
    expect(checkEmptyCellText("x".repeat(20)).ok).toBe(true);
    expect(checkEmptyCellText("x".repeat(21)).ok).toBe(false);
  });
});
