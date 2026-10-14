// The words and the checks behind the field menu (sort, label filters, value filters)
// and the value field menu (Show Values As, Number Format). Pure. The wording follows
// Excel's own PivotTable menus on Windows, which is what the user knows.

import type {
  FieldSort,
  LabelFilter,
  LabelOperator,
  ShowValuesAs,
  ValueFilter,
  ValueOperator,
} from "../types";

export interface OperatorChoice<Op extends string> {
  operator: Op;
  /** Menu wording, as in Excel. */
  label: string;
  /** Wording inside the form: "Show items for which the label <phrase> …". */
  phrase: string;
  /** Takes an upper bound as well. */
  two: boolean;
}

export const LABEL_OPERATORS: OperatorChoice<LabelOperator>[] = [
  { operator: "equals", label: "Equals…", phrase: "equals", two: false },
  { operator: "notEquals", label: "Does Not Equal…", phrase: "does not equal", two: false },
  { operator: "beginsWith", label: "Begins With…", phrase: "begins with", two: false },
  { operator: "notBeginsWith", label: "Does Not Begin With…", phrase: "does not begin with", two: false },
  { operator: "endsWith", label: "Ends With…", phrase: "ends with", two: false },
  { operator: "notEndsWith", label: "Does Not End With…", phrase: "does not end with", two: false },
  { operator: "contains", label: "Contains…", phrase: "contains", two: false },
  { operator: "notContains", label: "Does Not Contain…", phrase: "does not contain", two: false },
  { operator: "greaterThan", label: "Greater Than…", phrase: "is greater than", two: false },
  { operator: "greaterOrEqual", label: "Greater Than Or Equal To…", phrase: "is greater than or equal to", two: false },
  { operator: "lessThan", label: "Less Than…", phrase: "is less than", two: false },
  { operator: "lessOrEqual", label: "Less Than Or Equal To…", phrase: "is less than or equal to", two: false },
  { operator: "between", label: "Between…", phrase: "is between", two: true },
  { operator: "notBetween", label: "Not Between…", phrase: "is not between", two: true },
];

/** The comparisons; "Top 10…" is offered separately, as in Excel. */
export const VALUE_OPERATORS: OperatorChoice<ValueOperator>[] = [
  { operator: "equals", label: "Equals…", phrase: "equals", two: false },
  { operator: "notEquals", label: "Does Not Equal…", phrase: "does not equal", two: false },
  { operator: "greaterThan", label: "Greater Than…", phrase: "is greater than", two: false },
  { operator: "greaterOrEqual", label: "Greater Than Or Equal To…", phrase: "is greater than or equal to", two: false },
  { operator: "lessThan", label: "Less Than…", phrase: "is less than", two: false },
  { operator: "lessOrEqual", label: "Less Than Or Equal To…", phrase: "is less than or equal to", two: false },
  { operator: "between", label: "Between…", phrase: "is between", two: true },
  { operator: "notBetween", label: "Not Between…", phrase: "is not between", two: true },
];

export const TOP_LABEL = "Top 10…";

/** Longest label filter text the helper accepts. */
export const MAX_LABEL_TEXT = 200;
/** Longest custom number format the pane accepts. */
export const MAX_NUMBER_FORMAT = 100;

export type Checked<T> = { ok: true; value: T } | { ok: false; message: string };

function fail<T>(message: string): Checked<T> {
  return { ok: false, message };
}

function isTwo(operator: string): boolean {
  return operator === "between" || operator === "notBetween";
}

/** The label filter the form describes, or what is wrong with it. */
export function checkLabelFilter(input: {
  operator: LabelOperator;
  value: string;
  value2: string;
}): Checked<LabelFilter> {
  const value = input.value.trim();
  if (value === "") return fail("Type the text to compare with.");
  if (value.length > MAX_LABEL_TEXT) return fail(`Use at most ${MAX_LABEL_TEXT} characters.`);
  if (!isTwo(input.operator)) return { ok: true, value: { operator: input.operator, value } };
  const value2 = input.value2.trim();
  if (value2 === "") return fail("Type both ends of the range.");
  if (value2.length > MAX_LABEL_TEXT) return fail(`Use at most ${MAX_LABEL_TEXT} characters.`);
  // Labels compare without regard to case, so the range is checked the same way.
  if (value.toLowerCase() > value2.toLowerCase()) {
    return fail("The first text must come before the second, alphabetically.");
  }
  return { ok: true, value: { operator: input.operator, value, value2 } };
}

/**
 * A number as typed: "1500", "1,500", "-2.5", "12.5%"-less. Null when it is not one.
 * Thousands commas are accepted only where they would be in "1,234,567".
 */
export function parseNumber(text: string): number | null {
  const plain = text.trim();
  if (plain === "") return null;
  if (!/^[-+]?(\d+|\d{1,3}(,\d{3})+)(\.\d+)?$|^[-+]?\.\d+$/.test(plain)) return null;
  const n = Number(plain.replace(/,/g, ""));
  return isFinite(n) ? n : null;
}

/** The comparison value filter the form describes, or what is wrong with it. */
export function checkValueFilter(input: {
  operator: ValueOperator;
  measure: string;
  value: string;
  value2: string;
}): Checked<ValueFilter> {
  if (input.measure === "") return fail("Choose a value to compare.");
  const value = parseNumber(input.value);
  if (value === null) return fail("Type a number.");
  if (!isTwo(input.operator)) {
    return { ok: true, value: { operator: input.operator, measure: input.measure, value } };
  }
  const value2 = parseNumber(input.value2);
  if (value2 === null) return fail("Type a number for both ends of the range.");
  if (value > value2) return fail("The first number must not be larger than the second.");
  return { ok: true, value: { operator: input.operator, measure: input.measure, value, value2 } };
}

/** The Top 10 filter the form describes, or what is wrong with it. */
export function checkTopFilter(input: {
  direction: "top" | "bottom";
  count: string;
  topBy: "items" | "percent" | "sum";
  measure: string;
}): Checked<ValueFilter> {
  if (input.measure === "") return fail("Choose a value to rank by.");
  const count = parseNumber(input.count);
  if (count === null || count <= 0) return fail("Type a number larger than 0.");
  if (input.topBy === "items" && Math.floor(count) !== count) return fail("Type a whole number of items.");
  if (input.topBy === "percent" && count > 100) return fail("A percent can be at most 100.");
  return {
    ok: true,
    value: { operator: input.direction, measure: input.measure, value: count, topBy: input.topBy },
  };
}

/** A custom number format, or what is wrong with it. */
export function checkNumberFormat(text: string): Checked<string> {
  const format = text.trim();
  if (format === "") return fail("Type a number format, such as #,##0.00.");
  if (format.length > MAX_NUMBER_FORMAT) return fail(`Use at most ${MAX_NUMBER_FORMAT} characters.`);
  return { ok: true, value: format };
}

// ---------------------------------------------------------------------------
// Describing what is set, for the panel and the chips.
// ---------------------------------------------------------------------------

function quote(text: string): string {
  return "“" + text + "”";
}

export function describeSort(sort: FieldSort | null | undefined, measureCaption: (name: string) => string): string {
  if (!sort) return "Cube order";
  if (sort.by === "caption") return sort.order === "asc" ? "A to Z" : "Z to A";
  const by = measureCaption(sort.measure ?? "");
  return `${by}, ${sort.order === "asc" ? "smallest to largest" : "largest to smallest"}`;
}

export function describeLabelFilter(filter: LabelFilter): string {
  const choice = LABEL_OPERATORS.find((c) => c.operator === filter.operator);
  const phrase = choice ? choice.phrase : filter.operator;
  return choice && choice.two
    ? `Label ${phrase} ${quote(filter.value)} and ${quote(filter.value2 ?? "")}`
    : `Label ${phrase} ${quote(filter.value)}`;
}

function formatNumber(n: number): string {
  return n.toLocaleString("en-US", { maximumFractionDigits: 6 });
}

export function describeValueFilter(filter: ValueFilter, measureCaption: (name: string) => string): string {
  const by = measureCaption(filter.measure);
  if (filter.operator === "top" || filter.operator === "bottom") {
    const unit = filter.topBy === "percent" ? "percent" : filter.topBy === "sum" ? "sum" : "items";
    return `${filter.operator === "top" ? "Top" : "Bottom"} ${formatNumber(filter.value)} ${unit} by ${by}`;
  }
  const choice = VALUE_OPERATORS.find((c) => c.operator === filter.operator);
  const phrase = choice ? choice.phrase : filter.operator;
  return choice && choice.two
    ? `${by} ${phrase} ${formatNumber(filter.value)} and ${formatNumber(filter.value2 ?? 0)}`
    : `${by} ${phrase} ${formatNumber(filter.value)}`;
}

// ---------------------------------------------------------------------------
// The value field menu.
// ---------------------------------------------------------------------------

export const SHOW_AS_CHOICES: { showAs: ShowValuesAs; label: string }[] = [
  { showAs: "normal", label: "No Calculation" },
  { showAs: "percentOfGrandTotal", label: "% of Grand Total" },
  { showAs: "percentOfColumnTotal", label: "% of Column Total" },
  { showAs: "percentOfRowTotal", label: "% of Row Total" },
];

/** The number formats offered, in order. `format: null` is the cube's own. */
export const NUMBER_FORMAT_CHOICES: { format: string | null; label: string }[] = [
  { format: null, label: "Cube default" },
  { format: "General", label: "General" },
  { format: "#,##0", label: "Number  1,235" },
  { format: "#,##0.00", label: "Number  1,234.56" },
  { format: "0.00%", label: "Percent  12.35%" },
  { format: "$#,##0.00", label: "Currency  $1,234.56" },
];

/** True when `format` is one of the listed ones (otherwise it is a custom one). */
export function isListedFormat(format: string | null | undefined): boolean {
  const value = format ?? null;
  return NUMBER_FORMAT_CHOICES.some((choice) => choice.format === value);
}
