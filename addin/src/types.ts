// The helper API contract. The Python helper produces and consumes exactly these
// shapes (helper/src/cubebridge/pivot/model.py and pivot/grid.py). Change both sides
// together.

export type ErrorCategory =
  | "unreachable"
  | "auth"
  | "auth-mechanism"
  | "instance"
  | "too-large"
  | "server"
  | "cancelled"
  | "bad-request"
  | "not-found"
  | "forbidden"
  | "internal";

/** Body of every non-2xx response: `{ "error": ApiError }`. */
export interface ApiError {
  category: ErrorCategory;
  /** Plain-language sentence, safe to show the user as is. */
  message: string;
  /** Technical detail (server text), shown only on request. */
  detail?: string;
}

export interface Status {
  ok: true;
  version: string;
  /** Days until the localhost certificate expires; null if unknown. */
  certDaysLeft: number | null;
}

export interface Connection {
  id: string;
  name: string;
  kind: "ssas" | "demo";
  server: string;
  username: string;
}

export interface NewConnection {
  name?: string;
  /** "host", "host:port" or "host\\instance", as typed in Windows Excel. */
  server: string;
  /** "DOMAIN\\user" or "user@domain". */
  username: string;
  password: string;
}

export interface Catalog {
  name: string;
  description: string;
}

export interface Cube {
  name: string;
  caption: string;
}

export interface Measure {
  uniqueName: string;
  caption: string;
  /** Measure group; "" when the cube reports none. */
  group: string;
  /** Display folder, possibly nested with "\\"; "" for none. */
  folder: string;
}

export interface Hierarchy {
  uniqueName: string;
  caption: string;
  folder: string;
  kind: "attribute" | "user";
  /** Unique name of the All member; null when the hierarchy has none. */
  allMember: string | null;
  /** The field's top-level items are dates, so Date Filters apply to it. */
  isDate?: boolean;
}

export interface Dimension {
  uniqueName: string;
  caption: string;
  hierarchies: Hierarchy[];
}

export type KpiPart = "value" | "goal" | "status" | "trend";

/** A key performance indicator defined in the cube (Excel lists them under "KPIs"). */
export interface Kpi {
  /** KPI_NAME, as KPIValue("…") and its siblings take it. */
  name: string;
  caption: string;
  /** Display folder, possibly nested with "\\"; "" for none. */
  folder: string;
  /** The parts this KPI defines; "value" always. */
  parts: KpiPart[];
}

/** A named set defined in the cube (Excel lists them under their dimension, in "Sets"). */
export interface NamedSet {
  /** "[Set Name]", as MDX refers to it. */
  uniqueName: string;
  caption: string;
  folder: string;
  /** Dimension unique name it is listed under. */
  dimension: string;
  /**
   * The one hierarchy its items come from. Null when its items span several
   * hierarchies: such a set is listed but cannot be placed.
   */
  hierarchy: string | null;
  /** That hierarchy's All member, for totals; null when none or unknown. */
  allMember: string | null;
}

export interface Fields {
  measures: Measure[];
  dimensions: Dimension[];
  /** Absent from helpers older than KPI support; treat as empty. */
  kpis?: Kpi[];
  /** Absent from helpers older than named-set support; treat as empty. */
  sets?: NamedSet[];
}

export interface Member {
  uniqueName: string;
  caption: string;
  level: number;
  hasChildren: boolean;
}

export interface MembersRequest {
  catalog: string;
  cube: string;
  hierarchy: string;
  /** Children of this member; omitted for the top level. */
  parent?: string;
  /** Case-insensitive caption search across all levels; overrides `parent`. */
  search?: string;
  /** Default 500, maximum 5000. */
  limit?: number;
}

export interface MembersResponse {
  members: Member[];
  truncated: boolean;
}

/** A hierarchy placed in Rows, Columns or Filters. Same shape in all three. */
export interface PlacedField {
  hierarchy: string;
  caption: string;
  allMember: string | null;
  /** Drilled-down member unique names. Ignored in Filters. */
  expanded: string[];
  /** Selected member unique names; empty means no restriction. */
  members: string[];
  /** Captions for `members`, same order, for display. */
  memberCaptions: string[];
  /** Order of this field's items. Absent or null: the cube's own order. Ignored in Filters. */
  sort?: FieldSort | null;
  /** Keep only items whose caption matches. Ignored in Filters. */
  labelFilter?: LabelFilter | null;
  /** Keep only items whose value for a measure matches. Ignored in Filters. */
  valueFilter?: ValueFilter | null;
  /** Keep only items whose date matches; only for fields with `isDate`. Ignored in Filters. */
  dateFilter?: DateFilter | null;
  /**
   * Groups and calculated items, shown after the field's own items. Rows and Columns
   * only; at most 50 per field.
   */
  customItems?: CustomItem[] | null;
  /**
   * Set when this field is a named set: its `uniqueName`. `hierarchy`, `allMember` and
   * `caption` are then the set's. A set's items are the set's tuples; `expanded`,
   * `members`, `sort` and the filters do not apply. Rows and Columns only.
   */
  set?: string | null;
}

/**
 * A group (`members`) or a calculated item (`formula`); exactly one of the two.
 * Excel on Windows has neither for cube PivotTables (grouping is greyed out; OLAP
 * Tools > MDX Calculated Member is the nearest), so this goes further than Windows.
 */
export interface CustomItem {
  /** Caption of the new item: 1 to 100 characters, unique within the field. */
  name: string;
  /**
   * Group: top-level items of the field (member unique names), shown as one item that
   * adds them up. Each item may be in one group only; grouped items are not shown on
   * their own. 2 to 5000 items.
   */
  members?: string[] | null;
  /**
   * Calculated item: an MDX expression, e.g. "[Geo].[City].&[1] - [Geo].[City].&[2]".
   * At most 2000 characters; statements and ';' are refused.
   */
  formula?: string | null;
}

/** Excel's Date Filters. Dates are "yyyy-mm-dd" and refer to the Mac's local calendar. */
export type DateOperator =
  | "equals"
  | "before"
  | "after"
  | "between"
  | "tomorrow"
  | "today"
  | "yesterday"
  | "nextWeek"
  | "thisWeek"
  | "lastWeek"
  | "nextMonth"
  | "thisMonth"
  | "lastMonth"
  | "nextQuarter"
  | "thisQuarter"
  | "lastQuarter"
  | "nextYear"
  | "thisYear"
  | "lastYear"
  | "yearToDate"
  // "All Dates in the Period"
  | "quarter1"
  | "quarter2"
  | "quarter3"
  | "quarter4"
  | "month1"
  | "month2"
  | "month3"
  | "month4"
  | "month5"
  | "month6"
  | "month7"
  | "month8"
  | "month9"
  | "month10"
  | "month11"
  | "month12";

export interface DateFilter {
  operator: DateOperator;
  /** Required for equals, before, after and between. */
  value?: string | null;
  /** Required for between: the last day, included. */
  value2?: string | null;
}

/** A measure defined in the report with an MDX formula (Windows: OLAP Tools > MDX Calculated Measure). */
export interface CalculatedMeasure {
  /** 1 to 100 characters; must not be the name of one of the cube's measures. Its unique name is "[Measures].[name]". */
  name: string;
  /** MDX expression; at most 2000 characters; statements and ';' are refused. */
  formula: string;
  /** Format string as the cube takes it ("#,##0.00", "Percent", …). Absent or "": the cube's default. */
  formatString?: string | null;
}

/**
 * Item selection (`members`), label filter and value filter all apply together. Each
 * one narrows the field's top-level items, and totals then count only the items that
 * are left (Excel's default for cube PivotTables).
 */
export interface FieldSort {
  /** Sort by the items' captions, or by a measure's values. */
  by: "caption" | "measure";
  /** Measure unique name; required when `by` is "measure". */
  measure?: string;
  order: "asc" | "desc";
}

export type LabelOperator =
  | "equals"
  | "notEquals"
  | "beginsWith"
  | "notBeginsWith"
  | "endsWith"
  | "notEndsWith"
  | "contains"
  | "notContains"
  | "greaterThan"
  | "greaterOrEqual"
  | "lessThan"
  | "lessOrEqual"
  | "between"
  | "notBetween";

export interface LabelFilter {
  operator: LabelOperator;
  /** Compared without regard to case. At most 200 characters. */
  value: string;
  /** Upper bound; required for "between" and "notBetween". */
  value2?: string;
}

export type ValueOperator =
  | "equals"
  | "notEquals"
  | "greaterThan"
  | "greaterOrEqual"
  | "lessThan"
  | "lessOrEqual"
  | "between"
  | "notBetween"
  | "top"
  | "bottom";

export interface ValueFilter {
  operator: ValueOperator;
  /** Measure unique name the condition is evaluated on. Need not be in Values. */
  measure: string;
  /** The threshold; for "top" and "bottom", how many items, what percent, or what sum. */
  value: number;
  /** Upper bound; required for "between" and "notBetween". */
  value2?: number;
  /** For "top" and "bottom" only. Default "items". */
  topBy?: "items" | "percent" | "sum";
}

/**
 * Excel's "Show Values As" list. The cube calculates each one; the column's header
 * caption says which. Kinds marked (B) need `baseField`, kinds marked (B, I) need
 * `baseField` and `baseItem`.
 */
export type ShowValuesAs =
  | "normal"
  | "percentOfGrandTotal"
  | "percentOfColumnTotal"
  | "percentOfRowTotal"
  | "percentOf" // (B, I) value / value of the base item
  | "percentOfParentRowTotal" // value / value of its parent on rows (outer item or parent member)
  | "percentOfParentColumnTotal" // the same on columns
  | "percentOfParentTotal" // (B) value / value of the base field's item, everything inside it at All
  | "differenceFrom" // (B, I) value - value of the base item; blank for totals and the base item
  | "percentDifferenceFrom" // (B, I) (value - base) / base
  | "runningTotalIn" // (B) sum from the first item of the base field to this one, in the cube's order
  | "percentRunningTotalIn" // (B) running total / base field total
  | "rankAscending" // (B) 1 = smallest, within the base field and the other fields' current items
  | "rankDescending" // (B) 1 = largest
  | "index"; // (value × grand total) / (row total × column total)

/**
 * A field in Values. The same measure may be placed more than once, each copy shown its
 * own way (a second plain copy is headed "<caption> 2").
 */
export interface ValueField {
  /** A measure's unique name, a calculated measure's ("[Measures].[name]"), or, for a KPI part, the KPI's value measure. */
  measure: string;
  caption: string;
  /** Set for a KPI's Value, Goal, Status or Trend. Then `showAs` must be "normal". */
  kpi?: { name: string; part: KpiPart } | null;
  /** Default "normal". */
  showAs?: ShowValuesAs;
  /** Hierarchy unique name of a field in Rows or Columns; required by the kinds marked (B). */
  baseField?: string | null;
  /**
   * For the kinds marked (B, I): "previous", "next", or a member unique name of the
   * base field. "previous" and "next" follow the cube's order of the items.
   */
  baseItem?: string | null;
  /**
   * Excel number format chosen by the user for this field. Applied by the pane, which
   * uses it instead of the cube's format; the helper ignores it. Absent or null: the
   * cube's format.
   */
  numberFormat?: string | null;
}

export interface LayoutOptions {
  /** Both grand totals at once; the two optional fields below override it per axis. */
  grandTotals: boolean;
  /** The Grand Total row at the bottom (Excel: on for columns). Absent: `grandTotals`. */
  grandTotalRow?: boolean;
  /** The Grand Total column on the right (Excel: on for rows). Absent: `grandTotals`. */
  grandTotalColumn?: boolean;
  subtotals: boolean;
  /**
   * Excel's Report Layout. Absent: "tabular", the only form before compact and outline
   * existed, so older reports keep their look. New reports start in "compact", Excel's
   * default.
   */
  form?: "compact" | "outline" | "tabular";
  /**
   * Where subtotals go: "top" puts each subtotal line before the items it adds up
   * (the helper orders the axis that way), "bottom" after them. Absent: "bottom".
   * Tabular form always uses "bottom", as in Excel.
   */
  subtotalsAt?: "top" | "bottom";
  /** Hide rows and columns with no data. */
  nonEmpty: boolean;
  /** Pane only, the helper ignores it: write outer row labels on every line. Not in compact form. */
  repeatLabels?: boolean;
  /** Pane only: text written in value cells that have no value. Absent or "": empty. */
  emptyCellText?: string;
}

export interface Layout {
  rows: PlacedField[];
  columns: PlacedField[];
  values: ValueField[];
  filters: PlacedField[];
  options: LayoutOptions;
  /** Measures defined in this report; usable in Values, sorts and value filters. At most 50. */
  calculatedMeasures?: CalculatedMeasure[] | null;
}

export interface QueryRequest {
  catalog: string;
  cube: string;
  layout: Layout;
}

/**
 * "Show Details" for one value cell, like double-clicking a number in a PivotTable on
 * Windows: the fact rows behind it. `POST /connections/{id}/details`.
 */
export interface DetailsRequest {
  catalog: string;
  cube: string;
  layout: Layout;
  /**
   * The members behind the cell: from its row header cells and column header cells,
   * every `MemberRef.uniqueName` that is not a total. A total contributes nothing.
   */
  members: string[];
  /** The value field's measure (its `measure`, not the shown-as caption). */
  measure: string;
  /** Default 1000, maximum 100000. */
  maxRows?: number;
}

export interface DetailsResponse {
  columns: string[];
  rows: CellValue[][];
  /** True when there are more rows than `maxRows`. */
  truncated: boolean;
}

export interface MemberRef {
  uniqueName: string;
  hierarchy: string;
  caption: string;
  level: number;
  hasChildren: boolean;
  expanded: boolean;
  /** A group or calculated item of the report: no drill-down and no Show Details. */
  calculated?: boolean;
}

export type HeaderKind = "member" | "subtotal" | "grandTotal" | "measure" | "blank";

export interface HeaderCell {
  /** Text to write in the cell. "" for a repeated label that is suppressed. */
  text: string;
  /** Indent level, 0 or more. */
  indent: number;
  kind: HeaderKind;
  /** The member behind this cell; set for "member" and "blank", else null. */
  member: MemberRef | null;
}

export type CellValue = number | string | boolean | null;
export type LineKind = "data" | "subtotal" | "grandTotal";

/**
 * A query result, ready to write. With F = filterRows.length, H =
 * colHeaderRows.length, N = rowHeaders.length, R = rowFieldCaptions.length and C =
 * colKinds.length, the sheet block is:
 *
 *   F rows   caption | value                      (page filters)
 *   1 row    blank                                (only when F > 0)
 *   H rows   R corner cells | C header cells      (last row's corner = rowFieldCaptions)
 *   N rows   R row-header cells | C values
 */
export interface Grid {
  filterRows: { caption: string; value: string }[];
  /** Always at least one entry. */
  rowFieldCaptions: string[];
  /** Always at least one row; each row has C cells. */
  colHeaderRows: HeaderCell[][];
  /** N rows, each with R cells. N may be 0 (no data). */
  rowHeaders: HeaderCell[][];
  /** N rows of C values. */
  values: CellValue[][];
  /** Excel number formats. Index 0 is always "General". */
  formatPalette: string[];
  /** N rows of C indexes into `formatPalette`. */
  cellFormats: number[][];
  rowKinds: LineKind[];
  colKinds: LineKind[];
  /**
   * C entries: for each column, the index in `layout.values` of the value field it
   * shows, or -1 when it shows none of them (the cube's default measure). Read this
   * instead of computing it from the column's position: NON EMPTY can drop single
   * column-member × measure columns.
   */
  colValueIndex: number[];
  stats: {
    rows: number;
    columns: number;
    elapsedMs: number;
    mdx: string;
  };
}
