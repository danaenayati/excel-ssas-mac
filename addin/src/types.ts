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
}

export interface Dimension {
  uniqueName: string;
  caption: string;
  hierarchies: Hierarchy[];
}

export interface Fields {
  measures: Measure[];
  dimensions: Dimension[];
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

export type ShowValuesAs = "normal" | "percentOfGrandTotal" | "percentOfColumnTotal" | "percentOfRowTotal";

export interface ValueField {
  measure: string;
  caption: string;
  /** Default "normal". The cube calculates the percentage; its header caption says which. */
  showAs?: ShowValuesAs;
  /**
   * Excel number format chosen by the user for this field. Applied by the pane, which
   * uses it instead of the cube's format; the helper ignores it. Absent or null: the
   * cube's format.
   */
  numberFormat?: string | null;
}

export interface LayoutOptions {
  grandTotals: boolean;
  subtotals: boolean;
  /** Hide rows and columns with no data. */
  nonEmpty: boolean;
}

export interface Layout {
  rows: PlacedField[];
  columns: PlacedField[];
  values: ValueField[];
  filters: PlacedField[];
  options: LayoutOptions;
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
  stats: {
    rows: number;
    columns: number;
    elapsedMs: number;
    mdx: string;
  };
}
