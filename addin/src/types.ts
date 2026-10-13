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
}

export interface ValueField {
  measure: string;
  caption: string;
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
