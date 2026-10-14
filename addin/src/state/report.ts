// The report being edited in the pane: its definition, whether it is saved in the
// workbook yet, and where an unwritten one will go. Pure reducer, wrapping the layout
// reducer.

import type { CellAddress, Extent, ReportDef, SheetSelection } from "../sheet/host";
import type {
  FieldSort,
  LabelFilter,
  LabelOperator,
  Layout,
  PlacedField,
  ShowValuesAs,
  ValueField,
  ValueFilter,
  ValueOperator,
} from "../types";
import { DEFAULT_OPTIONS, emptyLayout, layoutReducer, type LayoutAction } from "./layout";

export interface Current {
  def: ReportDef;
  /** Saved in the workbook: it has been written to the sheet at least once. */
  persisted: boolean;
  /** Where a report that has no cells yet will be placed. Null means "the active cell
   *  at the moment of the first write". Ignored once the report has an extent. */
  anchor: CellAddress | null;
  /** Goes up whenever the definition changes in a way that needs a new query. */
  revision: number;
  /** True when the last change should be queried at once rather than after the
   *  typing delay (expand and collapse feel wrong with a delay). */
  urgent: boolean;
}

export type CurrentAction =
  /** Make a saved report, or a new draft, the one being edited. */
  | { type: "open"; def: ReportDef; persisted: boolean; anchor?: CellAddress | null }
  | { type: "close" }
  | { type: "layout"; action: LayoutAction; urgent?: boolean }
  /** Point the report at another connection, keeping database, cube and layout. */
  | { type: "rebind"; connectionId: string }
  /** Change the database: the cube and the layout no longer apply. */
  | { type: "setCatalog"; catalog: string }
  /** Change the cube: the layout no longer applies. */
  | { type: "setCube"; cube: string }
  /** Fill in a database or cube the user has not chosen yet. Not a change to query. */
  | { type: "defaults"; catalog?: string; cube?: string }
  | { type: "pin"; anchor: CellAddress }
  /** The report is in the workbook now. With `id`, only if that is the report being
   *  edited: the result of a run may arrive after the pane has moved to another one. */
  | { type: "saved"; id?: string }
  /** Ask for a re-query without changing anything (Refresh). */
  | { type: "touch" };

export function newReportId(random: () => number = Math.random): string {
  // Used inside a workbook name (CubeReport_<id>), so letters and digits only.
  let id = "";
  while (id.length < 10) id += Math.floor(random() * 36).toString(36);
  return id;
}

export function newReportDef(connectionId: string, catalog = "", cube = ""): ReportDef {
  return { id: newReportId(), connectionId, catalog, cube, layout: emptyLayout() };
}

export function currentReducer(state: Current | null, action: CurrentAction): Current | null {
  if (action.type === "open") {
    return {
      def: action.def,
      persisted: action.persisted,
      anchor: action.anchor ?? null,
      revision: 0,
      urgent: false,
    };
  }
  if (action.type === "close") return null;
  if (!state) return state;

  switch (action.type) {
    case "layout": {
      const layout = layoutReducer(state.def.layout, action.action);
      if (layout === state.def.layout) return state;
      return {
        ...state,
        def: { ...state.def, layout },
        revision: state.revision + 1,
        urgent: action.urgent === true,
      };
    }
    case "rebind":
      if (state.def.connectionId === action.connectionId) return state;
      return {
        ...state,
        def: { ...state.def, connectionId: action.connectionId },
        revision: state.revision + 1,
        urgent: true,
      };
    case "setCatalog":
      if (state.def.catalog === action.catalog) return state;
      return {
        ...state,
        def: { ...state.def, catalog: action.catalog, cube: "", layout: emptyLayout() },
        revision: state.revision + 1,
        urgent: false,
      };
    case "setCube":
      if (state.def.cube === action.cube) return state;
      return {
        ...state,
        def: { ...state.def, cube: action.cube, layout: emptyLayout() },
        revision: state.revision + 1,
        urgent: false,
      };
    case "defaults": {
      const catalog = state.def.catalog === "" ? action.catalog ?? "" : state.def.catalog;
      const cube = state.def.cube === "" ? action.cube ?? "" : state.def.cube;
      if (catalog === state.def.catalog && cube === state.def.cube) return state;
      return { ...state, def: { ...state.def, catalog, cube } };
    }
    case "pin":
      return { ...state, anchor: action.anchor };
    case "saved":
      if (action.id !== undefined && action.id !== state.def.id) return state;
      return state.persisted ? state : { ...state, persisted: true };
    case "touch":
      return { ...state, revision: state.revision + 1, urgent: true };
  }
}

// ---------------------------------------------------------------------------
// What the button next to the report's location does.
// ---------------------------------------------------------------------------

export type PlaceDecision =
  /** The selected cell belongs to a report: nothing can be put there. */
  | { kind: "blocked"; inside: "this" | "other" }
  /** The report being edited has no cells: it goes to the selected cell, fields and all. */
  | { kind: "place" }
  /** It is on a sheet: the button starts another report at the selected cell. */
  | { kind: "new" };

/**
 * `extent` must be the report's location read from the workbook just now, not a
 * remembered one: a report whose sheet was deleted still exists but has no cells, and
 * the button must then place that report rather than start an empty one.
 */
export function decidePlace(
  reportId: string,
  extent: Extent | null,
  selection: Pick<SheetSelection, "report">,
): PlaceDecision {
  if (selection.report) {
    return { kind: "blocked", inside: selection.report.id === reportId ? "this" : "other" };
  }
  return extent === null ? { kind: "place" } : { kind: "new" };
}

/** The button's wording for the same two cases. */
export function placeLabel(extent: Extent | null): string {
  return extent === null ? "Place here" : "New report here";
}

// ---------------------------------------------------------------------------
// Reading a definition back from the workbook. The file may come from another Mac,
// an older version, or be hand-edited, so nothing in it is trusted to have the right
// shape. A definition that cannot be read is skipped, never thrown on.
// ---------------------------------------------------------------------------

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function strings(value: unknown): string[] {
  return Array.isArray(value) ? value.filter((v): v is string => typeof v === "string") : [];
}

const LABEL_OPERATORS: readonly LabelOperator[] = [
  "equals",
  "notEquals",
  "beginsWith",
  "notBeginsWith",
  "endsWith",
  "notEndsWith",
  "contains",
  "notContains",
  "greaterThan",
  "greaterOrEqual",
  "lessThan",
  "lessOrEqual",
  "between",
  "notBetween",
];
const VALUE_OPERATORS: readonly ValueOperator[] = [
  "equals",
  "notEquals",
  "greaterThan",
  "greaterOrEqual",
  "lessThan",
  "lessOrEqual",
  "between",
  "notBetween",
  "top",
  "bottom",
];
const SHOW_AS: readonly ShowValuesAs[] = [
  "normal",
  "percentOfGrandTotal",
  "percentOfColumnTotal",
  "percentOfRowTotal",
];
const TWO_BOUNDS = ["between", "notBetween"];

function isOneOf<T extends string>(list: readonly T[], value: unknown): value is T {
  return typeof value === "string" && (list as readonly string[]).indexOf(value) >= 0;
}

function finite(value: unknown): value is number {
  return typeof value === "number" && isFinite(value);
}

function parseSort(value: unknown): FieldSort | null {
  if (!isRecord(value) || (value.order !== "asc" && value.order !== "desc")) return null;
  if (value.by === "caption") return { by: "caption", order: value.order };
  if (value.by === "measure" && typeof value.measure === "string" && value.measure !== "") {
    return { by: "measure", measure: value.measure, order: value.order };
  }
  return null;
}

function parseLabelFilter(value: unknown): LabelFilter | null {
  if (!isRecord(value) || !isOneOf(LABEL_OPERATORS, value.operator)) return null;
  if (typeof value.value !== "string") return null;
  if (TWO_BOUNDS.indexOf(value.operator) >= 0) {
    if (typeof value.value2 !== "string") return null;
    return { operator: value.operator, value: value.value, value2: value.value2 };
  }
  return { operator: value.operator, value: value.value };
}

function parseValueFilter(value: unknown): ValueFilter | null {
  if (!isRecord(value) || !isOneOf(VALUE_OPERATORS, value.operator)) return null;
  if (typeof value.measure !== "string" || value.measure === "" || !finite(value.value)) return null;
  const out: ValueFilter = { operator: value.operator, measure: value.measure, value: value.value };
  if (TWO_BOUNDS.indexOf(value.operator) >= 0) {
    if (!finite(value.value2)) return null;
    out.value2 = value.value2;
  }
  if (value.operator === "top" || value.operator === "bottom") {
    out.topBy = value.topBy === "percent" || value.topBy === "sum" ? value.topBy : "items";
  }
  return out;
}

function placedFields(value: unknown): PlacedField[] {
  if (!Array.isArray(value)) return [];
  const out: PlacedField[] = [];
  for (const item of value) {
    if (!isRecord(item) || typeof item.hierarchy !== "string" || item.hierarchy === "") continue;
    const members = strings(item.members);
    const captions = strings(item.memberCaptions);
    const field: PlacedField = {
      hierarchy: item.hierarchy,
      caption: typeof item.caption === "string" ? item.caption : item.hierarchy,
      allMember: typeof item.allMember === "string" ? item.allMember : null,
      expanded: strings(item.expanded),
      members,
      memberCaptions: members.map((m, i) => captions[i] ?? m),
    };
    // Optional settings that cannot be read are dropped, never guessed.
    const sort = parseSort(item.sort);
    if (sort) field.sort = sort;
    const labelFilter = parseLabelFilter(item.labelFilter);
    if (labelFilter) field.labelFilter = labelFilter;
    const valueFilter = parseValueFilter(item.valueFilter);
    if (valueFilter) field.valueFilter = valueFilter;
    out.push(field);
  }
  return out;
}

function valueFields(value: unknown): ValueField[] {
  if (!Array.isArray(value)) return [];
  const out: ValueField[] = [];
  for (const item of value) {
    if (!isRecord(item) || typeof item.measure !== "string" || item.measure === "") continue;
    const field: ValueField = {
      measure: item.measure,
      caption: typeof item.caption === "string" ? item.caption : item.measure,
    };
    if (isOneOf(SHOW_AS, item.showAs) && item.showAs !== "normal") field.showAs = item.showAs;
    if (typeof item.numberFormat === "string" && item.numberFormat.trim() !== "" && item.numberFormat.length <= 100) {
      field.numberFormat = item.numberFormat;
    }
    out.push(field);
  }
  return out;
}

export function parseLayout(value: unknown): Layout {
  if (!isRecord(value)) return emptyLayout();
  const options = isRecord(value.options) ? value.options : {};
  const flag = (key: keyof typeof DEFAULT_OPTIONS) =>
    typeof options[key] === "boolean" ? (options[key] as boolean) : DEFAULT_OPTIONS[key];

  // A hierarchy may sit in one area only; keep the first place it appears.
  const seen = new Set<string>();
  const unique = (fields: PlacedField[]) =>
    fields.filter((f) => (seen.has(f.hierarchy) ? false : (seen.add(f.hierarchy), true)));
  const rows = unique(placedFields(value.rows));
  const columns = unique(placedFields(value.columns));
  const filters = unique(placedFields(value.filters));

  // The same measure may be in Values twice, shown two ways (plain and as a percent).
  const shown = new Set<string>();
  const values = valueFields(value.values).filter((v) => {
    const key = v.measure + "\u0001" + (v.showAs ?? "normal");
    return shown.has(key) ? false : (shown.add(key), true);
  });

  return {
    rows,
    columns,
    values,
    filters,
    options: {
      grandTotals: flag("grandTotals"),
      subtotals: flag("subtotals"),
      nonEmpty: flag("nonEmpty"),
    },
  };
}

const ID_PATTERN = /^[A-Za-z0-9]{1,40}$/;

/** A stored definition as a `ReportDef`, or null when it cannot be used. */
export function parseReportDef(json: unknown, expectedId?: string): ReportDef | null {
  let value: unknown = json;
  if (typeof json === "string") {
    try {
      value = JSON.parse(json);
    } catch {
      return null;
    }
  }
  if (!isRecord(value)) return null;
  const id = typeof value.id === "string" ? value.id : expectedId;
  if (typeof id !== "string" || !ID_PATTERN.test(id)) return null;
  if (expectedId !== undefined && id !== expectedId) return null;
  return {
    id,
    connectionId: typeof value.connectionId === "string" ? value.connectionId : "",
    catalog: typeof value.catalog === "string" ? value.catalog : "",
    cube: typeof value.cube === "string" ? value.cube : "",
    layout: parseLayout(value.layout),
  };
}

export function serializeReportDef(def: ReportDef): string {
  return JSON.stringify({
    id: def.id,
    connectionId: def.connectionId,
    catalog: def.catalog,
    cube: def.cube,
    layout: def.layout,
  });
}
