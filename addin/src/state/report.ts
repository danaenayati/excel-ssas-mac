// The report being edited in the pane: its definition, whether it is saved in the
// workbook yet, and where an unwritten one will go. Pure reducer, wrapping the layout
// reducer.

import type { CellAddress, Extent, ReportDef, SheetSelection } from "../sheet/host";
import type {
  CalculatedMeasure,
  CustomItem,
  DateFilter,
  DateOperator,
  FieldSort,
  KpiPart,
  LabelFilter,
  LabelOperator,
  Layout,
  PlacedField,
  ShowValuesAs,
  ValueField,
  ValueFilter,
  ValueOperator,
} from "../types";
import {
  DEFAULT_OPTIONS,
  MAX_EMPTY_CELL_TEXT,
  emptyLayout,
  layoutReducer,
  needsBaseField,
  needsBaseItem,
  type LayoutAction,
} from "./layout";

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
// Which report the pane edits on the sheet in view.
// ---------------------------------------------------------------------------

export type SheetChoice =
  /** Go on with the report being edited. */
  | { kind: "keep" }
  /** Edit this report, which is on the sheet. */
  | { kind: "open"; id: string }
  /** Start a new report at the selected cell. */
  | { kind: "new" };

/**
 * The report the pane should edit when the user is on `sheetId` and the selected cell
 * is in no report. As in Excel, the field list belongs to the sheet in view: a field
 * ticked there must never change a report on another sheet, where nothing would show.
 *
 * The report being edited stays when it is on this sheet, when it has no cells (its
 * sheet was deleted; its next update puts it at the selected cell), or when it is a new
 * report that already has fields. Otherwise the first report on this sheet is opened,
 * and on a sheet without one a new report is started, unless an empty new one is
 * already being edited. `current` is null at start-up. `locations` must be read from
 * the workbook just now.
 */
export function reportForSheet(
  current: Pick<Current, "def" | "persisted"> | null,
  sheetId: string,
  reports: ReportDef[],
  locations: Record<string, Extent | null>,
): SheetChoice {
  if (current) {
    const where = current.persisted ? locations[current.def.id] ?? null : null;
    if (current.persisted && (where === null || where.sheetId === sheetId)) return { kind: "keep" };
    const layout = current.def.layout;
    const untouched =
      layout.rows.length + layout.columns.length + layout.values.length + layout.filters.length === 0;
    if (!current.persisted && !untouched) return { kind: "keep" };
  }
  const here = reports.find((d) => {
    const where = locations[d.id];
    return where !== null && where !== undefined && where.sheetId === sheetId;
  });
  if (here) return { kind: "open", id: here.id };
  return current && !current.persisted ? { kind: "keep" } : { kind: "new" };
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
  "percentOf",
  "percentOfParentRowTotal",
  "percentOfParentColumnTotal",
  "percentOfParentTotal",
  "differenceFrom",
  "percentDifferenceFrom",
  "runningTotalIn",
  "percentRunningTotalIn",
  "rankAscending",
  "rankDescending",
  "index",
];
const TWO_BOUNDS = ["between", "notBetween"];
const DATE_OPERATORS: readonly DateOperator[] = [
  "equals", "before", "after", "between", "tomorrow", "today", "yesterday",
  "nextWeek", "thisWeek", "lastWeek", "nextMonth", "thisMonth", "lastMonth",
  "nextQuarter", "thisQuarter", "lastQuarter", "nextYear", "thisYear", "lastYear", "yearToDate",
  "quarter1", "quarter2", "quarter3", "quarter4",
  "month1", "month2", "month3", "month4", "month5", "month6",
  "month7", "month8", "month9", "month10", "month11", "month12",
];
const KPI_PARTS: readonly KpiPart[] = ["value", "goal", "status", "trend"];
const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;
const MAX_NAME = 100;

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

function parseDateFilter(value: unknown): DateFilter | null {
  if (!isRecord(value) || !isOneOf(DATE_OPERATORS, value.operator)) return null;
  const needsDate = ["equals", "before", "after", "between"].indexOf(value.operator) >= 0;
  if (!needsDate) return { operator: value.operator };
  const date = (v: unknown) => (typeof v === "string" && ISO_DATE.test(v) ? v : null);
  const first = date(value.value);
  if (!first) return null;
  if (value.operator !== "between") return { operator: value.operator, value: first };
  const second = date(value.value2);
  return second ? { operator: "between", value: first, value2: second } : null;
}

function newName(value: unknown): string | null {
  return typeof value === "string" && value.trim() !== "" && value.length <= MAX_NAME ? value.trim() : null;
}

function parseCustomItems(value: unknown): CustomItem[] {
  if (!Array.isArray(value)) return [];
  const out: CustomItem[] = [];
  for (const item of value) {
    if (!isRecord(item)) continue;
    const name = newName(item.name);
    if (!name || out.some((o) => o.name.toLowerCase() === name.toLowerCase())) continue;
    const members = strings(item.members);
    if (members.length >= 2 && !item.formula) out.push({ name, members });
    else if (typeof item.formula === "string" && item.formula.trim() !== "" && members.length === 0) {
      out.push({ name, formula: item.formula });
    }
  }
  return out;
}

function parseCalculatedMeasures(value: unknown): CalculatedMeasure[] {
  if (!Array.isArray(value)) return [];
  const out: CalculatedMeasure[] = [];
  for (const item of value) {
    if (!isRecord(item)) continue;
    const name = newName(item.name);
    if (!name || typeof item.formula !== "string" || item.formula.trim() === "") continue;
    if (out.some((o) => o.name.toLowerCase() === name.toLowerCase())) continue;
    const measure: CalculatedMeasure = { name, formula: item.formula };
    if (typeof item.formatString === "string" && item.formatString !== "") measure.formatString = item.formatString;
    out.push(measure);
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
    const dateFilter = parseDateFilter(item.dateFilter);
    if (dateFilter && !field.labelFilter) field.dateFilter = dateFilter;
    const customItems = parseCustomItems(item.customItems);
    if (customItems.length > 0) field.customItems = customItems;
    if (typeof item.set === "string" && item.set !== "") {
      // A set's items are its own: nothing else about the field applies.
      out.push({ hierarchy: field.hierarchy, caption: field.caption, allMember: field.allMember,
        expanded: [], members: [], memberCaptions: [], set: item.set });
      continue;
    }
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
    if (isRecord(item.kpi) && typeof item.kpi.name === "string" && item.kpi.name !== "" && isOneOf(KPI_PARTS, item.kpi.part)) {
      if (out.some((o) => o.kpi && o.kpi.name === (item.kpi as { name: string }).name && o.kpi.part === (item.kpi as { part: string }).part)) continue;
      field.kpi = { name: item.kpi.name, part: item.kpi.part };
      out.push(field);
      continue; // a KPI's part is shown as it is
    }
    if (isOneOf(SHOW_AS, item.showAs) && item.showAs !== "normal") {
      // A kind whose base is missing cannot be shown as asked: it reads as No Calculation.
      const baseField = typeof item.baseField === "string" && item.baseField !== "" ? item.baseField : null;
      const baseItem = typeof item.baseItem === "string" && item.baseItem !== "" ? item.baseItem : null;
      const complete =
        (!needsBaseField(item.showAs) || baseField !== null) && (!needsBaseItem(item.showAs) || baseItem !== null);
      if (complete) {
        field.showAs = item.showAs;
        if (needsBaseField(item.showAs) && baseField) field.baseField = baseField;
        if (needsBaseItem(item.showAs) && baseItem) field.baseItem = baseItem;
      }
    }
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
  const flag = (key: "grandTotals" | "subtotals" | "nonEmpty") =>
    typeof options[key] === "boolean" ? (options[key] as boolean) : DEFAULT_OPTIONS[key];

  // A hierarchy may sit in one area only; keep the first place it appears.
  const seen = new Set<string>();
  const unique = (fields: PlacedField[]) =>
    fields.filter((f) => (seen.has(f.hierarchy) ? false : (seen.add(f.hierarchy), true)));
  const rows = unique(placedFields(value.rows));
  const columns = unique(placedFields(value.columns));
  const filters = unique(placedFields(value.filters));

  // The same measure may be in Values any number of times, each copy shown its own way.
  const values = valueFields(value.values);
  const calculatedMeasures = parseCalculatedMeasures(value.calculatedMeasures);

  const layout: Layout = {
    rows,
    columns,
    values,
    filters,
    options: parsedOptions(options, flag),
  };
  if (calculatedMeasures.length > 0) layout.calculatedMeasures = calculatedMeasures;
  return layout;
}

function parsedOptions(
  options: Record<string, unknown>,
  flag: (key: "grandTotals" | "subtotals" | "nonEmpty") => boolean,
): Layout["options"] {
  const out: Layout["options"] = {
    grandTotals: flag("grandTotals"),
    subtotals: flag("subtotals"),
    nonEmpty: flag("nonEmpty"),
  };
  // Optional settings are kept only when they are what they claim to be.
  if (typeof options.grandTotalRow === "boolean") out.grandTotalRow = options.grandTotalRow;
  if (typeof options.grandTotalColumn === "boolean") out.grandTotalColumn = options.grandTotalColumn;
  if (options.repeatLabels === true) out.repeatLabels = true;
  if (typeof options.emptyCellText === "string" && options.emptyCellText !== "") {
    out.emptyCellText = options.emptyCellText.slice(0, MAX_EMPTY_CELL_TEXT);
  }
  return out;
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
