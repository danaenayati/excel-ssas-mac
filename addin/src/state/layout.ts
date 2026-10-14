// The pivot layout reducer. Pure: no React, no I/O. Every action returns the same
// object when it changes nothing, so callers can tell "no change" by identity and
// skip a re-query.

import { DEFAULT_STYLE, isStyleName } from "../sheet/styles";
import type {
  CalculatedMeasure,
  CustomItem,
  DateFilter,
  FieldSort,
  KpiPart,
  LabelFilter,
  Layout,
  LayoutOptions,
  PlacedField,
  ShowValuesAs,
  ValueField,
  ValueFilter,
} from "../types";

/** Areas that hold hierarchies. */
export type FieldArea = "rows" | "columns" | "filters";
export type Area = FieldArea | "values";

export const FIELD_AREAS: readonly FieldArea[] = ["filters", "columns", "rows"];
export const AREAS: readonly Area[] = ["filters", "columns", "rows", "values"];

export const AREA_LABELS: Record<Area, string> = {
  filters: "Filters",
  columns: "Columns",
  rows: "Rows",
  values: "Values",
};

export interface MeasureInput {
  uniqueName: string;
  caption: string;
}

export interface HierarchyInput {
  uniqueName: string;
  caption: string;
  allMember: string | null;
  /** Place only this level of a user hierarchy (the field is then captioned with the level's name). */
  level?: { uniqueName: string; number: number; caption: string } | null;
}

export interface KpiInput {
  name: string;
  caption: string;
  /** The KPI's value measure. */
  measure: string;
}

/** A named set that can be placed: its items come from one hierarchy. */
export interface SetInput {
  uniqueName: string;
  caption: string;
  hierarchy: string;
  allMember: string | null;
  /** All hierarchies of its tuples, when it spans several (the first is `hierarchy`). */
  hierarchies?: string[];
  allMembers?: (string | null)[];
}

/** Excel's words for a KPI's parts, as its field list shows them. */
export const KPI_PART_LABELS: Record<KpiPart, string> = {
  value: "Value",
  goal: "Goal",
  status: "Status",
  trend: "Trend",
};

export type LayoutAction =
  /** Tick a measure: goes to Values. No-op when already there. */
  | { type: "addMeasure"; measure: MeasureInput; index?: number }
  /** Tick a KPI's Value, Goal, Status or Trend: goes to Values. No-op when already there. */
  | { type: "addKpi"; kpi: KpiInput; part: KpiPart; index?: number }
  /** Another copy of a value field, right after it, to be shown another way. */
  | { type: "duplicateValue"; key: string }
  /** Untick a measure in the field list: every plain copy of it leaves Values, as in Excel. */
  | { type: "removeMeasure"; measure: string }
  /** Tick a named set: goes to Rows unless an area is given. No-op when its field is placed. */
  | { type: "addSet"; set: SetInput; area?: "rows" | "columns"; index?: number }
  /** Tick a hierarchy: goes to Rows unless an area is given. If it is already placed
   *  and an area is given, it is moved there (keeping its drill state and filter). */
  | { type: "addHierarchy"; hierarchy: HierarchyInput; area?: FieldArea; index?: number }
  /** Remove a measure or hierarchy, by unique name, from whichever area holds it. */
  | { type: "remove"; name: string }
  /** Move to an area. `index` is the position in the target list after the field has
   *  been taken out of its old place; omitted means "at the end". */
  | { type: "move"; name: string; area: Area; index?: number }
  /** Move one step up (-1) or down (+1) inside its own area. */
  | { type: "shift"; name: string; delta: -1 | 1 }
  | { type: "setMembers"; hierarchy: string; members: string[]; captions: string[] }
  | { type: "expand"; hierarchy: string; member: string }
  | { type: "collapse"; hierarchy: string; member: string; descendants?: string[] }
  | { type: "setOption"; option: BooleanOption; value: boolean }
  /** Excel's Grand Totals choice: the Grand Total row (bottom) and column (right). */
  | { type: "setGrandTotals"; row: boolean; column: boolean }
  /** "For empty cells show": text written into value cells without a value. */
  | { type: "setEmptyCellText"; text: string }
  /** Excel's Report Layout: Compact, Outline or Tabular Form. */
  | { type: "setForm"; form: ReportForm }
  /** Excel's Subtotals choice (Design tab): none, at the bottom or at the top of each group. */
  | { type: "setSubtotals"; at: SubtotalsChoice }
  /** "Expand Entire Field": these members of the field, one level deeper. */
  | { type: "expandField"; hierarchy: string; members: string[] }
  /** "Collapse Entire Field". */
  | { type: "collapseField"; hierarchy: string }
  /** Order of a field's items; null for the cube's own order. */
  | { type: "setSort"; hierarchy: string; sort: FieldSort | null }
  | { type: "setLabelFilter"; hierarchy: string; filter: LabelFilter | null }
  | { type: "setValueFilter"; hierarchy: string; filter: ValueFilter | null }
  /** Excel's Date Filters. A date filter and a label filter replace each other, as in Excel. */
  | { type: "setDateFilter"; hierarchy: string; filter: DateFilter | null }
  /** A field's groups and calculated items, all at once. */
  | { type: "setCustomItems"; hierarchy: string; items: CustomItem[] }
  /** Add a calculated measure, or replace the one called `previousName` (a rename keeps its uses). */
  | { type: "setCalculatedMeasure"; measure: CalculatedMeasure; previousName?: string }
  /** Delete a calculated measure, with its value fields, and the sorts and filters that use it. */
  | { type: "removeCalculatedMeasure"; name: string }
  /** "Clear Filter From <field>": items, label filter and value filter. Not the sort. */
  | { type: "clearFilters"; hierarchy: string }
  /**
   * Show Values As, with the base field and item the kind needs (others are dropped).
   * `measure` is the value field's key (`valueKey`), which is its measure for the first copy.
   */
  | { type: "setShowAs"; measure: string; showAs: ShowValuesAs; baseField?: string | null; baseItem?: string | null }
  /** A value field's number format; null for the cube's own. `measure` is the field's key. */
  | { type: "setNumberFormat"; measure: string; format: string | null }
  /**
   * Value Field Settings > Custom Name; "" or null gives back the cube's caption.
   * `measure` is the field's key. The pane refuses a name in use first (state/valueNames.ts).
   */
  | { type: "setCustomName"; measure: string; name: string | null }
  /** Excel's PivotTable Styles: a name from sheet/styles.ts; null for the default (Light 16). */
  | { type: "setStyle"; style: string | null }
  /** Excel's Style Options: Row Headers, Column Headers, Banded Rows, Banded Columns. */
  | { type: "setStyleOption"; option: StyleOption; value: boolean }
  /**
   * Insert Slicer: the field gets a slicer, at the end. A field in Rows, Columns or
   * Filters moves there, keeping its chosen items. No-op for a named set, a field that
   * already has a slicer, or a 17th. Its items are chosen with `setMembers`.
   */
  | { type: "addSlicer"; hierarchy: HierarchyInput }
  | { type: "removeSlicer"; hierarchy: string }
  /** The Insert Slicers dialog: exactly these fields have slicers, in this order (as `addSlicer`). */
  | { type: "setSlicers"; hierarchies: HierarchyInput[] }
  | { type: "replace"; layout: Layout };

/** Most slicers a report has, as the helper allows. */
export const MAX_SLICERS = 16;

/** Excel's four Style Options (Design tab). */
export type StyleOption = "rowHeaders" | "columnHeaders" | "bandedRows" | "bandedColumns";

/** What each Style Option is when a report does not say: Excel's defaults. */
export const STYLE_OPTION_DEFAULTS: Record<StyleOption, boolean> = {
  rowHeaders: true,
  columnHeaders: true,
  bandedRows: false,
  bandedColumns: false,
};

/** The on/off options that `setOption` toggles. */
export type BooleanOption =
  | "grandTotals"
  | "grandTotalRow"
  | "grandTotalColumn"
  | "subtotals"
  | "nonEmpty"
  | "repeatLabels"
  | "includeFilteredItems";

/** Longest "For empty cells show" text. */
export const MAX_EMPTY_CELL_TEXT = 20;

/** Longest custom name of a value field, as Excel allows for a PivotTable field name. */
export const MAX_CUSTOM_NAME = 255;

/** Show Values As kinds that need a base field, and those that need a base item too. */
const NEEDS_BASE_FIELD: readonly ShowValuesAs[] = [
  "percentOf",
  "percentOfParentTotal",
  "differenceFrom",
  "percentDifferenceFrom",
  "runningTotalIn",
  "percentRunningTotalIn",
  "rankAscending",
  "rankDescending",
];
const NEEDS_BASE_ITEM: readonly ShowValuesAs[] = ["percentOf", "differenceFrom", "percentDifferenceFrom"];

export function needsBaseField(showAs: ShowValuesAs | undefined): boolean {
  return showAs !== undefined && NEEDS_BASE_FIELD.indexOf(showAs) >= 0;
}

export function needsBaseItem(showAs: ShowValuesAs | undefined): boolean {
  return showAs !== undefined && NEEDS_BASE_ITEM.indexOf(showAs) >= 0;
}

/** Whether the Grand Total row (bottom) and column (right) are shown. */
export function grandTotalsShown(options: LayoutOptions): { row: boolean; column: boolean } {
  return {
    row: options.grandTotalRow ?? options.grandTotals,
    column: options.grandTotalColumn ?? options.grandTotals,
  };
}

/** Excel's Report Layout. */
export type ReportForm = NonNullable<LayoutOptions["form"]>;
export const REPORT_FORMS: readonly ReportForm[] = ["compact", "outline", "tabular"];

/** Excel's three Subtotals choices: "Do Not Show", "at Bottom of Group", "at Top of Group". */
export type SubtotalsChoice = "none" | "bottom" | "top";

/** The form a report is shown in. Absent: tabular, the only form before the others existed. */
export function reportForm(options: LayoutOptions): ReportForm {
  return options.form === "compact" || options.form === "outline" ? options.form : "tabular";
}

/**
 * Where subtotals are shown, as the sheet shows them: tabular form has no subtotals at
 * the top of a group (Excel greys that choice out), so there it is "bottom" whatever is
 * stored. The stored choice is kept for when the form changes back, as in Excel.
 */
export function subtotalsChoice(options: LayoutOptions): SubtotalsChoice {
  if (!options.subtotals) return "none";
  return options.subtotalsAt === "top" && reportForm(options) !== "tabular" ? "top" : "bottom";
}

/**
 * What a stored definition reads as where it says nothing: the options from before
 * Report Layout existed (tabular form, no subtotals), so an old report keeps its look.
 */
export const DEFAULT_OPTIONS: LayoutOptions = {
  grandTotals: true,
  subtotals: false,
  nonEmpty: true,
};

/** A new report starts as Excel starts one: Compact Form, subtotals at the top of each group. */
export const NEW_REPORT_OPTIONS: LayoutOptions = {
  ...DEFAULT_OPTIONS,
  subtotals: true,
  form: "compact",
  subtotalsAt: "top",
};

/** No fields, and the options a stored definition without any reads as. */
export function emptyLayout(): Layout {
  return { rows: [], columns: [], values: [], filters: [], options: { ...DEFAULT_OPTIONS } };
}

/** The layout of a new report (or one whose cube changed): no fields, Excel's defaults. */
export function newLayout(): Layout {
  return { ...emptyLayout(), options: { ...NEW_REPORT_OPTIONS } };
}

/** True when there is nothing to query: no field in Rows, Columns or Values. */
export function isEmptyLayout(layout: Layout): boolean {
  return layout.rows.length === 0 && layout.columns.length === 0 && layout.values.length === 0;
}

/** What the sheet shows for a filter with no selection. */
export function noRestrictionLabel(field: Pick<PlacedField, "allMember">): string {
  return field.allMember !== null ? "All" : "(Default)";
}

/** The text of a filter row's value cell, exactly as the helper writes it. */
export function filterValueText(field: PlacedField): string {
  if (field.members.length === 0) return noRestrictionLabel(field);
  if (field.members.length === 1) return field.memberCaptions[0] ?? field.members[0];
  return "(Multiple Items)";
}

/** Any of the ways to narrow a field is in use: items, label, date or value filter. */
export function isFiltered(field: PlacedField): boolean {
  return (
    field.members.length > 0 || Boolean(field.labelFilter) || Boolean(field.valueFilter) || Boolean(field.dateFilter)
  );
}

/**
 * How a value field is told apart from the others in Values: its measure for the first
 * plain copy (so older reports keep their names), "<measure>#2" and so on for further
 * copies, and "kpi:<name>:<part>" for a KPI's part.
 */
export function valueKey(values: readonly ValueField[], index: number): string {
  const value = values[index];
  if (value.kpi) return `kpi:${value.kpi.name}:${value.kpi.part}`;
  let earlier = 0;
  for (let i = 0; i < index; i++) {
    if (!values[i].kpi && values[i].measure === value.measure) earlier++;
  }
  return earlier === 0 ? value.measure : `${value.measure}#${earlier + 1}`;
}

export function valueIndex(values: readonly ValueField[], key: string): number {
  for (let i = 0; i < values.length; i++) if (valueKey(values, i) === key) return i;
  return -1;
}

/**
 * The header Excel shows for each value field, as the helper names it: a second plain
 * copy of a measure reads "<caption> 2". Used for the chips in Values.
 */
export function valueLabels(values: readonly ValueField[]): string[] {
  const seen = new Map<string, number>();
  return values.map((value) => {
    if (value.kpi || (value.showAs ?? "normal") !== "normal") return value.caption;
    const n = (seen.get(value.measure) ?? 0) + 1;
    seen.set(value.measure, n);
    return n === 1 ? value.caption : `${value.caption} ${n}`;
  });
}

/** The same selection for a chip: "All", "2025", "3 selected". */
export function selectionSummary(field: PlacedField): string {
  return field.members.length > 1 ? `${field.members.length} selected` : filterValueText(field);
}

export interface Placement {
  area: Area;
  index: number;
}

/** Where a hierarchy or value field (by its `valueKey`) sits, or null when it is not placed. */
export function locate(layout: Layout, name: string): Placement | null {
  for (const area of FIELD_AREAS) {
    const index = layout[area].findIndex((f) => f.hierarchy === name);
    if (index >= 0) return { area, index };
  }
  const index = valueIndex(layout.values, name);
  return index >= 0 ? { area: "values", index } : null;
}

/** The hierarchies a placed field shows: one, or each of a multi-field set's. */
export function fieldHierarchies(field: PlacedField): string[] {
  return field.setHierarchies && field.setHierarchies.length > 0 ? field.setHierarchies : [field.hierarchy];
}

/** The field that shows `hierarchy`, also as the second or later field of a named set. */
export function fieldUsing(layout: Layout, hierarchy: string): PlacedField | null {
  for (const area of FIELD_AREAS) {
    const found = layout[area].find((f) => fieldHierarchies(f).indexOf(hierarchy) >= 0);
    if (found) return found;
  }
  return null;
}

export function findField(layout: Layout, hierarchy: string): PlacedField | null {
  for (const area of FIELD_AREAS) {
    const found = layout[area].find((f) => f.hierarchy === hierarchy);
    if (found) return found;
  }
  return null;
}

function clampIndex(index: number | undefined, length: number): number {
  if (index === undefined || index > length) return length;
  return index < 0 ? 0 : Math.floor(index);
}

function insertAt<T>(list: T[], item: T, index: number | undefined): T[] {
  const at = clampIndex(index, list.length);
  return [...list.slice(0, at), item, ...list.slice(at)];
}

function removeAt<T>(list: T[], index: number): T[] {
  return [...list.slice(0, index), ...list.slice(index + 1)];
}

// -- slicers ------------------------------------------------------------------
// A slicer filters like a page filter and has no cell: what matters is its field's
// chosen items. A hierarchy is in one place only: Rows, Columns, Filters or a slicer.

/** Where a hierarchy is among the report's slicers, or -1. */
export function slicerIndex(layout: Layout, hierarchy: string): number {
  return (layout.slicers ?? []).findIndex((field) => field.hierarchy === hierarchy);
}

/** A field as a slicer keeps it: which field, and which of its items are chosen. */
export function slicerField(field: PlacedField): PlacedField {
  return {
    hierarchy: field.hierarchy,
    caption: field.caption,
    allMember: field.allMember,
    expanded: [],
    members: field.members.slice(),
    memberCaptions: field.memberCaptions.slice(),
  };
}

/** The layout with these slicers; without the key when there are none, as before slicers existed. */
function withSlicers(layout: Layout, slicers: PlacedField[]): Layout {
  const next = { ...layout };
  if (slicers.length > 0) next.slicers = slicers;
  else delete next.slicers;
  return next;
}

function setSlicers(layout: Layout, wanted: readonly HierarchyInput[]): Layout {
  const current = layout.slicers ?? [];
  let out = layout;
  const slicers: PlacedField[] = [];
  for (const input of wanted) {
    if (slicers.length >= MAX_SLICERS) break;
    if (slicers.some((s) => s.hierarchy === input.uniqueName)) continue;
    const existing = current.find((s) => s.hierarchy === input.uniqueName);
    if (existing) {
      slicers.push(existing);
      continue;
    }
    const place = locate(out, input.uniqueName);
    if (place && place.area !== "values") {
      const field = out[place.area][place.index];
      if (field.set) continue; // a named set's items are its own: it cannot be a slicer
      out = { ...out, [place.area]: removeAt(out[place.area], place.index) };
      slicers.push(slicerField(field));
    } else {
      slicers.push({
        hierarchy: input.uniqueName,
        caption: input.caption,
        allMember: input.allMember,
        expanded: [],
        members: [],
        memberCaptions: [],
      });
    }
  }
  const same = out === layout && slicers.length === current.length && slicers.every((s, i) => s === current[i]);
  return same ? layout : withSlicers(out, slicers);
}

const slicerInputs = (layout: Layout): HierarchyInput[] =>
  (layout.slicers ?? []).map((s) => ({ uniqueName: s.hierarchy, caption: s.caption, allMember: s.allMember }));

/** Take a slicer's field out of the slicers, to be placed in an area. */
function takeSlicer(layout: Layout, hierarchy: string): { layout: Layout; field: PlacedField } | null {
  const index = slicerIndex(layout, hierarchy);
  if (index < 0) return null;
  const slicers = layout.slicers as PlacedField[];
  return { layout: withSlicers(layout, removeAt(slicers, index)), field: slicers[index] };
}

function updateField(
  layout: Layout,
  hierarchy: string,
  change: (field: PlacedField) => PlacedField,
): Layout {
  const place = locate(layout, hierarchy);
  if (!place || place.area === "values") {
    // A slicer's field: its chosen items change like a page filter's.
    const index = slicerIndex(layout, hierarchy);
    if (index < 0) return layout;
    const slicers = layout.slicers as PlacedField[];
    const after = change(slicers[index]);
    if (after === slicers[index]) return layout;
    const next = slicers.slice();
    next[index] = after;
    return { ...layout, slicers: next };
  }
  const list = layout[place.area];
  const before = list[place.index];
  const after = change(before);
  if (after === before) return layout;
  const next = list.slice();
  next[place.index] = after;
  return { ...layout, [place.area]: next };
}

function sameJson(a: unknown, b: unknown): boolean {
  return JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
}

/**
 * A copy of `item` with `key` set to `value`, or without `key` when `value` means
 * "the default" (null). Optional settings are left out rather than stored as null, so
 * a layout stays as small as before for fields that do not use them.
 */
function withOptional<T extends object, K extends keyof T>(item: T, key: K, value: T[K] | null): T {
  const next = { ...item };
  if (value === null || value === undefined) delete next[key];
  else next[key] = value;
  return next;
}

function updateValue(layout: Layout, key: string, change: (field: ValueField) => ValueField): Layout {
  const index = valueIndex(layout.values, key);
  if (index < 0) return layout;
  const before = layout.values[index];
  const after = change(before);
  if (after === before) return layout;
  const values = layout.values.slice();
  values[index] = after;
  return { ...layout, values };
}

function moveTo(layout: Layout, name: string, area: Area, index: number | undefined): Layout {
  const place = locate(layout, name);
  if (!place) {
    // From a slicer into an area, with its chosen items.
    const taken = area === "values" ? null : takeSlicer(layout, name);
    return taken ? { ...taken.layout, [area]: insertAt(taken.layout[area as FieldArea], taken.field, index) } : layout;
  }
  // A measure lives only in Values; a hierarchy never does.
  if ((place.area === "values") !== (area === "values")) return layout;

  if (place.area === "values") {
    const item: ValueField = layout.values[place.index];
    const without = removeAt(layout.values, place.index);
    const at = clampIndex(index, without.length);
    if (at === place.index) return layout;
    return { ...layout, values: insertAt(without, item, at) };
  }

  const from = place.area;
  const to = area as FieldArea;
  const item: PlacedField = layout[from][place.index];
  if (item.set && to === "filters") return layout; // a named set cannot be a page filter
  const without = removeAt(layout[from], place.index);
  if (from === to) {
    const at = clampIndex(index, without.length);
    if (at === place.index) return layout;
    return { ...layout, [from]: insertAt(without, item, at) };
  }
  return { ...layout, [from]: without, [to]: insertAt(layout[to], item, index) };
}

/**
 * Value fields whose Show Values As needs a base field that is no longer in Rows or
 * Columns go back to No Calculation, as in Excel. `reset` lists the ones changed so the
 * pane can say so.
 */
export function resetOrphanedShowAs(layout: Layout): { layout: Layout; reset: ValueField[] } {
  const axis = new Set([...layout.rows, ...layout.columns].map((f) => f.hierarchy));
  const reset: ValueField[] = [];
  const values = layout.values.map((value) => {
    if (!needsBaseField(value.showAs) || (value.baseField && axis.has(value.baseField))) return value;
    reset.push(value);
    let next = withOptional(value, "showAs", null);
    next = withOptional(next, "baseField", null);
    return withOptional(next, "baseItem", null);
  });
  return reset.length === 0 ? { layout, reset } : { layout: { ...layout, values }, reset };
}

export function layoutReducer(layout: Layout, action: LayoutAction): Layout {
  const next = applyAction(layout, action);
  return next === layout ? layout : resetOrphanedShowAs(next).layout;
}

function applyAction(layout: Layout, action: LayoutAction): Layout {
  switch (action.type) {
    case "addMeasure": {
      if (valueIndex(layout.values, action.measure.uniqueName) >= 0) return layout;
      const value: ValueField = {
        measure: action.measure.uniqueName,
        caption: action.measure.caption,
      };
      return { ...layout, values: insertAt(layout.values, value, action.index) };
    }

    case "addKpi": {
      const key = `kpi:${action.kpi.name}:${action.part}`;
      if (valueIndex(layout.values, key) >= 0) return layout;
      const value: ValueField = {
        measure: action.kpi.measure,
        caption: `${action.kpi.caption} ${KPI_PART_LABELS[action.part]}`,
        kpi: { name: action.kpi.name, part: action.part },
      };
      return { ...layout, values: insertAt(layout.values, value, action.index) };
    }

    case "duplicateValue": {
      const index = valueIndex(layout.values, action.key);
      if (index < 0 || layout.values[index].kpi) return layout;
      const { measure, caption } = layout.values[index];
      return { ...layout, values: insertAt(layout.values, { measure, caption }, index + 1) };
    }

    case "removeMeasure": {
      const values = layout.values.filter((v) => v.kpi || v.measure !== action.measure);
      return values.length === layout.values.length ? layout : { ...layout, values };
    }

    case "addSet": {
      const spans = action.set.hierarchies && action.set.hierarchies.length > 1 ? action.set.hierarchies : null;
      if ((spans ?? [action.set.hierarchy]).some((h) => fieldUsing(layout, h))) return layout;
      // The set takes its fields' places: a slicer of any of them goes.
      let base = layout;
      for (const h of spans ?? [action.set.hierarchy]) base = takeSlicer(base, h)?.layout ?? base;
      const area = action.area ?? "rows";
      const field: PlacedField = {
        hierarchy: action.set.hierarchy,
        caption: action.set.caption,
        allMember: action.set.allMember,
        expanded: [],
        members: [],
        memberCaptions: [],
        set: action.set.uniqueName,
      };
      if (spans) {
        field.setHierarchies = spans;
        field.setAllMembers = spans.map((_, i) => (action.set.allMembers ?? [])[i] ?? null);
      }
      return { ...base, [area]: insertAt(base[area], field, action.index) };
    }

    case "setDateFilter":
      return updateField(layout, action.hierarchy, (field) => {
        if (sameJson(field.dateFilter, action.filter)) return field;
        const next = withOptional(field, "dateFilter", action.filter);
        return action.filter ? withOptional(next, "labelFilter", null) : next;
      });

    case "setCustomItems":
      return updateField(layout, action.hierarchy, (field) =>
        sameJson(field.customItems ?? [], action.items)
          ? field
          : withOptional(field, "customItems", action.items.length > 0 ? action.items : null),
      );

    case "setCalculatedMeasure":
      return setCalculatedMeasure(layout, action.measure, action.previousName);

    case "removeCalculatedMeasure":
      return removeCalculatedMeasure(layout, action.name);

    case "addHierarchy": {
      const name = action.hierarchy.uniqueName;
      const level = action.hierarchy.level ?? null;
      const placed = findField(layout, name);
      if (placed && !placed.set && (placed.level?.uniqueName ?? null) !== (level?.uniqueName ?? null)) {
        // Another level of a field already in the report: show that level instead, in
        // the same place. What was drilled or picked belonged to the other level.
        return updateField(layout, name, (field) => {
          const next = withOptional(
            { ...field, caption: level ? level.caption : action.hierarchy.caption, expanded: [], members: [], memberCaptions: [] },
            "level",
            level,
          );
          return next;
        });
      }
      if (locate(layout, name)) {
        return action.area ? moveTo(layout, name, action.area, action.index) : layout;
      }
      if (fieldUsing(layout, name)) return layout; // a named set already shows it
      // Ticked or dropped while it is a slicer: it moves, keeping its chosen items.
      if (slicerIndex(layout, name) >= 0) return moveTo(layout, name, action.area ?? "rows", action.index);
      const area: FieldArea = action.area ?? "rows";
      const field: PlacedField = {
        hierarchy: name,
        caption: level ? level.caption : action.hierarchy.caption,
        allMember: action.hierarchy.allMember,
        expanded: [],
        members: [],
        memberCaptions: [],
      };
      if (level) field.level = level;
      return { ...layout, [area]: insertAt(layout[area], field, action.index) };
    }

    case "remove": {
      const place = locate(layout, action.name);
      if (!place) {
        const taken = takeSlicer(layout, action.name);
        return taken ? taken.layout : layout;
      }
      if (place.area === "values") {
        return { ...layout, values: removeAt(layout.values, place.index) };
      }
      return { ...layout, [place.area]: removeAt(layout[place.area], place.index) };
    }

    case "move":
      return moveTo(layout, action.name, action.area, action.index);

    case "shift": {
      const place = locate(layout, action.name);
      if (!place) return layout;
      const length = layout[place.area].length;
      const target = place.index + action.delta;
      if (target < 0 || target >= length) return layout;
      return moveTo(layout, action.name, place.area, target);
    }

    case "setMembers":
      return updateField(layout, action.hierarchy, (field) => {
        const members = action.members.slice();
        // Captions line up with members; a missing caption falls back to the name.
        const memberCaptions = members.map((m, i) => action.captions[i] ?? m);
        const same =
          members.length === field.members.length &&
          members.every((m, i) => m === field.members[i]) &&
          memberCaptions.every((c, i) => c === field.memberCaptions[i]);
        return same ? field : { ...field, members, memberCaptions };
      });

    case "expand":
      return updateField(layout, action.hierarchy, (field) =>
        field.expanded.includes(action.member)
          ? field
          : { ...field, expanded: [...field.expanded, action.member] },
      );

    case "collapse":
      return updateField(layout, action.hierarchy, (field) => {
        const drop = new Set<string>([action.member, ...(action.descendants ?? [])]);
        const expanded = field.expanded.filter((m) => !drop.has(m));
        return expanded.length === field.expanded.length ? field : { ...field, expanded };
      });

    case "setOption":
      if ((layout.options[action.option] ?? false) === action.value) return layout;
      return { ...layout, options: { ...layout.options, [action.option]: action.value } };

    case "setGrandTotals": {
      const now = grandTotalsShown(layout.options);
      const o = layout.options;
      if (now.row === action.row && now.column === action.column && o.grandTotalRow !== undefined) return layout;
      // `grandTotals` stays true when either is on, for helpers that only know it.
      return {
        ...layout,
        options: { ...o, grandTotals: action.row || action.column, grandTotalRow: action.row, grandTotalColumn: action.column },
      };
    }

    case "setEmptyCellText": {
      const text = action.text.slice(0, MAX_EMPTY_CELL_TEXT);
      if ((layout.options.emptyCellText ?? "") === text) return layout;
      const options = { ...layout.options };
      if (text === "") delete options.emptyCellText;
      else options.emptyCellText = text;
      return { ...layout, options };
    }

    case "setForm":
      if (REPORT_FORMS.indexOf(action.form) < 0 || reportForm(layout.options) === action.form) return layout;
      return { ...layout, options: { ...layout.options, form: action.form } };

    case "setSubtotals": {
      const o = layout.options;
      if (action.at === "none") return o.subtotals ? { ...layout, options: { ...o, subtotals: false } } : layout;
      if (action.at !== "top" && action.at !== "bottom") return layout;
      if (o.subtotals && (o.subtotalsAt ?? "bottom") === action.at) return layout;
      return { ...layout, options: { ...o, subtotals: true, subtotalsAt: action.at } };
    }

    case "expandField":
      return updateField(layout, action.hierarchy, (field) => {
        const fresh = action.members.filter((m, i) => field.expanded.indexOf(m) < 0 && action.members.indexOf(m) === i);
        return fresh.length === 0 ? field : { ...field, expanded: [...field.expanded, ...fresh] };
      });

    case "collapseField":
      return updateField(layout, action.hierarchy, (field) =>
        field.expanded.length === 0 ? field : { ...field, expanded: [] },
      );

    case "setSort":
      return updateField(layout, action.hierarchy, (field) =>
        sameJson(field.sort, action.sort) ? field : withOptional(field, "sort", action.sort),
      );

    case "setLabelFilter":
      return updateField(layout, action.hierarchy, (field) => {
        if (sameJson(field.labelFilter, action.filter)) return field;
        const next = withOptional(field, "labelFilter", action.filter);
        return action.filter ? withOptional(next, "dateFilter", null) : next;
      });

    case "setValueFilter":
      return updateField(layout, action.hierarchy, (field) =>
        sameJson(field.valueFilter, action.filter)
          ? field
          : withOptional(field, "valueFilter", action.filter),
      );

    case "clearFilters":
      return updateField(layout, action.hierarchy, (field) => {
        if (!isFiltered(field)) return field;
        let cleared = withOptional(withOptional(field, "labelFilter", null), "valueFilter", null);
        cleared = withOptional(cleared, "dateFilter", null);
        return { ...cleared, members: [], memberCaptions: [] };
      });

    case "setShowAs":
      return updateValue(layout, action.measure, (field) => {
        const baseField = needsBaseField(action.showAs) ? action.baseField ?? null : null;
        const baseItem = needsBaseItem(action.showAs) ? action.baseItem ?? null : null;
        if (
          (field.showAs ?? "normal") === action.showAs &&
          (field.baseField ?? null) === baseField &&
          (field.baseItem ?? null) === baseItem
        ) {
          return field;
        }
        let next = withOptional(field, "showAs", action.showAs === "normal" ? null : action.showAs);
        next = withOptional(next, "baseField", baseField);
        return withOptional(next, "baseItem", baseItem);
      });

    case "setNumberFormat":
      return updateValue(layout, action.measure, (field) => {
        const format = action.format === "" ? null : action.format;
        if ((field.numberFormat ?? null) === format) return field;
        return withOptional(field, "numberFormat", format);
      });

    case "setCustomName":
      return updateValue(layout, action.measure, (field) => {
        const name = (action.name ?? "").trim().slice(0, MAX_CUSTOM_NAME);
        if ((field.customName ?? "").trim() === name) return field;
        return withOptional(field, "customName", name === "" ? null : name);
      });

    // A report keeps a style or option only while it differs from the default, so one
    // that never chose any stays as it was.
    case "setStyle": {
      const style = action.style === null || action.style === DEFAULT_STYLE ? null : action.style;
      if (style !== null && !isStyleName(style)) return layout;
      if ((layout.options.style ?? null) === style) return layout;
      return { ...layout, options: withOptional(layout.options, "style", style) };
    }

    case "setStyleOption": {
      const fallback = STYLE_OPTION_DEFAULTS[action.option];
      if (fallback === undefined || (layout.options[action.option] ?? fallback) === action.value) return layout;
      const value = action.value === fallback ? null : action.value;
      return { ...layout, options: withOptional(layout.options, action.option, value) };
    }

    case "addSlicer":
      return setSlicers(layout, [...slicerInputs(layout), action.hierarchy]);

    case "removeSlicer": {
      const taken = takeSlicer(layout, action.hierarchy);
      return taken ? taken.layout : layout;
    }

    case "setSlicers":
      return setSlicers(layout, action.hierarchies);

    case "replace":
      return action.layout;
  }
}

/** The unique name of a calculated measure: "[Measures].[name]". */
export function calculatedMeasureName(name: string): string {
  return `[Measures].[${name.replace(/\]/g, "]]")}]`;
}

/** Every field's sort and value filter, with `rename` applied to the measures they use. */
function mapMeasureUses(layout: Layout, rename: (measure: string) => string | null): Layout {
  const fix = (field: PlacedField): PlacedField => {
    let next = field;
    if (field.sort && field.sort.by === "measure" && field.sort.measure) {
      const to = rename(field.sort.measure);
      if (to !== field.sort.measure) next = withOptional(next, "sort", to === null ? null : { ...field.sort, measure: to });
    }
    if (field.valueFilter) {
      const to = rename(field.valueFilter.measure);
      if (to !== field.valueFilter.measure) {
        next = withOptional(next, "valueFilter", to === null ? null : { ...field.valueFilter, measure: to });
      }
    }
    return next;
  };
  return { ...layout, rows: layout.rows.map(fix), columns: layout.columns.map(fix), filters: layout.filters.map(fix) };
}

function setCalculatedMeasure(layout: Layout, measure: CalculatedMeasure, previousName?: string): Layout {
  const list = layout.calculatedMeasures ?? [];
  const at = list.findIndex((m) => m.name === (previousName ?? measure.name));
  if (at >= 0 && sameJson(list[at], measure)) return layout;
  const next = list.slice();
  if (at >= 0) next[at] = measure;
  else next.push(measure);
  let out: Layout = { ...layout, calculatedMeasures: next };
  if (previousName !== undefined && previousName !== measure.name) {
    const from = calculatedMeasureName(previousName);
    const to = calculatedMeasureName(measure.name);
    out = {
      ...mapMeasureUses(out, (m) => (m === from ? to : m)),
      values: out.values.map((v) => (v.measure === from ? { ...v, measure: to, caption: measure.name } : v)),
    };
  }
  return out;
}

function removeCalculatedMeasure(layout: Layout, name: string): Layout {
  const list = layout.calculatedMeasures ?? [];
  if (!list.some((m) => m.name === name)) return layout;
  const gone = calculatedMeasureName(name);
  const rest = list.filter((m) => m.name !== name);
  const out: Layout = {
    ...mapMeasureUses(layout, (m) => (m === gone ? null : m)),
    values: layout.values.filter((v) => v.measure !== gone),
  };
  if (rest.length > 0) out.calculatedMeasures = rest;
  else delete out.calculatedMeasures;
  return out;
}
