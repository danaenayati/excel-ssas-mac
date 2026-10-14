// The pivot layout reducer. Pure: no React, no I/O. Every action returns the same
// object when it changes nothing, so callers can tell "no change" by identity and
// skip a re-query.

import type {
  FieldSort,
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
}

export type LayoutAction =
  /** Tick a measure: goes to Values. No-op when already there. */
  | { type: "addMeasure"; measure: MeasureInput; index?: number }
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
  | { type: "setOption"; option: keyof LayoutOptions; value: boolean }
  /** Order of a field's items; null for the cube's own order. */
  | { type: "setSort"; hierarchy: string; sort: FieldSort | null }
  | { type: "setLabelFilter"; hierarchy: string; filter: LabelFilter | null }
  | { type: "setValueFilter"; hierarchy: string; filter: ValueFilter | null }
  /** "Clear Filter From <field>": items, label filter and value filter. Not the sort. */
  | { type: "clearFilters"; hierarchy: string }
  | { type: "setShowAs"; measure: string; showAs: ShowValuesAs }
  /** A value field's number format; null for the cube's own. */
  | { type: "setNumberFormat"; measure: string; format: string | null }
  | { type: "replace"; layout: Layout };

export const DEFAULT_OPTIONS: LayoutOptions = {
  grandTotals: true,
  subtotals: false,
  nonEmpty: true,
};

export function emptyLayout(): Layout {
  return { rows: [], columns: [], values: [], filters: [], options: { ...DEFAULT_OPTIONS } };
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

/** Any of the three ways to narrow a field is in use: items, label filter, value filter. */
export function isFiltered(field: PlacedField): boolean {
  return field.members.length > 0 || Boolean(field.labelFilter) || Boolean(field.valueFilter);
}

/** The same selection for a chip: "All", "2025", "3 selected". */
export function selectionSummary(field: PlacedField): string {
  return field.members.length > 1 ? `${field.members.length} selected` : filterValueText(field);
}

export interface Placement {
  area: Area;
  index: number;
}

/** Where a measure or hierarchy currently sits, or null when it is not placed. */
export function locate(layout: Layout, name: string): Placement | null {
  for (const area of FIELD_AREAS) {
    const index = layout[area].findIndex((f) => f.hierarchy === name);
    if (index >= 0) return { area, index };
  }
  const index = layout.values.findIndex((v) => v.measure === name);
  return index >= 0 ? { area: "values", index } : null;
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

function updateField(
  layout: Layout,
  hierarchy: string,
  change: (field: PlacedField) => PlacedField,
): Layout {
  const place = locate(layout, hierarchy);
  if (!place || place.area === "values") return layout;
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

function updateValue(layout: Layout, measure: string, change: (field: ValueField) => ValueField): Layout {
  const index = layout.values.findIndex((v) => v.measure === measure);
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
  if (!place) return layout;
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
  const without = removeAt(layout[from], place.index);
  if (from === to) {
    const at = clampIndex(index, without.length);
    if (at === place.index) return layout;
    return { ...layout, [from]: insertAt(without, item, at) };
  }
  return { ...layout, [from]: without, [to]: insertAt(layout[to], item, index) };
}

export function layoutReducer(layout: Layout, action: LayoutAction): Layout {
  switch (action.type) {
    case "addMeasure": {
      if (locate(layout, action.measure.uniqueName)) return layout;
      const value: ValueField = {
        measure: action.measure.uniqueName,
        caption: action.measure.caption,
      };
      return { ...layout, values: insertAt(layout.values, value, action.index) };
    }

    case "addHierarchy": {
      const name = action.hierarchy.uniqueName;
      if (locate(layout, name)) {
        return action.area ? moveTo(layout, name, action.area, action.index) : layout;
      }
      const area: FieldArea = action.area ?? "rows";
      const field: PlacedField = {
        hierarchy: name,
        caption: action.hierarchy.caption,
        allMember: action.hierarchy.allMember,
        expanded: [],
        members: [],
        memberCaptions: [],
      };
      return { ...layout, [area]: insertAt(layout[area], field, action.index) };
    }

    case "remove": {
      const place = locate(layout, action.name);
      if (!place) return layout;
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
      if (layout.options[action.option] === action.value) return layout;
      return { ...layout, options: { ...layout.options, [action.option]: action.value } };

    case "setSort":
      return updateField(layout, action.hierarchy, (field) =>
        sameJson(field.sort, action.sort) ? field : withOptional(field, "sort", action.sort),
      );

    case "setLabelFilter":
      return updateField(layout, action.hierarchy, (field) =>
        sameJson(field.labelFilter, action.filter)
          ? field
          : withOptional(field, "labelFilter", action.filter),
      );

    case "setValueFilter":
      return updateField(layout, action.hierarchy, (field) =>
        sameJson(field.valueFilter, action.filter)
          ? field
          : withOptional(field, "valueFilter", action.filter),
      );

    case "clearFilters":
      return updateField(layout, action.hierarchy, (field) => {
        if (!isFiltered(field)) return field;
        const cleared = withOptional(withOptional(field, "labelFilter", null), "valueFilter", null);
        return { ...cleared, members: [], memberCaptions: [] };
      });

    case "setShowAs":
      return updateValue(layout, action.measure, (field) => {
        const current = field.showAs ?? "normal";
        if (current === action.showAs) return field;
        return withOptional(field, "showAs", action.showAs === "normal" ? null : action.showAs);
      });

    case "setNumberFormat":
      return updateValue(layout, action.measure, (field) => {
        const format = action.format === "" ? null : action.format;
        if ((field.numberFormat ?? null) === format) return field;
        return withOptional(field, "numberFormat", format);
      });

    case "replace":
      return action.layout;
  }
}
