// Value field names: Windows' Value Field Settings > Custom Name. The name is the pane's
// own (the helper never sees it): the pane writes it into the report's header cells in
// place of the cube's caption, and shows it on the field's chip.

import type { Layout, ValueField } from "../types";
import { MAX_CUSTOM_NAME, valueIndex, valueLabels } from "./layout";

export { MAX_CUSTOM_NAME };

/** A value field's own name, if it has one. */
export function customNameOf(value: Pick<ValueField, "customName">): string | null {
  const name = typeof value.customName === "string" ? value.customName.trim() : "";
  return name === "" ? null : name;
}

/**
 * What each value field is called in the report, in Values order: its custom name, or
 * the header the helper gives it ("Rides", "Rides 2" for a second copy, the KPI part's
 * caption).
 */
export function valueFieldNames(values: readonly ValueField[]): string[] {
  const labels = valueLabels(values);
  return values.map((value, i) => customNameOf(value) ?? labels[i]);
}

/**
 * Why `name` cannot be the custom name of the value field with key `key`, or null when it
 * can. As Excel does, a name another field of the report already has is refused: another
 * value field's, or that of a field in Rows, Columns, Filters or a slicer. Letters are
 * compared without regard to case. An empty name is fine: it gives back the cube's.
 */
export function customNameProblem(layout: Layout, key: string, name: string): string | null {
  const wanted = name.trim();
  if (wanted === "") return null;
  if (wanted.length > MAX_CUSTOM_NAME) return `A name can be at most ${MAX_CUSTOM_NAME} characters long.`;
  const index = valueIndex(layout.values, key);
  const same = (other: string) => other.trim().toLowerCase() === wanted.toLowerCase();
  const names = valueFieldNames(layout.values);
  const taken =
    names.some((other, i) => i !== index && same(other)) ||
    [...layout.rows, ...layout.columns, ...layout.filters, ...(layout.slicers ?? [])].some((field) => same(field.caption));
  return taken ? `This report already has a field called “${wanted}”. Type another name.` : null;
}
