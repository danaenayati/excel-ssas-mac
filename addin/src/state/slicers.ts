// What a click on a slicer's button does, as in Excel. Pure.
//
//   - A plain click selects that item only.
//   - With Multi-Select on, or Ctrl/Cmd held, a click adds the item or takes it out.
//     With nothing filtered every item counts as selected, so the first such click
//     takes its item out, as Excel does. The last selected item cannot be taken out.
//   - Selecting every item is no filter at all ("All", nothing stored).
//
// The buttons are the field's top-level items. Items chosen elsewhere that are not among
// them (a month picked while the field was a page filter) are kept by a multi-select
// click and dropped by a plain one.

import type { PlacedField } from "../types";

export interface SlicerItem {
  uniqueName: string;
  caption: string;
}

export type SlicerClick =
  /** The new selection; empty means no filter. */
  | { members: string[]; captions: string[] }
  /** Nothing to change. */
  | { unchanged: true }
  /**
   * Not possible: "truncated" when leaving one item out of everything needs the whole
   * list and only part of it was loaded; "last" when it is the only item left.
   */
  | { blocked: "truncated" | "last" };

/** Whether a button shows as selected: every item when nothing is filtered. */
export function isSlicerSelected(field: Pick<PlacedField, "members">, item: string): boolean {
  return field.members.length === 0 || field.members.indexOf(item) >= 0;
}

/** Chosen items of the field that have no button. */
export function slicerExtras(field: Pick<PlacedField, "members" | "memberCaptions">, items: readonly SlicerItem[]): SlicerItem[] {
  const shown = new Set(items.map((item) => item.uniqueName));
  return field.members
    .map((uniqueName, i) => ({ uniqueName, caption: field.memberCaptions[i] ?? uniqueName }))
    .filter((member) => !shown.has(member.uniqueName));
}

export function slicerClick(
  field: Pick<PlacedField, "members" | "memberCaptions">,
  items: readonly SlicerItem[],
  truncated: boolean,
  item: SlicerItem,
  multi: boolean,
): SlicerClick {
  if (!multi) {
    const only = field.members.length === 1 && field.members[0] === item.uniqueName;
    return only ? { unchanged: true } : { members: [item.uniqueName], captions: [item.caption] };
  }
  const extras = slicerExtras(field, items);
  let chosen: Set<string>;
  if (field.members.length === 0) {
    if (truncated) return { blocked: "truncated" };
    chosen = new Set(items.map((i) => i.uniqueName));
  } else {
    chosen = new Set(field.members);
  }
  if (chosen.has(item.uniqueName)) chosen.delete(item.uniqueName);
  else chosen.add(item.uniqueName);
  if (chosen.size === 0) return { blocked: "last" };

  const listed = items.filter((i) => chosen.has(i.uniqueName));
  const kept = extras.filter((e) => chosen.has(e.uniqueName));
  // Every item selected, and nothing else: no filter.
  if (!truncated && kept.length === 0 && listed.length === items.length) return { members: [], captions: [] };
  const all = [...listed, ...kept];
  return { members: all.map((i) => i.uniqueName), captions: all.map((i) => i.caption) };
}
