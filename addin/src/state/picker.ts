// Selection logic of the member picker. Pure.
//
// The contract only knows an include list: `PlacedField.members`, where an empty list
// means "no restriction". The picker shows that as a tree of checkboxes with the usual
// three states. A ticked member stands for itself and everything below it, so the
// selection is kept as the smallest set of members that covers what is ticked. A
// virtual root above the top level stands for "everything". The helper lists the
// hierarchy's top-level members directly under it; the All member is never in the tree.

import type { Member, PlacedField } from "../types";

/** Key of the virtual root. Not a valid unique name, so it cannot collide with one. */
export const ROOT = "(root)";

export interface PickerNode {
  key: string;
  caption: string;
  level: number;
  hasChildren: boolean;
  /** Key of the parent; ROOT for top-level members; null for the root itself. */
  parent: string | null;
  /** Keys of the children, once loaded. */
  children?: string[];
  /** The helper cut the list of children off at its limit. */
  truncated?: boolean;
}

export interface PickerState {
  nodes: ReadonlyMap<string, PickerNode>;
  /** Minimal covering set of ticked keys. Contains ROOT when everything is ticked. */
  selected: ReadonlySet<string>;
  /** Captions of selected members that are not (yet) in the tree. */
  captions: ReadonlyMap<string, string>;
}

export type CheckState = "on" | "off" | "partial";

export function initPicker(field: Pick<PlacedField, "members" | "memberCaptions">): PickerState {
  const nodes = new Map<string, PickerNode>();
  nodes.set(ROOT, { key: ROOT, caption: "", level: -1, hasChildren: true, parent: null });
  const captions = new Map<string, string>();
  field.members.forEach((m, i) => captions.set(m, field.memberCaptions[i] ?? m));
  return {
    nodes,
    selected: new Set(field.members.length > 0 ? field.members : [ROOT]),
    captions,
  };
}

/** Record the children of `parentKey` (ROOT for the top level) as loaded. */
export function setChildren(
  state: PickerState,
  parentKey: string,
  members: Member[],
  truncated: boolean,
): PickerState {
  const parent = state.nodes.get(parentKey);
  if (!parent) return state;
  const nodes = new Map(state.nodes);
  const keys: string[] = [];
  for (const m of members) {
    if (m.uniqueName === ROOT) continue;
    keys.push(m.uniqueName);
    const known = nodes.get(m.uniqueName);
    nodes.set(m.uniqueName, {
      key: m.uniqueName,
      caption: m.caption,
      level: m.level,
      hasChildren: m.hasChildren,
      parent: parentKey,
      children: known ? known.children : undefined,
      truncated: known ? known.truncated : undefined,
    });
  }
  nodes.set(parentKey, { ...parent, children: keys, truncated });
  return { ...state, nodes };
}

function ancestors(state: PickerState, key: string): string[] {
  const out: string[] = [];
  let node = state.nodes.get(key);
  while (node && node.parent !== null) {
    out.push(node.parent);
    node = state.nodes.get(node.parent);
  }
  return out;
}

function loadedDescendants(state: PickerState, key: string, out: string[] = []): string[] {
  const node = state.nodes.get(key);
  if (!node || !node.children) return out;
  for (const child of node.children) {
    out.push(child);
    loadedDescendants(state, child, out);
  }
  return out;
}

/** True when everything is ticked. */
export function isAll(state: PickerState): boolean {
  return state.selected.has(ROOT);
}

export function checkState(state: PickerState, key: string): CheckState {
  if (state.selected.has(key)) return "on";
  for (const a of ancestors(state, key)) if (state.selected.has(a)) return "on";
  if (state.selected.size === 0) return "off";
  for (const d of loadedDescendants(state, key)) if (state.selected.has(d)) return "partial";
  return "off";
}

export interface ToggleResult {
  state: PickerState;
  /**
   * Set when the click could not be honoured: unticking one item means "all the
   * others", and the others are not all known because their list was cut off.
   */
  blocked?: "truncated";
}

/** Tick or untick a member in the tree. */
export function toggle(state: PickerState, key: string): ToggleResult {
  const node = state.nodes.get(key);
  if (!node) return { state };
  const selected = new Set(state.selected);

  if (checkState(state, key) === "on") {
    if (selected.has(key)) {
      selected.delete(key);
      return { state: { ...state, selected } };
    }
    // Ticked through an ancestor: replace that ancestor by everything under it
    // except the path down to this member.
    const up = ancestors(state, key);
    const top = up.findIndex((a) => selected.has(a));
    if (top < 0) return { state };
    const path = [...up.slice(0, top + 1).reverse(), key]; // ancestor ... parent, key
    for (let i = 0; i < path.length - 1; i++) {
      const p = state.nodes.get(path[i]);
      if (!p || !p.children || p.truncated) return { state, blocked: "truncated" };
    }
    selected.delete(path[0]);
    for (let i = 0; i < path.length - 1; i++) {
      const p = state.nodes.get(path[i]) as PickerNode;
      for (const child of p.children as string[]) {
        if (child !== path[i + 1]) selected.add(child);
      }
    }
    return { state: { ...state, selected } };
  }

  // Off or partial: tick it, which covers everything below it.
  for (const d of loadedDescendants(state, key)) selected.delete(d);
  selected.add(key);
  // If that completes a parent, the parent stands for its children.
  let parentKey = node.parent;
  while (parentKey !== null) {
    const parent = state.nodes.get(parentKey);
    if (!parent || !parent.children || parent.truncated) break;
    if (!parent.children.every((child) => selected.has(child))) break;
    for (const child of parent.children) selected.delete(child);
    selected.add(parent.key);
    parentKey = parent.parent;
  }
  return { state: { ...state, selected } };
}

/**
 * Tick or untick a search result. Results are a flat list whose ancestors are not
 * known, so there is no three-state logic here: a result is ticked when it is itself in
 * the selection. Ticking one while everything is selected means "only this".
 */
export function toggleFlat(state: PickerState, member: Member, allMember: string | null): PickerState {
  const key = member.uniqueName;
  const captions = new Map(state.captions);
  captions.set(key, member.caption);
  const everything = isAll(state) || (allMember !== null && state.selected.has(allMember));
  if (everything) return { ...state, selected: new Set([key]), captions };
  const selected = new Set(state.selected);
  if (selected.has(key)) selected.delete(key);
  else selected.add(key);
  return { ...state, selected, captions };
}

export function selectAll(state: PickerState): PickerState {
  return { ...state, selected: new Set([ROOT]) };
}

export function selectNone(state: PickerState): PickerState {
  return { ...state, selected: new Set<string>() };
}

export function captionOf(state: PickerState, key: string): string {
  const node = state.nodes.get(key);
  return node ? node.caption : state.captions.get(key) ?? key;
}

export interface PickerResult {
  members: string[];
  captions: string[];
}

/**
 * The selection as the contract wants it, or null when nothing is ticked (which is not
 * a valid filter). Everything ticked is "no restriction". So is a selection that holds
 * the All member, which the tree never offers but a search across all levels, or a
 * workbook saved by an earlier version, can still bring in.
 */
export function pickerResult(state: PickerState, allMember: string | null): PickerResult | null {
  if (state.selected.size === 0) return null;
  if (isAll(state) || (allMember !== null && state.selected.has(allMember))) {
    return { members: [], captions: [] };
  }
  const members = Array.from(state.selected);
  return { members, captions: members.map((m) => captionOf(state, m)) };
}

/** Short text describing the selection, for the picker footer and the chip tooltip. */
export function describeSelection(state: PickerState, allMember: string | null): string {
  const result = pickerResult(state, allMember);
  if (result === null) return "Nothing selected";
  if (result.members.length === 0) return "All selected";
  if (result.members.length <= 3) return result.captions.join(", ");
  return `${result.members.length} selected`;
}
