// The field list: the report's calculated measures, the cube's measures grouped by
// measure group then display folder, its KPIs (Value, Goal, Status, Trend each), and
// dimensions with their hierarchies and named sets. A ticked measure or KPI part is in
// Values; a ticked hierarchy or set is in Rows, Columns or Filters. Measures and
// hierarchies can also be dragged onto an area.

import { useEffect, useMemo, useState } from "react";
import {
  AREA_LABELS,
  KPI_PART_LABELS,
  calculatedMeasureName,
  fieldUsing,
  locate,
  valueIndex,
} from "../state/layout";
import type { UiError } from "../state/runner";
import type { CalculatedMeasure, Fields, Hierarchy, Kpi, KpiPart, Layout, Level, Measure, NamedSet } from "../types";
import { cx } from "../util";
import { ErrorNote, Icon, Spinner } from "./common";
import { endDrag, startDrag } from "./drag";

interface Branch {
  type: "branch";
  id: string;
  label: string;
  icon: "measure" | "table" | "folder" | "kpi" | "set" | "calc";
  children: Entry[];
}
interface MeasureLeaf {
  type: "measure";
  id: string;
  measure: Measure;
}
interface HierarchyLeaf {
  type: "hierarchy";
  id: string;
  hierarchy: Hierarchy;
}
interface KpiLeaf {
  type: "kpi";
  id: string;
  kpi: Kpi;
  part: KpiPart;
}
interface SetLeaf {
  type: "set";
  id: string;
  set: NamedSet;
}
interface CalculatedLeaf {
  type: "calculated";
  id: string;
  measure: CalculatedMeasure;
}
type Leaf = MeasureLeaf | HierarchyLeaf | KpiLeaf | SetLeaf | CalculatedLeaf;
type Entry = Branch | Leaf;

/** Display folders nest with "\"; a field listed in several (";") is shown in the first. */
function folderPath(folder: string): string[] {
  const first = folder.split(";")[0];
  return first
    .split("\\")
    .map((part) => part.trim())
    .filter((part) => part !== "");
}

function branchFor(parent: Branch, path: string[]): Branch {
  let at = parent;
  for (const part of path) {
    const id = `${at.id}/${part}`;
    let next = at.children.find((c): c is Branch => c.type === "branch" && c.id === id);
    if (!next) {
      next = { type: "branch", id, label: part, icon: "folder", children: [] };
      at.children.push(next);
    }
    at = next;
  }
  return at;
}

/** Fields directly in a branch come before its folders, as in Excel's own field list. */
function sortFoldersLast(branch: Branch): void {
  const leaves = branch.children.filter((c) => c.type !== "branch");
  const folders = branch.children.filter((c): c is Branch => c.type === "branch");
  folders.forEach(sortFoldersLast);
  branch.children = [...leaves, ...folders];
}

export function buildFieldTree(fields: Fields, calculated: readonly CalculatedMeasure[] = []): Branch[] {
  const roots: Branch[] = [];
  // Always listed, so that "New Calculated Measure…" has a home.
  roots.push({
    type: "branch",
    id: "c:",
    label: "Calculated Measures",
    icon: "calc",
    children: calculated.map((measure) => ({ type: "calculated", id: `c:${measure.name}`, measure })),
  });
  const groups = new Map<string, Branch>();
  for (const measure of fields.measures) {
    let group = groups.get(measure.group);
    if (!group) {
      group = {
        type: "branch",
        id: `m:${measure.group}`,
        label: measure.group || "Measures",
        icon: "measure",
        children: [],
      };
      groups.set(measure.group, group);
      roots.push(group);
    }
    branchFor(group, folderPath(measure.folder)).children.push({
      type: "measure",
      id: measure.uniqueName,
      measure,
    });
  }
  const kpis = fields.kpis ?? [];
  if (kpis.length > 0) {
    const branch: Branch = { type: "branch", id: "k:", label: "KPIs", icon: "kpi", children: [] };
    for (const kpi of kpis) {
      const own: Branch = { type: "branch", id: `k:${kpi.name}`, label: kpi.caption, icon: "kpi", children: [] };
      for (const part of kpi.parts) own.children.push({ type: "kpi", id: `kpi:${kpi.name}:${part}`, kpi, part });
      branchFor(branch, folderPath(kpi.folder)).children.push(own);
    }
    roots.push(branch);
  }
  const sets = fields.sets ?? [];
  const setsOf = (dimension: string) => sets.filter((s) => s.dimension === dimension);
  const setBranch = (id: string, list: NamedSet[]): Branch => {
    const branch: Branch = { type: "branch", id, label: "Sets", icon: "set", children: [] };
    for (const set of list) {
      branchFor(branch, folderPath(set.folder)).children.push({ type: "set", id: `s:${set.uniqueName}`, set });
    }
    return branch;
  };
  for (const dimension of fields.dimensions) {
    if (dimension.hierarchies.length === 0) continue;
    const branch: Branch = {
      type: "branch",
      id: `d:${dimension.uniqueName}`,
      label: dimension.caption,
      icon: "table",
      children: [],
    };
    for (const hierarchy of dimension.hierarchies) {
      branchFor(branch, folderPath(hierarchy.folder)).children.push({
        type: "hierarchy",
        id: hierarchy.uniqueName,
        hierarchy,
      });
    }
    const own = setsOf(dimension.uniqueName);
    if (own.length > 0) branch.children.push(setBranch(`ds:${dimension.uniqueName}`, own));
    roots.push(branch);
  }
  const loose = sets.filter((s) => !fields.dimensions.some((d) => d.uniqueName === s.dimension && d.hierarchies.length > 0));
  if (loose.length > 0) roots.push(setBranch("s:", loose));
  roots.forEach(sortFoldersLast);
  return roots;
}

function captionOf(entry: Leaf): string {
  switch (entry.type) {
    case "measure":
      return entry.measure.caption;
    case "hierarchy":
      return entry.hierarchy.caption;
    case "kpi":
      return `${entry.kpi.caption} ${KPI_PART_LABELS[entry.part]}`;
    case "set":
      return entry.set.caption;
    case "calculated":
      return entry.measure.name;
  }
}

/** Whether a leaf is in the report, where, and whether it can be ticked now. */
function leafState(entry: Leaf, layout: Layout): { placed: string | null; checked: boolean; disabled: string | null } {
  switch (entry.type) {
    case "measure":
    case "calculated": {
      const name = entry.type === "measure" ? entry.measure.uniqueName : calculatedMeasureName(entry.measure.name);
      const on = valueIndex(layout.values, name) >= 0;
      return { placed: null, checked: on, disabled: null };
    }
    case "kpi":
      return { placed: null, checked: valueIndex(layout.values, entry.id) >= 0, disabled: null };
    case "hierarchy": {
      const field = fieldUsing(layout, entry.hierarchy.uniqueName);
      if (field && field.set) {
        return { placed: null, checked: false, disabled: `The set \u201c${field.caption}\u201d is using this field.` };
      }
      const place = locate(layout, entry.hierarchy.uniqueName);
      // With one of its levels placed, the whole hierarchy is not: ticking it shows it all.
      const whole = place !== null && !(field && field.level);
      return { placed: whole && place ? AREA_LABELS[place.area] : null, checked: whole, disabled: null };
    }
    case "set": {
      if (!entry.set.hierarchy) {
        return {
          placed: null,
          checked: false,
          disabled: "This set\u2019s fields are not all in this cube\u2019s field list.",
        };
      }
      const spans = entry.set.hierarchies && entry.set.hierarchies.length > 0 ? entry.set.hierarchies : [entry.set.hierarchy];
      const mine = fieldUsing(layout, entry.set.hierarchy);
      if (mine && mine.set === entry.set.uniqueName) {
        const place = locate(layout, entry.set.hierarchy);
        return { placed: place ? AREA_LABELS[place.area] : null, checked: true, disabled: null };
      }
      const taken = spans.map((h) => fieldUsing(layout, h)).find((f) => f !== null);
      if (taken) {
        return { placed: null, checked: false, disabled: `\u201c${taken.caption}\u201d already uses one of this set\u2019s fields.` };
      }
      return { placed: null, checked: false, disabled: null };
    }
  }
}

const LEAF_ICONS: Record<Leaf["type"], "attribute" | "hierarchy" | "kpi" | "set" | "calc" | null> = {
  measure: null,
  hierarchy: "attribute",
  kpi: null,
  set: "set",
  calculated: null,
};

/** Keep the leaves whose caption contains `needle`, and the branches above them. */
export function filterFieldTree(entries: Entry[], needle: string): Entry[] {
  if (needle === "") return entries;
  const out: Entry[] = [];
  for (const entry of entries) {
    if (entry.type === "branch") {
      const children = filterFieldTree(entry.children, needle);
      if (children.length > 0) out.push({ ...entry, children });
    } else if (captionOf(entry).toLowerCase().indexOf(needle) >= 0) {
      out.push(entry);
    }
  }
  return out;
}

function countLeaves(entries: Entry[]): number {
  let n = 0;
  for (const entry of entries) n += entry.type === "branch" ? countLeaves(entry.children) : 1;
  return n;
}

/** With a big cube, start with the dimensions folded so the list is scannable. */
const AUTO_COLLAPSE_ABOVE = 40;

interface Props {
  fields: Fields | null;
  loading: boolean;
  error: UiError | null;
  onRetry: () => void;
  /** Shown when there is no cube to list yet. */
  placeholder: string;
  layout: Layout;
  search: string;
  onSearch: (text: string) => void;
  onToggleMeasure: (measure: Measure, on: boolean) => void;
  onToggleHierarchy: (hierarchy: Hierarchy, on: boolean) => void;
  onToggleKpi: (kpi: Kpi, part: KpiPart, on: boolean) => void;
  onToggleSet: (set: NamedSet, on: boolean) => void;
  onToggleCalculated: (measure: CalculatedMeasure, on: boolean) => void;
  /** Place one level of a user hierarchy (Windows: tick "Month" under "Calendar"). */
  onToggleLevel: (hierarchy: Hierarchy, level: Level, on: boolean) => void;
  /** Open the calculated measure dialog: a new one (null), or this one to edit. */
  onEditCalculated: (measure: CalculatedMeasure | null) => void;
}

export function FieldList(props: Props) {
  const { fields, layout, search } = props;
  const calculated = layout.calculatedMeasures;
  const tree = useMemo(() => (fields ? buildFieldTree(fields, calculated ?? []) : []), [fields, calculated]);
  const [collapsed, setCollapsed] = useState<ReadonlySet<string>>(new Set());
  /** User hierarchies whose levels are listed under them; folded by default, as in Excel. */
  const [openLevels, setOpenLevels] = useState<ReadonlySet<string>>(new Set());

  useEffect(() => {
    const total = countLeaves(tree);
    setCollapsed(
      new Set(total > AUTO_COLLAPSE_ABOVE ? tree.filter((b) => b.icon === "table").map((b) => b.id) : []),
    );
    // Only when the cube changes: a new calculated measure must not fold the list.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fields]);

  const needle = search.trim().toLowerCase();
  const visible = useMemo(() => filterFieldTree(tree, needle), [tree, needle]);

  const toggleBranch = (id: string) =>
    setCollapsed((before) => {
      const next = new Set(before);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });

  const render = (entries: Entry[], depth: number): JSX.Element[] =>
    entries.map((entry) => {
      const indent = { paddingLeft: 6 + depth * 14 };
      if (entry.type === "branch") {
        // While searching, everything that matches is shown unfolded.
        const open = needle !== "" || !collapsed.has(entry.id);
        return (
          <li key={entry.id} role="treeitem" aria-expanded={open}>
            <button
              type="button"
              className="field-row branch"
              style={indent}
              onClick={() => toggleBranch(entry.id)}
              disabled={needle !== ""}
            >
              <Icon name={open ? "chevron-down" : "chevron-right"} className="twist" />
              <Icon name={entry.icon} className="kind" />
              <span className="field-caption">{entry.label}</span>
            </button>
            {open && (
              <ul role="group">
                {render(entry.children, depth + 1)}
                {entry.id === "c:" && needle === "" && (
                  <li role="treeitem">
                    <button
                      type="button"
                      className="field-row leaf action-row"
                      style={{ paddingLeft: 6 + (depth + 1) * 14 }}
                      onClick={() => props.onEditCalculated(null)}
                    >
                      <Icon name="plus" className="kind" />
                      <span className="field-caption">New Calculated Measure…</span>
                    </button>
                  </li>
                )}
              </ul>
            )}
          </li>
        );
      }
      const caption = captionOf(entry);
      const state = leafState(entry, layout);
      // Measures, calculated measures and hierarchies can be dragged onto an area.
      const drag =
        entry.type === "measure" || entry.type === "calculated"
          ? {
              kind: "measure" as const,
              name: entry.type === "measure" ? entry.measure.uniqueName : calculatedMeasureName(entry.measure.name),
              caption,
              allMember: null,
            }
          : entry.type === "hierarchy"
            ? { kind: "hierarchy" as const, name: entry.hierarchy.uniqueName, caption, allMember: entry.hierarchy.allMember }
            : null;
      const toggle = (on: boolean) => {
        if (entry.type === "measure") props.onToggleMeasure(entry.measure, on);
        else if (entry.type === "hierarchy") props.onToggleHierarchy(entry.hierarchy, on);
        else if (entry.type === "kpi") props.onToggleKpi(entry.kpi, entry.part, on);
        else if (entry.type === "set") props.onToggleSet(entry.set, on);
        else props.onToggleCalculated(entry.measure, on);
      };
      const icon =
        entry.type === "hierarchy" && entry.hierarchy.kind === "user" ? "hierarchy" : LEAF_ICONS[entry.type];
      const levels = entry.type === "hierarchy" ? entry.hierarchy.levels ?? [] : [];
      const levelsOpen = levels.length > 1 && openLevels.has(entry.id);
      const levelRows =
        entry.type === "hierarchy" && levelsOpen
          ? levels.map((level) => {
              const field = fieldUsing(layout, entry.hierarchy.uniqueName);
              const on = Boolean(field && !field.set && field.level && field.level.uniqueName === level.uniqueName);
              const place = on ? locate(layout, entry.hierarchy.uniqueName) : null;
              return (
                <li key={`lv:${level.uniqueName}`} role="treeitem" aria-selected={on}>
                  <label
                    className={cx("field-row leaf", on && "placed", field && field.set && "unavailable")}
                    style={{ paddingLeft: 6 + (depth + 1) * 14 }}
                    title={`${entry.hierarchy.caption}: ${level.caption} only`}
                  >
                    <input
                      type="checkbox"
                      checked={on}
                      disabled={Boolean(field && field.set)}
                      onChange={(event) => props.onToggleLevel(entry.hierarchy, level, event.target.checked)}
                    />
                    <Icon name="attribute" className="kind" />
                    <span className="field-caption">{level.caption}</span>
                    {place && <span className="field-area">{AREA_LABELS[place.area]}</span>}
                  </label>
                </li>
              );
            })
          : [];
      return (
        <li key={entry.id} role="treeitem" aria-selected={state.checked} aria-expanded={levels.length > 1 ? levelsOpen : undefined}>
          <label
            className={cx("field-row leaf", state.checked && "placed", state.disabled !== null && "unavailable")}
            style={indent}
            title={state.disabled ?? caption}
            draggable={drag !== null && state.disabled === null}
            onDragStart={drag ? (event) => startDrag(event, drag) : undefined}
            onDragEnd={drag ? endDrag : undefined}
          >
            <input
              type="checkbox"
              checked={state.checked}
              disabled={state.disabled !== null}
              onChange={(event) => toggle(event.target.checked)}
            />
            {icon && <Icon name={icon} className="kind" />}
            <span className="field-caption">{entry.type === "kpi" ? KPI_PART_LABELS[entry.part] : caption}</span>
            {state.placed !== null && <span className="field-area">{state.placed}</span>}
            {levels.length > 1 && (
              <button
                type="button"
                className="icon-button tiny"
                aria-label={levelsOpen ? `Hide the levels of ${caption}` : `Show the levels of ${caption}`}
                title={levelsOpen ? "Hide levels" : "Show levels, to place one alone"}
                onClick={(event) => {
                  event.preventDefault();
                  setOpenLevels((before) => {
                    const next = new Set(before);
                    if (next.has(entry.id)) next.delete(entry.id);
                    else next.add(entry.id);
                    return next;
                  });
                }}
              >
                <Icon name={levelsOpen ? "chevron-down" : "chevron-right"} />
              </button>
            )}
            {entry.type === "calculated" && (
              <button
                type="button"
                className="icon-button tiny"
                aria-label={`Edit ${caption}`}
                title="Edit or delete this calculated measure"
                onClick={(event) => {
                  event.preventDefault();
                  props.onEditCalculated(entry.measure);
                }}
              >
                <Icon name="more" />
              </button>
            )}
          </label>
          {levelRows.length > 0 && <ul role="group">{levelRows}</ul>}
        </li>
      );
    });

  return (
    <section className="fields" aria-label="Fields">
      <div className="search">
        <Icon name="search" className="search-icon" />
        <input
          className="input search-input"
          type="search"
          value={search}
          onChange={(event) => props.onSearch(event.target.value)}
          placeholder="Search fields"
          aria-label="Search fields"
          autoCapitalize="off"
          autoCorrect="off"
          spellCheck={false}
          disabled={!fields}
        />
      </div>
      <div className="field-box">
        {props.error ? (
          <ErrorNote error={props.error} action={{ label: "Try again", run: props.onRetry }} />
        ) : props.loading ? (
          <p className="empty">
            <Spinner label="Loading fields" /> Loading fields…
          </p>
        ) : !fields ? (
          <p className="empty">{props.placeholder}</p>
        ) : tree.length === 0 ? (
          <p className="empty">This cube has no fields you can use.</p>
        ) : visible.length === 0 ? (
          <p className="empty">No field is called &ldquo;{search.trim()}&rdquo;.</p>
        ) : (
          <ul className="field-tree" role="tree">
            {render(visible, 0)}
          </ul>
        )}
      </div>
    </section>
  );
}
