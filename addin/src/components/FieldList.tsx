// The field list: measures grouped by measure group then display folder, and
// dimensions with their hierarchies. A ticked measure is in Values; a ticked hierarchy
// is in Rows, Columns or Filters. Rows can also be dragged onto an area.

import { useEffect, useMemo, useState } from "react";
import { AREA_LABELS, locate } from "../state/layout";
import type { UiError } from "../state/runner";
import type { Fields, Hierarchy, Layout, Measure } from "../types";
import { cx } from "../util";
import { ErrorNote, Icon, Spinner } from "./common";
import { endDrag, startDrag } from "./drag";

interface Branch {
  type: "branch";
  id: string;
  label: string;
  icon: "measure" | "table" | "folder";
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
type Entry = Branch | MeasureLeaf | HierarchyLeaf;

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

export function buildFieldTree(fields: Fields): Branch[] {
  const roots: Branch[] = [];
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
    roots.push(branch);
  }
  roots.forEach(sortFoldersLast);
  return roots;
}

function captionOf(entry: MeasureLeaf | HierarchyLeaf): string {
  return entry.type === "measure" ? entry.measure.caption : entry.hierarchy.caption;
}

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
}

export function FieldList(props: Props) {
  const { fields, layout, search } = props;
  const tree = useMemo(() => (fields ? buildFieldTree(fields) : []), [fields]);
  const [collapsed, setCollapsed] = useState<ReadonlySet<string>>(new Set());

  useEffect(() => {
    const total = countLeaves(tree);
    setCollapsed(
      new Set(total > AUTO_COLLAPSE_ABOVE ? tree.filter((b) => b.icon === "table").map((b) => b.id) : []),
    );
  }, [tree]);

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
            {open && <ul role="group">{render(entry.children, depth + 1)}</ul>}
          </li>
        );
      }
      const isMeasure = entry.type === "measure";
      const name = entry.id;
      const caption = captionOf(entry);
      const place = locate(layout, name);
      return (
        <li key={entry.id} role="treeitem" aria-selected={place !== null}>
          <label
            className={cx("field-row leaf", place !== null && "placed")}
            style={indent}
            title={caption}
            draggable
            onDragStart={(event) =>
              startDrag(event, {
                kind: isMeasure ? "measure" : "hierarchy",
                name,
                caption,
                allMember: entry.type === "hierarchy" ? entry.hierarchy.allMember : null,
              })
            }
            onDragEnd={endDrag}
          >
            <input
              type="checkbox"
              checked={place !== null}
              onChange={(event) =>
                entry.type === "measure"
                  ? props.onToggleMeasure(entry.measure, event.target.checked)
                  : props.onToggleHierarchy(entry.hierarchy, event.target.checked)
              }
            />
            {!isMeasure && (
              <Icon
                name={entry.type === "hierarchy" && entry.hierarchy.kind === "user" ? "hierarchy" : "attribute"}
                className="kind"
              />
            )}
            <span className="field-caption">{caption}</span>
            {place !== null && !isMeasure && <span className="field-area">{AREA_LABELS[place.area]}</span>}
          </label>
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
