// Member picker for one placed hierarchy: a tree loaded one level at a time, a search
// across all levels, Select all / Select none, and a note when a list was cut off.
// The selection rules live in state/picker.ts. For a Rows or Columns field the panel
// starts with the field menu (sorting, label and value filters, see FieldMenu.tsx), as
// Excel's Row Labels / Column Labels dropdown does; a page filter has the list only.

import { useCallback, useEffect, useRef, useState } from "react";
import type { Api } from "../api";
import {
  ROOT,
  checkState,
  describeSelection,
  initPicker,
  isAll,
  pickerResult,
  selectAll,
  selectNone,
  setChildren,
  toggle,
  toggleFlat,
  type PickerState,
} from "../state/picker";
import { toUiError, type UiError } from "../state/runner";
import type { LayoutAction } from "../state/layout";
import type { Measure, Member, PlacedField } from "../types";
import { FieldMenu } from "./FieldMenu";
import { ErrorNote, Icon, Spinner, TriCheckbox } from "./common";

interface Props {
  api: Api;
  connectionId: string;
  catalog: string;
  cube: string;
  field: PlacedField;
  onApply: (members: string[], captions: string[]) => void;
  onCancel: () => void;
  /** Present for a Rows or Columns field: the sort and filter menu above the list. */
  menu?: {
    measures: Measure[];
    defaultMeasure: string;
    /** Apply a sort or filter change; the panel is closed by the caller. */
    onAction: (action: LayoutAction) => void;
    /** For Expand / Collapse Entire Field; null without a result. */
    expandable: { members: string[]; anyChildren: boolean } | null;
    /** The field's items are dates: offer Date Filters. */
    isDate?: boolean;
    /** Open the field's groups, and its calculated items. */
    onGroups?: () => void;
    onCalculatedItems?: () => void;
  };
  /**
   * Opened from compact form's "Row Labels" or "Column Labels": the fields of that area,
   * to choose from first (Windows' "Select field:"). Shown when there are several.
   */
  choose?: {
    fields: PlacedField[];
    onChoose: (hierarchy: string) => void;
  };
}

const SEARCH_DELAY_MS = 300;

interface SearchState {
  text: string;
  loading: boolean;
  members: Member[];
  truncated: boolean;
  error: UiError | null;
}

export function MemberPicker({ api, connectionId, catalog, cube, field, onApply, onCancel, menu, choose }: Props) {
  const [state, setState] = useState<PickerState>(() => initPicker(field));
  const [open, setOpen] = useState<ReadonlySet<string>>(new Set());
  const [loading, setLoading] = useState<ReadonlySet<string>>(new Set([ROOT]));
  const [error, setError] = useState<UiError | null>(null);
  const [hint, setHint] = useState<string | null>(null);
  const [query, setQuery] = useState("");
  const [search, setSearch] = useState<SearchState | null>(null);
  const alive = useRef(true);
  const box = useRef<HTMLDivElement>(null);

  useEffect(() => {
    alive.current = true;
    return () => {
      alive.current = false;
    };
  }, []);

  const hierarchy = field.hierarchy;

  /** Load the children of `key` (the top level for ROOT). */
  const load = useCallback(
    async (key: string): Promise<void> => {
      setLoading((before) => new Set(before).add(key));
      setError(null);
      try {
        // A field placed as one level starts from that level's items.
        const base = { catalog, cube, hierarchy, ...(field.level ? { level: field.level.uniqueName } : {}) };
        const response = await api.members(connectionId, key === ROOT ? base : { ...base, parent: key });
        if (!alive.current) return;
        setState((before) => setChildren(before, key, response.members, response.truncated));
      } catch (failure) {
        if (alive.current) setError(toUiError(failure));
      } finally {
        if (alive.current) {
          setLoading((before) => {
            const next = new Set(before);
            next.delete(key);
            return next;
          });
        }
      }
    },
    [api, connectionId, catalog, cube, hierarchy],
  );

  // The top level: the hierarchy's first level of members.
  useEffect(() => {
    void load(ROOT);
  }, [load]);

  // Search, after a short pause in typing. A newer search replaces an older one.
  useEffect(() => {
    const text = query.trim();
    if (text === "") {
      setSearch(null);
      return;
    }
    let cancelled = false;
    setSearch((before) => ({
      text,
      loading: true,
      members: before ? before.members : [],
      truncated: false,
      error: null,
    }));
    const timer = window.setTimeout(async () => {
      try {
        const response = await api.members(connectionId, { catalog, cube, hierarchy, search: text });
        if (cancelled) return;
        setSearch({ text, loading: false, members: response.members, truncated: response.truncated, error: null });
      } catch (failure) {
        if (cancelled) return;
        setSearch({ text, loading: false, members: [], truncated: false, error: toUiError(failure) });
      }
    }, SEARCH_DELAY_MS);
    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [query, api, connectionId, catalog, cube, hierarchy]);

  // Keyboard focus starts where the panel starts: the field menu, or else the search.
  useEffect(() => {
    const first =
      box.current?.querySelector<HTMLElement>(".fm-item:not(:disabled)") ??
      box.current?.querySelector<HTMLInputElement>("input[type=search]");
    first?.focus();
  }, []);

  const twist = (key: string) => {
    const node = state.nodes.get(key);
    if (!node) return;
    const next = new Set(open);
    if (next.has(key)) {
      next.delete(key);
    } else {
      next.add(key);
      if (!node.children && !loading.has(key)) void load(key);
    }
    setOpen(next);
  };

  const tick = (key: string) => {
    const result = toggle(state, key);
    setState(result.state);
    setHint(
      result.blocked === "truncated"
        ? "This list is too long to leave one item out. Choose Select none, then tick the items you want."
        : null,
    );
  };

  const result = pickerResult(state, field.allMember);

  const renderNodes = (keys: string[], depth: number): JSX.Element[] =>
    keys.map((key) => {
      const node = state.nodes.get(key);
      if (!node) return <li key={key} />;
      const isOpen = open.has(key);
      return (
        <li key={key} role="treeitem" aria-expanded={node.hasChildren ? isOpen : undefined}>
          <div className="member-row" style={{ paddingLeft: 4 + depth * 14 }}>
            {node.hasChildren ? (
              <button
                type="button"
                className="twist-button"
                onClick={() => twist(key)}
                aria-label={isOpen ? `Collapse ${node.caption}` : `Expand ${node.caption}`}
              >
                <Icon name={isOpen ? "chevron-down" : "chevron-right"} />
              </button>
            ) : (
              <span className="twist-gap" />
            )}
            <label className="member-label" title={node.caption}>
              <TriCheckbox state={checkState(state, key)} onChange={() => tick(key)} label={node.caption} />
              <span className="member-caption">{node.caption}</span>
            </label>
          </div>
          {isOpen && (
            <ul role="group">
              {node.children && renderNodes(node.children, depth + 1)}
              {loading.has(key) && (
                <li className="member-note" style={{ paddingLeft: 24 + (depth + 1) * 14 }}>
                  <Spinner label="Loading" /> Loading…
                </li>
              )}
              {node.truncated && node.children && (
                <li className="member-note" style={{ paddingLeft: 24 + (depth + 1) * 14 }}>
                  Showing the first {node.children.length.toLocaleString("en-US")} items. Use search to
                  find the others.
                </li>
              )}
            </ul>
          )}
        </li>
      );
    });

  const root = state.nodes.get(ROOT);
  const everything = isAll(state) || (field.allMember !== null && state.selected.has(field.allMember));

  return (
    <div className="sheet" role="dialog" aria-modal="true" aria-label={`Filter ${field.caption}`} ref={box}
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.stopPropagation();
          onCancel();
        }
      }}
    >
      <header className="bar">
        <h2 className="bar-title">
          {menu ? "Sort and filter" : "Filter"}: {field.caption}
        </h2>
        <button type="button" className="icon-button" onClick={onCancel} aria-label="Close">
          <Icon name="close" />
        </button>
      </header>

      {choose && choose.fields.length > 1 && (
        <label className="row-field picker-choose">
          <span className="row-label">Select field:</span>
          <select
            className="select"
            value={field.hierarchy}
            onChange={(event) => choose.onChoose(event.target.value)}
          >
            {choose.fields.map((f) => (
              <option key={f.hierarchy} value={f.hierarchy}>
                {f.caption}
              </option>
            ))}
          </select>
        </label>
      )}

      <div className="picker-body">
      {menu && (
        <>
          <FieldMenu
            field={field}
            measures={menu.measures}
            defaultMeasure={menu.defaultMeasure}
            onAction={menu.onAction}
            expandable={menu.expandable}
            isDate={menu.isDate}
            onGroups={menu.onGroups}
            onCalculatedItems={menu.onCalculatedItems}
          />
          <h3 className="picker-heading">Select items</h3>
        </>
      )}

      <div className="search picker-search">
        <Icon name="search" className="search-icon" />
        <input
          className="input search-input"
          type="search"
          value={query}
          onChange={(event) => setQuery(event.target.value)}
          placeholder={`Search ${field.caption}`}
          aria-label={`Search ${field.caption}`}
          autoCapitalize="off"
          autoCorrect="off"
          spellCheck={false}
        />
      </div>

      <div className="picker-tools">
        <button type="button" className="link" onClick={() => { setState(selectAll(state)); setHint(null); }}>
          Select all
        </button>
        <button type="button" className="link" onClick={() => { setState(selectNone(state)); setHint(null); }}>
          Select none
        </button>
      </div>

      <div className="picker-list">
        {error && <ErrorNote error={error} action={{ label: "Try again", run: () => void load(ROOT) }} />}
        {search ? (
          <>
            {search.error && <ErrorNote error={search.error} />}
            {everything && search.members.length > 0 && (
              <p className="member-note">Everything is selected now. Tick an item to show only that item.</p>
            )}
            <ul className="member-tree">
              {search.members.map((member) => (
                <li key={member.uniqueName}>
                  <div className="member-row" style={{ paddingLeft: 8 }}>
                    <label className="member-label" title={member.caption}>
                      <input
                        type="checkbox"
                        checked={state.selected.has(member.uniqueName)}
                        onChange={() => setState(toggleFlat(state, member, field.allMember))}
                      />
                      <span className="member-caption">{member.caption}</span>
                    </label>
                  </div>
                </li>
              ))}
            </ul>
            {search.loading && (
              <p className="member-note">
                <Spinner label="Searching" /> Searching…
              </p>
            )}
            {!search.loading && !search.error && search.members.length === 0 && (
              <p className="member-note">Nothing in {field.caption} matches &ldquo;{search.text}&rdquo;.</p>
            )}
            {search.truncated && (
              <p className="member-note">
                Showing the first {search.members.length.toLocaleString("en-US")} matches. Type more to
                narrow the search.
              </p>
            )}
          </>
        ) : (
          <>
            <ul className="member-tree" role="tree">
              {root && root.children && renderNodes(root.children, 0)}
            </ul>
            {loading.has(ROOT) && (
              <p className="member-note">
                <Spinner label="Loading" /> Loading…
              </p>
            )}
            {root && root.truncated && root.children && (
              <p className="member-note">
                Showing the first {root.children.length.toLocaleString("en-US")} items. Use search to find
                the others.
              </p>
            )}
          </>
        )}
      </div>

      {hint && <p className="picker-hint">{hint}</p>}
      </div>

      <footer className="picker-foot">
        <span className="picker-summary" title={result ? result.captions.join(", ") : undefined}>
          {result === null ? "Tick at least one item." : describeSelection(state, field.allMember)}
        </span>
        <button type="button" className="button" onClick={onCancel}>
          Cancel
        </button>
        <button
          type="button"
          className="button primary"
          disabled={result === null}
          onClick={() => result && onApply(result.members, result.captions)}
        >
          Apply
        </button>
      </footer>
    </div>
  );
}
