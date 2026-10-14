// The report's slicers, as cards in the panel: Excel puts slicers on the sheet, but the
// pane cannot draw buttons there. Each card has the field's name, Excel's Multi-Select
// and Clear Filter buttons, a button to remove the slicer, and one button per top-level
// item. Selected items are highlighted, every item when nothing is filtered, as in Excel.
// What a click does is in state/slicers.ts; the choice re-runs the report like a filter.

import { useCallback, useEffect, useState, type KeyboardEvent, type MouseEvent } from "react";
import type { Api } from "../api";
import type { LayoutAction } from "../state/layout";
import { toUiError, type UiError } from "../state/runner";
import { isSlicerSelected, slicerClick, slicerExtras, type SlicerItem } from "../state/slicers";
import type { PlacedField } from "../types";
import { cx } from "../util";
import { ErrorNote, Icon, Spinner } from "./common";
import "./Slicers.css";

interface Props {
  api: Api;
  /** Null when the report has no connection on this Mac: the items cannot be listed. */
  connectionId: string | null;
  catalog: string;
  cube: string;
  slicers: PlacedField[];
  dispatch: (action: LayoutAction) => void;
}

export function Slicers({ api, connectionId, catalog, cube, slicers, dispatch }: Props) {
  return (
    <section className="slicers" aria-label="Slicers">
      <h3 className="area-title slicers-title">Slicers</h3>
      {slicers.map((field) => (
        <SlicerCard
          key={field.hierarchy}
          api={api}
          connectionId={connectionId}
          catalog={catalog}
          cube={cube}
          field={field}
          onSelect={(members, captions) => dispatch({ type: "setMembers", hierarchy: field.hierarchy, members, captions })}
          onRemove={() => dispatch({ type: "removeSlicer", hierarchy: field.hierarchy })}
        />
      ))}
    </section>
  );
}

interface CardProps {
  api: Api;
  connectionId: string | null;
  catalog: string;
  cube: string;
  field: PlacedField;
  onSelect: (members: string[], captions: string[]) => void;
  onRemove: () => void;
}

interface Items {
  loading: boolean;
  list: SlicerItem[];
  truncated: boolean;
  error: UiError | null;
}

/** Excel's Multi-Select button: a list with ticks. */
function MultiSelectIcon() {
  return (
    <svg className="icon" viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" focusable="false">
      <path d="M2 4l1.2 1.2L5.5 3M2 9l1.2 1.2L5.5 8M7.5 4.2h6.5M7.5 9.2h6.5M7.5 13.2h6.5M2.5 13.2h2" />
    </svg>
  );
}

/** Excel's Clear Filter button: a funnel with a red cross. */
function ClearFilterIcon() {
  return (
    <svg className="icon" viewBox="0 0 16 16" width="14" height="14" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true" focusable="false">
      <path d="M1.5 2.5h10L8 7v4.5l-3 1.5V7z" />
      <path className="slicer-cross" d="M10.5 10l4 4M14.5 10l-4 4" />
    </svg>
  );
}

function SlicerCard({ api, connectionId, catalog, cube, field, onSelect, onRemove }: CardProps) {
  const [items, setItems] = useState<Items>({ loading: true, list: [], truncated: false, error: null });
  const [multi, setMulti] = useState(false);
  const [hint, setHint] = useState<string | null>(null);
  const [attempt, setAttempt] = useState(0);

  const hierarchy = field.hierarchy;
  useEffect(() => {
    if (!connectionId) {
      setItems({ loading: false, list: [], truncated: false, error: null });
      return;
    }
    let current = true;
    setItems((before) => ({ ...before, loading: true, error: null }));
    api.members(connectionId, { catalog, cube, hierarchy }).then(
      (response) => {
        if (!current) return;
        setItems({
          loading: false,
          list: response.members.map((m) => ({ uniqueName: m.uniqueName, caption: m.caption })),
          truncated: response.truncated,
          error: null,
        });
      },
      (failure) => {
        if (current) setItems({ loading: false, list: [], truncated: false, error: toUiError(failure) });
      },
    );
    return () => {
      current = false;
    };
  }, [api, connectionId, catalog, cube, hierarchy, attempt]);

  const click = useCallback(
    (item: SlicerItem, add: boolean) => {
      const result = slicerClick(field, items.list, items.truncated, item, multi || add);
      if ("blocked" in result) {
        setHint(
          result.blocked === "truncated"
            ? "This list is too long to leave one item out. Click the item you want, then ⌘-click the others."
            : null,
        );
        return;
      }
      setHint(null);
      if ("members" in result) onSelect(result.members, result.captions);
    },
    [field, items.list, items.truncated, multi, onSelect],
  );

  const filtered = field.members.length > 0;
  const clear = () => {
    setHint(null);
    if (filtered) onSelect([], []);
  };
  const extras = slicerExtras(field, items.list);

  const onKeyDown = (event: KeyboardEvent) => {
    // Excel's shortcuts for a selected slicer: Alt+S Multi-Select, Alt+C Clear Filter.
    if (!event.altKey || event.metaKey || event.ctrlKey) return;
    if (event.code === "KeyS") {
      event.preventDefault();
      setMulti((on) => !on);
    } else if (event.code === "KeyC") {
      event.preventDefault();
      clear();
    }
  };

  return (
    <div className="slicer" role="group" aria-label={`Slicer ${field.caption}`} onKeyDown={onKeyDown}>
      <div className="slicer-head">
        <span className="slicer-caption" title={field.caption}>
          {field.caption}
        </span>
        <button
          type="button"
          className={cx("slicer-tool", multi && "on")}
          aria-pressed={multi}
          title="Multi-Select (Alt+S)"
          aria-label="Multi-Select"
          onClick={() => setMulti(!multi)}
        >
          <MultiSelectIcon />
        </button>
        <button
          type="button"
          className="slicer-tool slicer-clear"
          disabled={!filtered}
          title="Clear Filter (Alt+C)"
          aria-label={`Clear Filter From ${field.caption}`}
          onClick={clear}
        >
          <ClearFilterIcon />
        </button>
        <button type="button" className="slicer-tool" title={`Remove “${field.caption}”`} aria-label={`Remove slicer ${field.caption}`} onClick={onRemove}>
          <Icon name="close" />
        </button>
      </div>

      {!connectionId ? (
        <p className="slicer-note">Choose a connection for this report to see the items.</p>
      ) : items.error ? (
        <ErrorNote error={items.error} action={{ label: "Try again", run: () => setAttempt((n) => n + 1) }} />
      ) : (
        <div className="slicer-items" role="group" aria-label={`Items of ${field.caption}`}>
          {items.list.map((item) => {
            const selected = isSlicerSelected(field, item.uniqueName);
            return (
              <button
                key={item.uniqueName}
                type="button"
                className={cx("slicer-item", selected && "selected")}
                aria-pressed={selected}
                title={item.caption}
                onClick={(event: MouseEvent) => click(item, event.metaKey || event.ctrlKey)}
              >
                {item.caption}
              </button>
            );
          })}
          {items.loading && (
            <span className="slicer-note">
              <Spinner label="Loading" /> Loading…
            </span>
          )}
        </div>
      )}
      {items.truncated && (
        <p className="slicer-note">Showing the first {items.list.length.toLocaleString("en-US")} items.</p>
      )}
      {extras.length > 0 && (
        <p className="slicer-note" title={extras.map((e) => e.caption).join(", ")}>
          Also selected: {extras.map((e) => e.caption).join(", ")}
        </p>
      )}
      {hint && <p className="slicer-note slicer-hint">{hint}</p>}
    </div>
  );
}
