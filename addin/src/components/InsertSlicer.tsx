// Windows' Insert Slicers dialog (PivotTable Analyze > Insert Slicer): the cube's fields
// by dimension, with a search and a tick for each. The fields that already have a slicer
// start ticked; OK makes exactly the ticked ones slicers. A field in Rows, Columns or
// Filters is marked so: ticking it moves it to a slicer, as a field is in one place only.

import { useEffect, useMemo, useRef, useState } from "react";
import { AREA_LABELS, MAX_SLICERS, locate, slicerIndex, type HierarchyInput } from "../state/layout";
import type { Fields, Layout } from "../types";
import { cx, plural } from "../util";
import { Icon } from "./common";
import "./Slicers.css";

interface Props {
  fields: Fields;
  layout: Layout;
  onApply: (hierarchies: HierarchyInput[]) => void;
  onCancel: () => void;
}

export function InsertSlicer({ fields, layout, onApply, onCancel }: Props) {
  const [ticked, setTicked] = useState<string[]>(() => (layout.slicers ?? []).map((s) => s.hierarchy));
  const [search, setSearch] = useState("");
  const box = useRef<HTMLDivElement>(null);

  useEffect(() => {
    box.current?.querySelector<HTMLInputElement>("input[type=search]")?.focus();
  }, []);

  /** Every hierarchy the dialog can name, by unique name. */
  const known = useMemo(() => {
    const map = new Map<string, HierarchyInput>();
    for (const d of fields.dimensions) {
      for (const h of d.hierarchies) map.set(h.uniqueName, { uniqueName: h.uniqueName, caption: h.caption, allMember: h.allMember });
    }
    for (const s of layout.slicers ?? []) {
      if (!map.has(s.hierarchy)) map.set(s.hierarchy, { uniqueName: s.hierarchy, caption: s.caption, allMember: s.allMember });
    }
    return map;
  }, [fields, layout.slicers]);

  const needle = search.trim().toLowerCase();
  const groups = fields.dimensions
    .map((d) => ({
      dimension: d,
      hierarchies: d.hierarchies.filter(
        (h) => needle === "" || h.caption.toLowerCase().indexOf(needle) >= 0 || d.caption.toLowerCase().indexOf(needle) >= 0,
      ),
    }))
    .filter((g) => g.hierarchies.length > 0);
  const full = ticked.length >= MAX_SLICERS;

  const toggle = (name: string) =>
    setTicked((list) => (list.indexOf(name) >= 0 ? list.filter((n) => n !== name) : [...list, name]));

  const apply = () =>
    onApply(ticked.map((name) => known.get(name)).filter((h): h is HierarchyInput => h !== undefined));

  return (
    <div
      className="sheet"
      role="dialog"
      aria-modal="true"
      aria-label="Insert Slicers"
      ref={box}
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.stopPropagation();
          onCancel();
        }
      }}
    >
      <header className="bar">
        <h2 className="bar-title">Insert Slicers</h2>
        <button type="button" className="icon-button" onClick={onCancel} aria-label="Close">
          <Icon name="close" />
        </button>
      </header>

      <div className="search picker-search">
        <Icon name="search" className="search-icon" />
        <input
          className="input search-input"
          type="search"
          value={search}
          onChange={(event) => setSearch(event.target.value)}
          placeholder="Search fields"
          aria-label="Search fields"
          autoCapitalize="off"
          autoCorrect="off"
          spellCheck={false}
        />
      </div>

      <div className="picker-list insert-slicer-list">
        {groups.length === 0 && <p className="member-note">No field matches &ldquo;{search.trim()}&rdquo;.</p>}
        {groups.map(({ dimension, hierarchies }) => (
          <div key={dimension.uniqueName} role="group" aria-label={dimension.caption}>
            <div className="insert-slicer-dimension">{dimension.caption}</div>
            {hierarchies.map((h) => {
              const on = ticked.indexOf(h.uniqueName) >= 0;
              const place = locate(layout, h.uniqueName);
              const field = place && place.area !== "values" ? layout[place.area][place.index] : null;
              const set = field !== null && Boolean(field.set);
              const where = place && slicerIndex(layout, h.uniqueName) < 0 ? AREA_LABELS[place.area] : null;
              const disabled = set || (!on && full);
              return (
                <label
                  key={h.uniqueName}
                  className={cx("member-row insert-slicer-row", disabled && "unavailable")}
                  title={
                    set
                      ? `The set “${field?.caption}” is using this field.`
                      : where
                        ? `In ${where} now. A slicer takes it from there.`
                        : h.caption
                  }
                >
                  <input type="checkbox" checked={on} disabled={disabled} onChange={() => toggle(h.uniqueName)} />
                  <span className="member-caption field-caption">{h.caption}</span>
                  {where && <span className="field-area">{where}</span>}
                </label>
              );
            })}
          </div>
        ))}
      </div>

      <footer className="picker-foot">
        <span className="picker-summary">
          {full ? `At most ${MAX_SLICERS} slicers.` : ticked.length === 0 ? "No slicers." : plural(ticked.length, "slicer", "slicers")}
        </span>
        <button type="button" className="button" onClick={onCancel}>
          Cancel
        </button>
        <button type="button" className="button primary" onClick={apply}>
          OK
        </button>
      </footer>
    </div>
  );
}
