// The top of the panel for a Rows or Columns field, as in Excel's Row Labels / Column
// Labels dropdown on Windows: sorting, Label Filters, Value Filters, and clearing every
// filter of the field. The item checklist follows below it (MemberPicker). Each choice
// applies at once and closes the panel, as Excel's menu does.

import { useState, type FormEvent, type ReactNode } from "react";
import {
  LABEL_OPERATORS,
  TOP_LABEL,
  VALUE_OPERATORS,
  checkLabelFilter,
  checkTopFilter,
  checkValueFilter,
  describeLabelFilter,
  describeSort,
  describeValueFilter,
} from "../state/fieldMenu";
import { isFiltered, type LayoutAction } from "../state/layout";
import type { LabelOperator, Measure, PlacedField, ValueOperator } from "../types";
import { cx } from "../util";
import { Icon } from "./common";

interface Props {
  field: PlacedField;
  /** Every measure of the cube, for sorting by a value and for value filters. */
  measures: Measure[];
  /** Measure the forms start on: the first value field, else the cube's first. */
  defaultMeasure: string;
  /** Apply a change to this field. The panel closes afterwards. */
  onAction: (action: LayoutAction) => void;
  /**
   * Expand / Collapse Entire Field: the members of the last result that can be opened
   * one level deeper, and whether any shown item has children. Null when there is no
   * result to go by.
   */
  expandable: { members: string[]; anyChildren: boolean } | null;
}

type Mode =
  | { kind: "menu" }
  | { kind: "sort" }
  | { kind: "labels" }
  | { kind: "label"; operator: LabelOperator }
  | { kind: "values" }
  | { kind: "value"; operator: ValueOperator }
  | { kind: "top" };

function MenuButton(props: {
  children: ReactNode;
  onClick: () => void;
  checked?: boolean;
  disabled?: boolean;
  expanded?: boolean;
  danger?: boolean;
  indent?: boolean;
}) {
  return (
    <button
      type="button"
      className={cx("fm-item", props.indent && "indent", props.danger && "danger")}
      onClick={props.onClick}
      disabled={props.disabled}
      aria-expanded={props.expanded}
      aria-pressed={props.checked === undefined ? undefined : props.checked}
    >
      <span className="fm-check" aria-hidden="true">
        {props.checked ? "✓" : ""}
      </span>
      <span className="fm-label">{props.children}</span>
      {props.expanded !== undefined && (
        <Icon name={props.expanded ? "chevron-down" : "chevron-right"} className="fm-twist" />
      )}
    </button>
  );
}

function FormButtons({ onCancel }: { onCancel: () => void }) {
  return (
    <div className="fm-buttons">
      <button type="button" className="button small" onClick={onCancel}>
        Cancel
      </button>
      <button type="submit" className="button small primary">
        OK
      </button>
    </div>
  );
}

function MeasureSelect(props: {
  measures: Measure[];
  value: string;
  onChange: (value: string) => void;
  label: string;
}) {
  const known = props.measures.some((m) => m.uniqueName === props.value);
  return (
    <select
      className="select"
      aria-label={props.label}
      value={props.value}
      onChange={(event) => props.onChange(event.target.value)}
    >
      {!known && <option value={props.value}>{props.value || "Choose a value"}</option>}
      {props.measures.map((m) => (
        <option key={m.uniqueName} value={m.uniqueName}>
          {m.caption}
        </option>
      ))}
    </select>
  );
}

export function FieldMenu({ field, measures, defaultMeasure, onAction, expandable }: Props) {
  const [mode, setMode] = useState<Mode>({ kind: "menu" });
  const [error, setError] = useState<string | null>(null);
  const caption = (name: string) => measures.find((m) => m.uniqueName === name)?.caption ?? name;

  // Form state, started from what the field has now.
  const label = field.labelFilter;
  const value = field.valueFilter;
  const isTop = value && (value.operator === "top" || value.operator === "bottom");
  const [text1, setText1] = useState(label ? label.value : "");
  const [text2, setText2] = useState(label && label.value2 ? label.value2 : "");
  const [number1, setNumber1] = useState(value && !isTop ? String(value.value) : "");
  const [number2, setNumber2] = useState(value && !isTop && value.value2 !== undefined ? String(value.value2) : "");
  const [measure, setMeasure] = useState(value ? value.measure : defaultMeasure);
  const [direction, setDirection] = useState<"top" | "bottom">(value && value.operator === "bottom" ? "bottom" : "top");
  const [count, setCount] = useState(isTop && value ? String(value.value) : "10");
  const [topBy, setTopBy] = useState<"items" | "percent" | "sum">(isTop && value ? value.topBy ?? "items" : "items");
  const [sortOrder, setSortOrder] = useState<"asc" | "desc">(field.sort ? field.sort.order : "asc");
  const [sortMeasure, setSortMeasure] = useState(
    field.sort && field.sort.by === "measure" && field.sort.measure ? field.sort.measure : defaultMeasure,
  );

  const go = (next: Mode) => {
    setError(null);
    setMode(next);
  };
  const back = () => go({ kind: "menu" });
  const hierarchy = field.hierarchy;
  const sort = field.sort ?? null;

  // ---- forms -------------------------------------------------------------

  if (mode.kind === "sort") {
    const submit = (event: FormEvent) => {
      event.preventDefault();
      if (sortMeasure === "") {
        setError("Choose a value to sort by.");
        return;
      }
      onAction({ type: "setSort", hierarchy, sort: { by: "measure", measure: sortMeasure, order: sortOrder } });
    };
    return (
      <form className="fm-form" onSubmit={submit} aria-label="More sort options">
        <h3 className="fm-title">Sort ({field.caption})</h3>
        <label className="fm-field">
          <span className="field-label">Order</span>
          <select className="select" value={sortOrder} onChange={(e) => setSortOrder(e.target.value as "asc" | "desc")}>
            <option value="asc">Ascending (smallest to largest)</option>
            <option value="desc">Descending (largest to smallest)</option>
          </select>
        </label>
        <label className="fm-field">
          <span className="field-label">By</span>
          <MeasureSelect measures={measures} value={sortMeasure} onChange={setSortMeasure} label="Sort by" />
        </label>
        {error && <p className="fm-error">{error}</p>}
        <FormButtons onCancel={back} />
      </form>
    );
  }

  if (mode.kind === "label") {
    const choice = LABEL_OPERATORS.find((c) => c.operator === mode.operator) ?? LABEL_OPERATORS[0];
    const submit = (event: FormEvent) => {
      event.preventDefault();
      const checked = checkLabelFilter({ operator: choice.operator, value: text1, value2: text2 });
      if (!checked.ok) {
        setError(checked.message);
        return;
      }
      onAction({ type: "setLabelFilter", hierarchy, filter: checked.value });
    };
    return (
      <form className="fm-form" onSubmit={submit} aria-label="Label filter">
        <h3 className="fm-title">Label Filter ({field.caption})</h3>
        <p className="fm-text">Show items for which the label</p>
        <select
          className="select"
          aria-label="Condition"
          value={choice.operator}
          onChange={(e) => go({ kind: "label", operator: e.target.value as LabelOperator })}
        >
          {LABEL_OPERATORS.map((c) => (
            <option key={c.operator} value={c.operator}>
              {c.phrase}
            </option>
          ))}
        </select>
        <input
          className="input"
          value={text1}
          onChange={(e) => setText1(e.target.value)}
          aria-label="Text"
          autoFocus
          autoCapitalize="off"
          autoCorrect="off"
          spellCheck={false}
        />
        {choice.two && (
          <>
            <span className="fm-text">and</span>
            <input
              className="input"
              value={text2}
              onChange={(e) => setText2(e.target.value)}
              aria-label="Second text"
              autoCapitalize="off"
              autoCorrect="off"
              spellCheck={false}
            />
          </>
        )}
        {error && <p className="fm-error">{error}</p>}
        <FormButtons onCancel={back} />
      </form>
    );
  }

  if (mode.kind === "value") {
    const choice = VALUE_OPERATORS.find((c) => c.operator === mode.operator) ?? VALUE_OPERATORS[0];
    const submit = (event: FormEvent) => {
      event.preventDefault();
      const checked = checkValueFilter({ operator: choice.operator, measure, value: number1, value2: number2 });
      if (!checked.ok) {
        setError(checked.message);
        return;
      }
      onAction({ type: "setValueFilter", hierarchy, filter: checked.value });
    };
    return (
      <form className="fm-form" onSubmit={submit} aria-label="Value filter">
        <h3 className="fm-title">Value Filter ({field.caption})</h3>
        <p className="fm-text">Show items for which</p>
        <MeasureSelect measures={measures} value={measure} onChange={setMeasure} label="Value" />
        <select
          className="select"
          aria-label="Condition"
          value={choice.operator}
          onChange={(e) => go({ kind: "value", operator: e.target.value as ValueOperator })}
        >
          {VALUE_OPERATORS.map((c) => (
            <option key={c.operator} value={c.operator}>
              {c.phrase}
            </option>
          ))}
        </select>
        <input
          className="input"
          value={number1}
          onChange={(e) => setNumber1(e.target.value)}
          aria-label="Number"
          inputMode="decimal"
          autoFocus
        />
        {choice.two && (
          <>
            <span className="fm-text">and</span>
            <input
              className="input"
              value={number2}
              onChange={(e) => setNumber2(e.target.value)}
              aria-label="Second number"
              inputMode="decimal"
            />
          </>
        )}
        {error && <p className="fm-error">{error}</p>}
        <FormButtons onCancel={back} />
      </form>
    );
  }

  if (mode.kind === "top") {
    const submit = (event: FormEvent) => {
      event.preventDefault();
      const checked = checkTopFilter({ direction, count, topBy, measure });
      if (!checked.ok) {
        setError(checked.message);
        return;
      }
      onAction({ type: "setValueFilter", hierarchy, filter: checked.value });
    };
    return (
      <form className="fm-form" onSubmit={submit} aria-label="Top 10 filter">
        <h3 className="fm-title">Top 10 Filter ({field.caption})</h3>
        <p className="fm-text">Show</p>
        <div className="fm-row">
          <select
            className="select"
            aria-label="Top or bottom"
            value={direction}
            onChange={(e) => setDirection(e.target.value as "top" | "bottom")}
          >
            <option value="top">Top</option>
            <option value="bottom">Bottom</option>
          </select>
          <input
            className="input fm-count"
            value={count}
            onChange={(e) => setCount(e.target.value)}
            aria-label="How many"
            inputMode="decimal"
            autoFocus
          />
          <select
            className="select"
            aria-label="Items, percent or sum"
            value={topBy}
            onChange={(e) => setTopBy(e.target.value as "items" | "percent" | "sum")}
          >
            <option value="items">Items</option>
            <option value="percent">Percent</option>
            <option value="sum">Sum</option>
          </select>
        </div>
        <span className="fm-text">by</span>
        <MeasureSelect measures={measures} value={measure} onChange={setMeasure} label="By" />
        {error && <p className="fm-error">{error}</p>}
        <FormButtons onCancel={back} />
      </form>
    );
  }

  // ---- the menu ----------------------------------------------------------

  const labelsOpen = mode.kind === "labels";
  const valuesOpen = mode.kind === "values";
  return (
    <section className="fm" aria-label={`Sort and filter ${field.caption}`}>
      <MenuButton
        checked={sort !== null && sort.by === "caption" && sort.order === "asc"}
        onClick={() => onAction({ type: "setSort", hierarchy, sort: { by: "caption", order: "asc" } })}
      >
        Sort A to Z
      </MenuButton>
      <MenuButton
        checked={sort !== null && sort.by === "caption" && sort.order === "desc"}
        onClick={() => onAction({ type: "setSort", hierarchy, sort: { by: "caption", order: "desc" } })}
      >
        Sort Z to A
      </MenuButton>
      <MenuButton checked={sort !== null && sort.by === "measure"} onClick={() => go({ kind: "sort" })}>
        More sort options…
        {sort && sort.by === "measure" && <span className="fm-state">{describeSort(sort, caption)}</span>}
      </MenuButton>
      <MenuButton checked={sort === null} onClick={() => onAction({ type: "setSort", hierarchy, sort: null })}>
        Cube order
      </MenuButton>

      <div className="fm-sep" role="separator" />

      <MenuButton expanded={labelsOpen} onClick={() => go(labelsOpen ? { kind: "menu" } : { kind: "labels" })}>
        Label Filters
        {label && <span className="fm-state">{describeLabelFilter(label)}</span>}
      </MenuButton>
      {labelsOpen && (
        <div className="fm-sub" role="group" aria-label="Label Filters">
          <MenuButton indent disabled={!label} onClick={() => onAction({ type: "setLabelFilter", hierarchy, filter: null })}>
            Clear Label Filter
          </MenuButton>
          {LABEL_OPERATORS.map((c) => (
            <MenuButton
              key={c.operator}
              indent
              checked={Boolean(label && label.operator === c.operator)}
              onClick={() => go({ kind: "label", operator: c.operator })}
            >
              {c.label}
            </MenuButton>
          ))}
        </div>
      )}

      <MenuButton expanded={valuesOpen} onClick={() => go(valuesOpen ? { kind: "menu" } : { kind: "values" })}>
        Value Filters
        {value && <span className="fm-state">{describeValueFilter(value, caption)}</span>}
      </MenuButton>
      {valuesOpen && (
        <div className="fm-sub" role="group" aria-label="Value Filters">
          <MenuButton indent disabled={!value} onClick={() => onAction({ type: "setValueFilter", hierarchy, filter: null })}>
            Clear Value Filter
          </MenuButton>
          {VALUE_OPERATORS.map((c) => (
            <MenuButton
              key={c.operator}
              indent
              checked={Boolean(value && value.operator === c.operator)}
              onClick={() => go({ kind: "value", operator: c.operator })}
            >
              {c.label}
            </MenuButton>
          ))}
          <MenuButton indent checked={Boolean(isTop)} onClick={() => go({ kind: "top" })}>
            {TOP_LABEL}
          </MenuButton>
        </div>
      )}

      <div className="fm-sep" role="separator" />
      <MenuButton
        disabled={!expandable || !expandable.anyChildren || expandable.members.length === 0}
        onClick={() => expandable && onAction({ type: "expandField", hierarchy, members: expandable.members })}
      >
        Expand Entire Field
      </MenuButton>
      <MenuButton
        disabled={!expandable || !expandable.anyChildren || field.expanded.length === 0}
        onClick={() => onAction({ type: "collapseField", hierarchy })}
      >
        Collapse Entire Field
      </MenuButton>

      <div className="fm-sep" role="separator" />
      <MenuButton
        disabled={!isFiltered(field)}
        danger
        onClick={() => onAction({ type: "clearFilters", hierarchy })}
      >
        Clear Filter From &ldquo;{field.caption}&rdquo;
      </MenuButton>
    </section>
  );
}
