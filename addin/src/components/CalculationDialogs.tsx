// Dialogs for what the report calculates itself: a field's groups ("Group Items…"), its
// calculated items ("Calculated Item…"), and calculated measures (Windows: OLAP Tools >
// MDX Calculated Measure). Each one edits a copy and hands back the whole list on OK.
// A formula can be tried on the cube first ("Check"), as Windows' "Test MDX" does.

import { useEffect, useMemo, useRef, useState } from "react";
import type { Api } from "../api";
import { CALCULATED_FORMATS, checkFormula, checkNewName } from "../state/fieldMenu";
import { DEFAULT_OPTIONS, calculatedMeasureName } from "../state/layout";
import { toUiError, type UiError } from "../state/runner";
import type { CalculatedMeasure, CustomItem, Layout, Measure, Member, PlacedField, ValueField } from "../types";
import { Dialog, ErrorNote, Spinner } from "./common";

interface CubeRef {
  api: Api;
  connectionId: string;
  catalog: string;
  cube: string;
}

type CheckState = { kind: "idle" } | { kind: "checking" } | { kind: "ok" } | { kind: "failed"; error: UiError };

/** Run a tiny query with the formula in it: the cube's answer says whether it works. */
function useFormulaCheck(ref: CubeRef) {
  const [state, setState] = useState<CheckState>({ kind: "idle" });
  const alive = useRef(true);
  useEffect(
    () => () => {
      alive.current = false;
    },
    [],
  );
  const check = async (layout: Layout) => {
    setState({ kind: "checking" });
    try {
      await ref.api.query(ref.connectionId, { catalog: ref.catalog, cube: ref.cube, layout });
      if (alive.current) setState({ kind: "ok" });
    } catch (failure) {
      if (alive.current) setState({ kind: "failed", error: toUiError(failure) });
    }
  };
  return { state, check, reset: () => setState({ kind: "idle" }) };
}

function CheckResult({ state }: { state: CheckState }) {
  if (state.kind === "checking") {
    return (
      <p className="dialog-note">
        <Spinner label="Checking" /> Asking the cube…
      </p>
    );
  }
  if (state.kind === "ok") return <p className="dialog-note ok">The cube accepts this formula.</p>;
  if (state.kind === "failed") return <ErrorNote error={state.error} />;
  return null;
}

/** The field's top-level items, for picking group members and inserting into formulas. */
function useTopItems(ref: CubeRef, hierarchy: string) {
  const [items, setItems] = useState<Member[] | null>(null);
  const [error, setError] = useState<UiError | null>(null);
  const [truncated, setTruncated] = useState(false);
  useEffect(() => {
    let alive = true;
    ref.api
      .members(ref.connectionId, { catalog: ref.catalog, cube: ref.cube, hierarchy, limit: 5000 })
      .then(
        (found) => {
          if (!alive) return;
          setItems(found.members);
          setTruncated(found.truncated);
        },
        (failure) => alive && setError(toUiError(failure)),
      );
    return () => {
      alive = false;
    };
  }, [ref.api, ref.connectionId, ref.catalog, ref.cube, hierarchy]);
  return { items, error, truncated };
}

/** Insert `text` into a textarea at its caret, and put the caret after it. */
function insertAtCaret(area: HTMLTextAreaElement | null, value: string, text: string, set: (next: string) => void) {
  if (!area) {
    set(value + text);
    return;
  }
  const start = area.selectionStart ?? value.length;
  const end = area.selectionEnd ?? value.length;
  set(value.slice(0, start) + text + value.slice(end));
  requestAnimationFrame(() => {
    area.focus();
    area.selectionStart = area.selectionEnd = start + text.length;
  });
}

function nextName(prefix: string, taken: readonly string[]): string {
  for (let n = 1; ; n++) {
    const name = `${prefix}${n}`;
    if (!taken.some((t) => t.toLowerCase() === name.toLowerCase())) return name;
  }
}

// ---------------------------------------------------------------------------
// Group Items
// ---------------------------------------------------------------------------

export function GroupDialog(
  props: CubeRef & { field: PlacedField; onApply: (items: CustomItem[]) => void; onCancel: () => void },
) {
  const { field } = props;
  const { items, error, truncated } = useTopItems(props, field.hierarchy);
  const [list, setList] = useState<CustomItem[]>(field.customItems ?? []);
  const taken = list.map((i) => i.name);
  const [name, setName] = useState(() => nextName("Group", taken));
  const [chosen, setChosen] = useState<ReadonlySet<string>>(new Set());
  const [search, setSearch] = useState("");
  const [problem, setProblem] = useState<string | null>(null);

  const groupOf = useMemo(() => {
    const map = new Map<string, string>();
    for (const item of list) for (const m of item.members ?? []) map.set(m, item.name);
    return map;
  }, [list]);
  const needle = search.trim().toLowerCase();
  const shown = (items ?? []).filter((m) => needle === "" || m.caption.toLowerCase().indexOf(needle) >= 0);
  const captionOf = (unique: string) => items?.find((m) => m.uniqueName === unique)?.caption ?? unique;

  const addGroup = (): CustomItem[] | null => {
    if (chosen.size === 0) return list;
    const checked = checkNewName(name, taken, "group");
    if (!checked.ok) {
      setProblem(checked.message);
      return null;
    }
    if (chosen.size < 2) {
      setProblem("A group needs at least two items.");
      return null;
    }
    const members = (items ?? []).map((m) => m.uniqueName).filter((u) => chosen.has(u));
    return [...list, { name: checked.value, members }];
  };

  return (
    <Dialog
      title={`Group Items (${field.caption})`}
      wide
      onCancel={props.onCancel}
      buttons={[
        { label: "Cancel", run: props.onCancel },
        {
          label: "OK",
          primary: true,
          run: () => {
            const next = addGroup();
            if (next) props.onApply(next);
          },
        },
      ]}
    >
      <p className="dialog-text">
        Tick two or more items to show them as one item that adds them up. Grouped items are not shown on their own.
        Excel on Windows cannot group a cube field; Cube Explorer can.
      </p>
      {list.some((i) => i.members && i.members.length > 0) && (
        <ul className="calc-list" aria-label="Groups">
          {list
            .filter((i) => i.members && i.members.length > 0)
            .map((item) => (
              <li key={item.name}>
                <span className="calc-name">{item.name}</span>
                <span className="calc-detail">{(item.members ?? []).map(captionOf).join(", ")}</span>
                <button
                  type="button"
                  className="button small"
                  onClick={() => setList((before) => before.filter((i) => i.name !== item.name))}
                >
                  Ungroup
                </button>
              </li>
            ))}
        </ul>
      )}
      <label className="dialog-field">
        <span className="field-label">New group&rsquo;s name</span>
        <input
          className="input"
          value={name}
          onChange={(e) => {
            setName(e.target.value);
            setProblem(null);
          }}
          maxLength={100}
          data-autofocus
        />
      </label>
      <input
        className="input"
        type="search"
        placeholder="Search items"
        aria-label="Search items"
        value={search}
        onChange={(e) => setSearch(e.target.value)}
      />
      <div className="calc-items">
        {error ? (
          <ErrorNote error={error} />
        ) : items === null ? (
          <p className="dialog-note">
            <Spinner label="Loading items" /> Loading items…
          </p>
        ) : (
          shown.map((member) => {
            const inGroup = groupOf.get(member.uniqueName);
            return (
              <label key={member.uniqueName} className="calc-item" title={inGroup ? `In ${inGroup}` : member.caption}>
                <input
                  type="checkbox"
                  disabled={inGroup !== undefined}
                  checked={chosen.has(member.uniqueName)}
                  onChange={(e) => {
                    setProblem(null);
                    setChosen((before) => {
                      const next = new Set(before);
                      if (e.target.checked) next.add(member.uniqueName);
                      else next.delete(member.uniqueName);
                      return next;
                    });
                  }}
                />
                <span className="calc-caption">{member.caption}</span>
                {inGroup && <span className="calc-detail">{inGroup}</span>}
              </label>
            );
          })
        )}
        {truncated && <p className="dialog-note">Only the first 5,000 items are listed.</p>}
      </div>
      {problem && <p className="fm-error">{problem}</p>}
    </Dialog>
  );
}

// ---------------------------------------------------------------------------
// Calculated Item
// ---------------------------------------------------------------------------

export function CalculatedItemDialog(
  props: CubeRef & {
    field: PlacedField;
    /** A value field to try the formula with. */
    value: ValueField | null;
    onApply: (items: CustomItem[]) => void;
    onCancel: () => void;
  },
) {
  const { field } = props;
  const { items, error } = useTopItems(props, field.hierarchy);
  const [list, setList] = useState<CustomItem[]>(field.customItems ?? []);
  const [editing, setEditing] = useState<string | null>(null);
  const others = list.filter((i) => i.name !== editing).map((i) => i.name);
  const [name, setName] = useState(() => nextName("Formula", list.map((i) => i.name)));
  const [formula, setFormula] = useState("");
  const [problem, setProblem] = useState<string | null>(null);
  const area = useRef<HTMLTextAreaElement>(null);
  const checker = useFormulaCheck(props);

  const candidate = (): CustomItem | null => {
    const checkedName = checkNewName(name, others, "calculated item");
    if (!checkedName.ok) {
      setProblem(checkedName.message);
      return null;
    }
    const checkedFormula = checkFormula(formula);
    if (!checkedFormula.ok) {
      setProblem(checkedFormula.message);
      return null;
    }
    return { name: checkedName.value, formula: checkedFormula.value };
  };
  const withCandidate = (item: CustomItem): CustomItem[] =>
    editing !== null ? list.map((i) => (i.name === editing ? item : i)) : [...list, item];
  const formulas = list.filter((i) => i.formula);

  return (
    <Dialog
      title={`Calculated Item (${field.caption})`}
      wide
      onCancel={props.onCancel}
      buttons={[
        { label: "Cancel", run: props.onCancel },
        {
          label: "OK",
          primary: true,
          run: () => {
            if (formula.trim() === "" && editing === null) {
              props.onApply(list);
              return;
            }
            const item = candidate();
            if (item) props.onApply(withCandidate(item));
          },
        },
      ]}
    >
      <p className="dialog-text">
        A new item of {field.caption}, worked out by the cube from an MDX formula, for example one item minus
        another. Pick items below to put their names in the formula.
      </p>
      {formulas.length > 0 && (
        <ul className="calc-list" aria-label="Calculated items">
          {formulas.map((item) => (
            <li key={item.name}>
              <span className="calc-name">{item.name}</span>
              <code className="calc-detail">{item.formula}</code>
              <button
                type="button"
                className="button small"
                onClick={() => {
                  setEditing(item.name);
                  setName(item.name);
                  setFormula(item.formula ?? "");
                  checker.reset();
                }}
              >
                Edit
              </button>
              <button
                type="button"
                className="button small"
                onClick={() => {
                  setList((before) => before.filter((i) => i.name !== item.name));
                  if (editing === item.name) {
                    setEditing(null);
                    setFormula("");
                  }
                }}
              >
                Delete
              </button>
            </li>
          ))}
        </ul>
      )}
      <label className="dialog-field">
        <span className="field-label">Name</span>
        <input
          className="input"
          value={name}
          onChange={(e) => {
            setName(e.target.value);
            setProblem(null);
          }}
          maxLength={100}
          data-autofocus
        />
      </label>
      <label className="dialog-field">
        <span className="field-label">Formula</span>
        <textarea
          ref={area}
          className="input formula"
          value={formula}
          rows={4}
          spellCheck={false}
          autoCapitalize="off"
          autoCorrect="off"
          onChange={(e) => {
            setFormula(e.target.value);
            setProblem(null);
            checker.reset();
          }}
        />
      </label>
      <div className="dialog-row">
        <select
          className="select"
          aria-label="Insert an item"
          value=""
          disabled={!items}
          onChange={(e) => {
            if (e.target.value) insertAtCaret(area.current, formula, e.target.value, setFormula);
          }}
        >
          <option value="">{error ? "Items could not be loaded" : items ? "Insert an item…" : "Loading items…"}</option>
          {(items ?? []).map((m) => (
            <option key={m.uniqueName} value={m.uniqueName}>
              {m.caption}
            </option>
          ))}
        </select>
        <button
          type="button"
          className="button small"
          disabled={checker.state.kind === "checking"}
          onClick={() => {
            const item = candidate();
            if (!item) return;
            const first = items && items.length > 0 ? [items[0].uniqueName] : [];
            const tried: PlacedField = {
              ...field,
              members: first,
              memberCaptions: first,
              customItems: [item],
              labelFilter: null,
              valueFilter: null,
              dateFilter: null,
              sort: null,
            };
            void checker.check({
              rows: [tried],
              columns: [],
              filters: [],
              values: props.value ? [props.value] : [],
              options: { ...DEFAULT_OPTIONS, nonEmpty: false },
            });
          }}
        >
          Check
        </button>
      </div>
      <CheckResult state={checker.state} />
      {problem && <p className="fm-error">{problem}</p>}
    </Dialog>
  );
}

// ---------------------------------------------------------------------------
// Calculated Measure
// ---------------------------------------------------------------------------

export function CalculatedMeasureDialog(
  props: CubeRef & {
    /** The cube's measures: for "Insert a measure", and names that are taken. */
    measures: Measure[];
    existing: CalculatedMeasure[];
    /** The measure to edit; null for a new one. */
    editing: CalculatedMeasure | null;
    onSave: (measure: CalculatedMeasure, previousName?: string) => void;
    onDelete: (name: string) => void;
    onCancel: () => void;
  },
) {
  const { editing } = props;
  const taken = [
    ...props.measures.map((m) => m.caption),
    ...props.existing.filter((m) => !editing || m.name !== editing.name).map((m) => m.name),
  ];
  const [name, setName] = useState(editing ? editing.name : nextName("Measure", taken));
  const [formula, setFormula] = useState(editing ? editing.formula : "");
  const [format, setFormat] = useState(editing?.formatString ?? "");
  const [problem, setProblem] = useState<string | null>(null);
  const area = useRef<HTMLTextAreaElement>(null);
  const checker = useFormulaCheck(props);
  const formats = CALCULATED_FORMATS.some((f) => f.format === format)
    ? CALCULATED_FORMATS
    : [...CALCULATED_FORMATS, { label: format, format }];

  const candidate = (): CalculatedMeasure | null => {
    const checkedName = checkNewName(name, taken, "calculated measure");
    if (!checkedName.ok) {
      setProblem(checkedName.message);
      return null;
    }
    const checkedFormula = checkFormula(formula);
    if (!checkedFormula.ok) {
      setProblem(checkedFormula.message);
      return null;
    }
    const measure: CalculatedMeasure = { name: checkedName.value, formula: checkedFormula.value };
    if (format !== "") measure.formatString = format;
    return measure;
  };

  const buttons = [
    ...(editing ? [{ label: "Delete", danger: true, run: () => props.onDelete(editing.name) }] : []),
    { label: "Cancel", run: props.onCancel },
    {
      label: "OK",
      primary: true,
      run: () => {
        const measure = candidate();
        if (measure) props.onSave(measure, editing ? editing.name : undefined);
      },
    },
  ];

  return (
    <Dialog title={editing ? "Edit Calculated Measure" : "New Calculated Measure"} wide onCancel={props.onCancel} buttons={buttons}>
      <p className="dialog-text">
        A measure worked out by the cube from an MDX formula, kept in this report. For example the revenue per ride:
        <code> [Measures].[Revenue] / [Measures].[Rides]</code>.
      </p>
      <label className="dialog-field">
        <span className="field-label">Name</span>
        <input
          className="input"
          value={name}
          onChange={(e) => {
            setName(e.target.value);
            setProblem(null);
          }}
          maxLength={100}
          data-autofocus
        />
      </label>
      <label className="dialog-field">
        <span className="field-label">Formula</span>
        <textarea
          ref={area}
          className="input formula"
          value={formula}
          rows={5}
          spellCheck={false}
          autoCapitalize="off"
          autoCorrect="off"
          onChange={(e) => {
            setFormula(e.target.value);
            setProblem(null);
            checker.reset();
          }}
        />
      </label>
      <div className="dialog-row">
        <select
          className="select"
          aria-label="Insert a measure"
          value=""
          onChange={(e) => {
            if (e.target.value) insertAtCaret(area.current, formula, e.target.value, setFormula);
          }}
        >
          <option value="">Insert a measure…</option>
          {props.measures.map((m) => (
            <option key={m.uniqueName} value={m.uniqueName}>
              {m.caption}
            </option>
          ))}
          {props.existing
            .filter((m) => !editing || m.name !== editing.name)
            .map((m) => (
              <option key={m.name} value={calculatedMeasureName(m.name)}>
                {m.name}
              </option>
            ))}
        </select>
        <button
          type="button"
          className="button small"
          disabled={checker.state.kind === "checking"}
          onClick={() => {
            const measure = candidate();
            if (!measure) return;
            const others = props.existing.filter((m) => !editing || m.name !== editing.name);
            void checker.check({
              rows: [],
              columns: [],
              filters: [],
              values: [{ measure: calculatedMeasureName(measure.name), caption: measure.name }],
              options: { ...DEFAULT_OPTIONS, nonEmpty: false },
              calculatedMeasures: [...others, measure],
            });
          }}
        >
          Check
        </button>
      </div>
      <label className="dialog-field">
        <span className="field-label">Number format</span>
        <select className="select" value={format} onChange={(e) => setFormat(e.target.value)}>
          {formats.map((f) => (
            <option key={f.format} value={f.format}>
              {f.label}
            </option>
          ))}
        </select>
      </label>
      <CheckResult state={checker.state} />
      {problem && <p className="fm-error">{problem}</p>}
    </Dialog>
  );
}
