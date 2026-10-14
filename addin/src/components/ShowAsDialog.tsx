// The small form behind the Show Values As items marked "…", like Excel's dialog: a
// base field from the fields in Rows and Columns, and for % Of, Difference From and
// % Difference From a base item: (previous), (next), or one of the field's top-level
// items, loaded from the cube.

import { useEffect, useState } from "react";
import type { Api } from "../api";
import { RELATIVE_ITEMS, checkShowAs, showAsChoice } from "../state/fieldMenu";
import type { Member, PlacedField, ShowValuesAs, ValueField } from "../types";
import { Dialog, Spinner } from "./common";

interface Props {
  api: Api;
  connectionId: string;
  catalog: string;
  cube: string;
  value: ValueField;
  showAs: ShowValuesAs;
  /** The fields now in Rows and Columns, in that order. */
  axisFields: PlacedField[];
  onApply: (choice: { showAs: ShowValuesAs; baseField: string | null; baseItem: string | null }) => void;
  onCancel: () => void;
}

export function ShowAsDialog({ api, connectionId, catalog, cube, value, showAs, axisFields, onApply, onCancel }: Props) {
  const choice = showAsChoice(showAs);
  const sameKind = (value.showAs ?? "normal") === showAs;
  const [baseField, setBaseField] = useState(
    sameKind && value.baseField && axisFields.some((f) => f.hierarchy === value.baseField)
      ? value.baseField
      : axisFields.length > 0
        ? axisFields[0].hierarchy
        : "",
  );
  const [baseItem, setBaseItem] = useState(sameKind && value.baseItem ? value.baseItem : "previous");
  const [items, setItems] = useState<Member[] | null>(null);
  const [itemsNote, setItemsNote] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const wantsItem = choice.needs === "fieldAndItem";

  // The base field's own items, top level only, as Excel lists them.
  useEffect(() => {
    if (!wantsItem || baseField === "") return;
    let cancelled = false;
    setItems(null);
    setItemsNote(null);
    api.members(connectionId, { catalog, cube, hierarchy: baseField }).then(
      (response) => {
        if (cancelled) return;
        setItems(response.members);
        if (response.truncated) setItemsNote(`Only the first ${response.members.length} items are listed.`);
      },
      () => {
        if (cancelled) return;
        setItems([]);
        setItemsNote("The items could not be loaded; (previous) and (next) still work.");
      },
    );
    return () => {
      cancelled = true;
    };
  }, [api, connectionId, catalog, cube, baseField, wantsItem]);

  const apply = () => {
    const checked = checkShowAs({
      showAs,
      baseField,
      baseItem,
      axisFields: axisFields.map((f) => f.hierarchy),
    });
    if (checked.ok) onApply(checked.value);
    else setError(checked.message);
  };

  return (
    <Dialog
      title={`Show Values As (${value.caption})`}
      onCancel={onCancel}
      buttons={[
        { label: "Cancel", run: onCancel },
        { label: "OK", run: apply, primary: true },
      ]}
    >
      <form
        className="fm-form plain"
        onSubmit={(event) => {
          event.preventDefault();
          apply();
        }}
      >
        <p className="fm-text">Calculation: {choice.label.replace("…", "")}</p>
        <label className="fm-field">
          <span className="field-label">Base field</span>
          <select
            className="select"
            value={baseField}
            onChange={(event) => {
              setBaseField(event.target.value);
              setBaseItem("previous");
              setError(null);
            }}
            data-autofocus
          >
            {axisFields.map((field) => (
              <option key={field.hierarchy} value={field.hierarchy}>
                {field.caption}
              </option>
            ))}
          </select>
        </label>
        {wantsItem && (
          <label className="fm-field">
            <span className="field-label">Base item</span>
            <select
              className="select"
              value={baseItem}
              onChange={(event) => {
                setBaseItem(event.target.value);
                setError(null);
              }}
            >
              {RELATIVE_ITEMS.map((item) => (
                <option key={item.item} value={item.item}>
                  {item.label}
                </option>
              ))}
              {/* Keep a stored item visible even before (or without) the list. */}
              {baseItem !== "previous" &&
                baseItem !== "next" &&
                !(items ?? []).some((m) => m.uniqueName === baseItem) && <option value={baseItem}>{baseItem}</option>}
              {(items ?? []).map((member) => (
                <option key={member.uniqueName} value={member.uniqueName}>
                  {member.caption}
                </option>
              ))}
            </select>
            {items === null && (
              <span className="hint">
                <Spinner label="Loading items" /> Loading items…
              </span>
            )}
            {itemsNote && <span className="hint">{itemsNote}</span>}
          </label>
        )}
        {error && <p className="fm-error">{error}</p>}
      </form>
    </Dialog>
  );
}
