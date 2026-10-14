// Value Field Settings, as on Windows: the field's Source Name and a Custom Name box.
// The custom name heads the field's columns in the report and names its chip; an empty
// box gives back the cube's name. A name another field of the report already has is
// refused, as Excel refuses it.

import { useState } from "react";
import { customNameProblem, valueFieldNames } from "../state/valueNames";
import { valueIndex, valueLabels } from "../state/layout";
import type { Layout } from "../types";
import { Dialog } from "./common";

interface Props {
  layout: Layout;
  /** The value field's key (`valueKey`). */
  valueKey: string;
  /** The new name; null for the cube's own. */
  onApply: (name: string | null) => void;
  onCancel: () => void;
}

export function ValueFieldSettings({ layout, valueKey, onApply, onCancel }: Props) {
  const index = valueIndex(layout.values, valueKey);
  const value = layout.values[index];
  const own = index >= 0 ? valueLabels(layout.values)[index] : "";
  const [name, setName] = useState(() => (index >= 0 ? valueFieldNames(layout.values)[index] : ""));
  const [problem, setProblem] = useState<string | null>(null);
  if (!value) return null;

  const apply = () => {
    const wanted = name.trim();
    // The field's own header is no custom name: storing it would only pin it.
    if (wanted === "" || wanted === own) {
      onApply(null);
      return;
    }
    const refused = customNameProblem(layout, valueKey, wanted);
    if (refused) {
      setProblem(refused);
      return;
    }
    onApply(wanted);
  };

  return (
    <Dialog
      title="Value Field Settings"
      onCancel={onCancel}
      buttons={[
        { label: "Cancel", run: onCancel },
        { label: "OK", run: apply, primary: true },
      ]}
    >
      <p>
        <span className="dialog-text">Source Name:</span> {value.caption}
      </p>
      <label className="dialog-field">
        <span className="field-label">Custom Name:</span>
        <input
          className="input"
          value={name}
          maxLength={255}
          data-autofocus
          aria-invalid={problem !== null}
          onFocus={(event) => event.currentTarget.select()}
          onChange={(event) => {
            setName(event.target.value);
            setProblem(null);
          }}
          onKeyDown={(event) => {
            if (event.key === "Enter") {
              event.preventDefault();
              apply();
            }
          }}
        />
      </label>
      {problem ? (
        <p className="fm-error" role="alert">
          {problem}
        </p>
      ) : (
        <p className="dialog-note">Leave it empty to use the name from the cube.</p>
      )}
    </Dialog>
  );
}
