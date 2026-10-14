// "Number Format ▸ Custom…" for a value field: one text box, checked before it is used.

import { useState } from "react";
import { checkNumberFormat, MAX_NUMBER_FORMAT } from "../state/fieldMenu";
import { Dialog } from "./common";

interface Props {
  caption: string;
  initial: string;
  onApply: (format: string) => void;
  onCancel: () => void;
}

export function NumberFormatDialog({ caption, initial, onApply, onCancel }: Props) {
  const [text, setText] = useState(initial);
  const [error, setError] = useState<string | null>(null);
  const apply = () => {
    const checked = checkNumberFormat(text);
    if (checked.ok) onApply(checked.value);
    else setError(checked.message);
  };
  return (
    <Dialog
      title={`Number format: ${caption}`}
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
        <label className="fm-field">
          <span className="field-label">Excel number format</span>
          <input
            className="input"
            value={text}
            maxLength={MAX_NUMBER_FORMAT + 20}
            onChange={(event) => {
              setText(event.target.value);
              setError(null);
            }}
            placeholder="#,##0.00"
            autoFocus
            data-autofocus
            autoCapitalize="off"
            autoCorrect="off"
            spellCheck={false}
          />
        </label>
        <p className="hint">The same codes as in Excel&rsquo;s Format Cells, for example #,##0 or 0.0%.</p>
        {error && <p className="fm-error">{error}</p>}
      </form>
    </Dialog>
  );
}
