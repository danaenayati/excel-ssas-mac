// The report's look, as on Excel's Design tab: the PivotTable Styles gallery (here a
// pop-up list with a small sample of the chosen style) and the four Style Options,
// laid out as Excel lays them out: Row Headers and Banded Rows, then Column Headers and
// Banded Columns. Rendered inside the options grid, whose two columns the checkboxes use.

import type { StyleOption } from "../state/layout";
import { PIVOT_STYLES, styleChoice, styleSwatch, type PivotStyle } from "../sheet/styles";
import type { LayoutOptions } from "../types";
import "./StyleOptions.css";

interface Props {
  options: LayoutOptions;
  onStyle: (name: string) => void;
  onOption: (option: StyleOption, value: boolean) => void;
}

/** Excel's words, in the order of its Design tab (two columns). */
const STYLE_OPTIONS: [StyleOption, string][] = [
  ["rowHeaders", "Row Headers"],
  ["bandedRows", "Banded Rows"],
  ["columnHeaders", "Column Headers"],
  ["bandedColumns", "Banded Columns"],
];

const GROUPS: PivotStyle["group"][] = ["Light", "Medium", "Dark"];

/** A tiny sample of the style: its header band over its body, framed in its line colour. */
function Swatch({ style }: { style: PivotStyle }) {
  const { header, body, line } = styleSwatch(style);
  return (
    <span
      className="style-swatch"
      aria-hidden="true"
      style={{ borderColor: line, background: `linear-gradient(${header} 0 40%, ${line} 40% 50%, ${body} 50%)` }}
    />
  );
}

export function StyleOptions({ options, onStyle, onOption }: Props) {
  const chosen = styleChoice(options);
  const current: Record<StyleOption, boolean> = {
    rowHeaders: chosen.rowHeaders,
    columnHeaders: chosen.columnHeaders,
    bandedRows: chosen.bandedRows,
    bandedColumns: chosen.bandedColumns,
  };
  return (
    <>
      <label className="row-field options-wide">
        <span className="row-label">Style</span>
        <Swatch style={chosen.style} />
        <select
          className="select"
          aria-label="PivotTable Style"
          value={chosen.style.name}
          onChange={(event) => onStyle(event.target.value)}
        >
          {PIVOT_STYLES.filter((s) => s.group === "None").map((s) => (
            <option key={s.name} value={s.name}>
              {s.name}
            </option>
          ))}
          {GROUPS.map((group) => (
            <optgroup key={group} label={group}>
              {PIVOT_STYLES.filter((s) => s.group === group).map((s) => (
                <option key={s.name} value={s.name}>
                  {s.name}
                </option>
              ))}
            </optgroup>
          ))}
        </select>
      </label>
      {STYLE_OPTIONS.map(([option, label]) => (
        <label key={option} className="check">
          <input type="checkbox" checked={current[option]} onChange={(event) => onOption(option, event.target.checked)} />
          {label}
        </label>
      ))}
    </>
  );
}
