// The four areas (Filters, Columns, Rows, Values) and the chips in them.
//
// A chip in Rows, Columns or Values is a button that opens a menu: move to another
// area, move up or down, sort and filter, remove; for a value field also Show Values
// As and Number Format. A small funnel marks a field that is filtered in any way
// (items, label filter, value filter), arrows a field with its own sort order, and "%"
// a value shown as a percentage.
//
// A chip in Filters says what it is filtering by ("Years: 2025") and opens the member
// picker when clicked, because choosing members is what a filter is for; its menu is
// behind the small button at its right.
//
// Dragging a chip or a field onto an area does the same moves faster, but nothing
// depends on it.

import { useState, type DragEvent } from "react";
import { NUMBER_FORMAT_CHOICES, SHOW_AS_CHOICES, isListedFormat } from "../state/fieldMenu";
import {
  AREA_LABELS,
  FIELD_AREAS,
  isFiltered,
  locate,
  selectionSummary,
  valueKey,
  type Area,
  type FieldArea,
  type LayoutAction,
} from "../state/layout";
import { valueFieldNames } from "../state/valueNames";
import type { Layout, PlacedField, ShowValuesAs, ValueField } from "../types";
import { cx } from "../util";
import { Icon } from "./common";
import { dragged, endDrag, startDrag } from "./drag";
import { Menu, type MenuEntry } from "./Menu";
import { ValueFieldSettings } from "./ValueFieldSettings";

interface Props {
  layout: Layout;
  dispatch: (action: LayoutAction) => void;
  /** Open the sort and filter panel (or, for a page filter, the member picker). */
  onFilter: (hierarchy: string) => void;
  /** Ask for a custom number format for a value field. */
  onCustomFormat: (measure: string) => void;
  /** Ask for the base field (and item) of a Show Values As kind that needs one. */
  onShowAsForm: (measure: string, showAs: ShowValuesAs) => void;
}

interface ChipInfo {
  name: string;
  caption: string;
  field: PlacedField | null;
  value: ValueField | null;
}

const ORDER: Area[] = ["filters", "columns", "rows", "values"];

/** A one-glance mark on a value chip for how its values are shown. */
const SHOW_AS_MARKS: Record<ShowValuesAs, string> = {
  normal: "",
  percentOfGrandTotal: "%",
  percentOfColumnTotal: "%",
  percentOfRowTotal: "%",
  percentOf: "%",
  percentOfParentRowTotal: "%",
  percentOfParentColumnTotal: "%",
  percentOfParentTotal: "%",
  differenceFrom: "Δ",
  percentDifferenceFrom: "Δ%",
  runningTotalIn: "Σ",
  percentRunningTotalIn: "Σ%",
  rankAscending: "#",
  rankDescending: "#",
  index: "ix",
};

function chipsOf(layout: Layout, area: Area): ChipInfo[] {
  if (area === "values") {
    // Copies of one measure are told apart by their key, and by "Rides 2" on the chip;
    // a field with a custom name shows that.
    const labels = valueFieldNames(layout.values);
    return layout.values.map((v, i) => ({ name: valueKey(layout.values, i), caption: labels[i], field: null, value: v }));
  }
  return layout[area].map((f) => ({ name: f.hierarchy, caption: f.caption, field: f, value: null }));
}

export function Areas({ layout, dispatch, onFilter, onCustomFormat, onShowAsForm }: Props) {
  const [menu, setMenu] = useState<{ area: Area; name: string; anchor: HTMLElement } | null>(null);
  const [hover, setHover] = useState<{ area: Area; index: number | null } | null>(null);
  /** The value field whose Value Field Settings are open, by key. */
  const [settings, setSettings] = useState<string | null>(null);

  const accepts = (area: Area): boolean => {
    const item = dragged();
    return item !== null && (item.kind === "measure") === (area === "values");
  };

  const over = (event: DragEvent, area: Area, index: number | null) => {
    if (!accepts(area)) return;
    event.preventDefault();
    event.stopPropagation();
    try {
      event.dataTransfer.dropEffect = "move";
    } catch {
      // not settable in every web view
    }
    if (!hover || hover.area !== area || hover.index !== index) setHover({ area, index });
  };

  const drop = (event: DragEvent, area: Area, index: number | null) => {
    const item = dragged();
    setHover(null);
    if (!item || !accepts(area)) return;
    event.preventDefault();
    event.stopPropagation();
    endDrag();
    const place = locate(layout, item.name);
    // `index` counts the chips as drawn; the reducer wants the position after the
    // dragged chip has been taken out of its old place.
    let at = index === null ? undefined : index;
    if (at !== undefined && place && place.area === area && place.index < at) at -= 1;
    if (place) {
      dispatch({ type: "move", name: item.name, area, index: at });
    } else if (item.kind === "measure") {
      dispatch({ type: "addMeasure", measure: { uniqueName: item.name, caption: item.caption }, index: at });
    } else {
      dispatch({
        type: "addHierarchy",
        hierarchy: { uniqueName: item.name, caption: item.caption, allMember: item.allMember },
        area: area as FieldArea,
        index: at,
      });
    }
  };

  const menuItems = (area: Area, name: string): MenuEntry[] => {
    const list = chipsOf(layout, area);
    const index = list.findIndex((c) => c.name === name);
    if (index < 0) return [];
    const chip = list[index];
    const items: MenuEntry[] = [
      { label: "Move up", run: () => dispatch({ type: "shift", name, delta: -1 }), disabled: index === 0 },
      {
        label: "Move down",
        run: () => dispatch({ type: "shift", name, delta: 1 }),
        disabled: index === list.length - 1,
      },
    ];
    if (area !== "values") {
      const isSet = Boolean(chip.field && chip.field.set);
      items.push("separator");
      for (const target of FIELD_AREAS) {
        if (target === area || (isSet && target === "filters")) continue;
        items.push({
          label: `Move to ${AREA_LABELS[target]}`,
          run: () => dispatch({ type: "move", name, area: target }),
        });
      }
      if (!isSet) {
        items.push("separator", {
          label: area === "filters" ? "Filter members…" : "Sort and Filter…",
          run: () => onFilter(name),
        });
      }
      if (chip.field && isFiltered(chip.field)) {
        items.push({ label: "Clear Filter", run: () => dispatch({ type: "clearFilters", hierarchy: name }) });
      }
    } else if (chip.value) {
      const value = chip.value;
      const hasAxisFields = layout.rows.length + layout.columns.length > 0;
      const showAs = value.showAs ?? "normal";
      const format = value.numberFormat ?? null;
      if (!value.kpi) {
        items.push("separator", {
          label: "Duplicate",
          run: () => dispatch({ type: "duplicateValue", key: name }),
        });
      }
      items.push(
        "separator",
        {
          label: "Show Values As",
          disabled: Boolean(value.kpi),
          submenu: [
            ...(hasAxisFields
              ? []
              : [{ note: "The items with … need a field in Rows or Columns." } as MenuEntry]),
            ...SHOW_AS_CHOICES.map(
              (choice): MenuEntry => ({
                label: choice.label,
                checked: choice.showAs === showAs,
                disabled: choice.needs !== "nothing" && !hasAxisFields,
                run: () =>
                  choice.needs === "nothing"
                    ? dispatch({ type: "setShowAs", measure: name, showAs: choice.showAs })
                    : onShowAsForm(name, choice.showAs),
              }),
            ),
          ],
        },
        {
          label: "Number Format",
          submenu: [
            ...NUMBER_FORMAT_CHOICES.map((choice) => ({
              label: choice.label,
              checked: choice.format === format,
              run: () => dispatch({ type: "setNumberFormat", measure: name, format: choice.format }),
            })),
            "separator" as const,
            {
              label: isListedFormat(format) ? "Custom…" : `Custom: ${format}`,
              checked: !isListedFormat(format),
              run: () => onCustomFormat(name),
            },
          ],
        },
        // Windows' name for where a value field is renamed (Custom Name).
        { label: "Value Field Settings…", run: () => setSettings(name) },
      );
    }
    items.push("separator", {
      label: "Remove",
      run: () => dispatch({ type: "remove", name }),
      danger: true,
    });
    return items;
  };

  return (
    <section className="areas" aria-label="Layout">
      {ORDER.map((area) => {
        const chips = chipsOf(layout, area);
        const target = hover !== null && hover.area === area;
        return (
          <div
            key={area}
            className={cx("area", target && hover.index === null && "drop-end")}
            onDragOver={(event) => over(event, area, null)}
            onDragLeave={(event) => {
              if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setHover(null);
            }}
            onDrop={(event) => drop(event, area, null)}
          >
            <h3 className="area-title" id={`area-${area}`}>
              {AREA_LABELS[area]}
            </h3>
            <ul className="chips" aria-labelledby={`area-${area}`}>
              {chips.map((chip, index) => {
                const field = chip.field;
                const restricted = field !== null && isFiltered(field);
                const sorted = field !== null && Boolean(field.sort);
                const percent = chip.value !== null && (chip.value.showAs ?? "normal") !== "normal";
                const marks = [
                  chip.value && chip.value.customName ? `Source Name: ${chip.value.caption}` : "",
                  field && field.members.length > 0 ? `items: ${field.memberCaptions.join(", ")}` : "",
                  field && field.labelFilter ? "label filter" : "",
                  field && field.valueFilter ? "value filter" : "",
                  sorted ? "sorted" : "",
                  percent ? SHOW_AS_CHOICES.find((c) => c.showAs === chip.value?.showAs)?.label ?? "" : "",
                  chip.value && chip.value.numberFormat ? `format ${chip.value.numberFormat}` : "",
                ].filter((text) => text !== "");
                const detail = marks.length > 0 ? `${chip.caption} (${marks.join("; ")})` : chip.caption;
                const open = menu !== null && menu.area === area && menu.name === chip.name;
                const pageFilter = area === "filters" && field !== null;
                return (
                  <li
                    key={chip.name}
                    className={cx(
                      "chip",
                      pageFilter && "split",
                      target && hover.index === index && "drop-before",
                    )}
                    draggable
                    onDragStart={(event) =>
                      startDrag(event, {
                        kind: area === "values" ? "measure" : "hierarchy",
                        name: chip.name,
                        caption: chip.caption,
                        allMember: chip.field ? chip.field.allMember : null,
                      })
                    }
                    onDragEnd={() => {
                      endDrag();
                      setHover(null);
                    }}
                    onDragOver={(event) => over(event, area, index)}
                    onDrop={(event) => drop(event, area, index)}
                  >
                    {pageFilter ? (
                      <>
                        <button
                          type="button"
                          className="chip-button chip-main"
                          title={`${detail}. Click to choose.`}
                          aria-label={`Filter ${chip.caption}: ${selectionSummary(field)}`}
                          onClick={() => {
                            setMenu(null);
                            onFilter(chip.name);
                          }}
                        >
                          <span className="chip-caption">
                            {chip.caption}: <span className="chip-value">{selectionSummary(field)}</span>
                          </span>
                        </button>
                        <button
                          type="button"
                          className={cx("chip-more", open && "open")}
                          aria-label={`${chip.caption}: move or remove`}
                          aria-haspopup="menu"
                          aria-expanded={open}
                          title="Move or remove"
                          onClick={(event) =>
                            setMenu(open ? null : { area, name: chip.name, anchor: event.currentTarget })
                          }
                        >
                          <Icon name="more" />
                        </button>
                      </>
                    ) : (
                      <button
                        type="button"
                        className={cx("chip-button", open && "open")}
                        aria-haspopup="menu"
                        aria-expanded={open}
                        aria-label={
                          chip.caption +
                          (restricted ? ", filtered" : "") +
                          (sorted ? ", sorted" : "") +
                          (percent ? ", shown as a percentage" : "")
                        }
                        title={detail}
                        onClick={(event) =>
                          setMenu(open ? null : { area, name: chip.name, anchor: event.currentTarget })
                        }
                      >
                        <span className="chip-caption">{chip.caption}</span>
                        {percent && chip.value && (
                          <span className="chip-mark" aria-hidden="true">
                            {SHOW_AS_MARKS[chip.value.showAs ?? "normal"]}
                          </span>
                        )}
                        {sorted && <Icon name="sort" className="chip-funnel" />}
                        {restricted && <Icon name="funnel" className="chip-funnel" />}
                        <Icon name="chevron-down" className="chip-caret" />
                      </button>
                    )}
                  </li>
                );
              })}
            </ul>
          </div>
        );
      })}
      {menu && (
        <Menu
          anchor={menu.anchor}
          label={`${AREA_LABELS[menu.area]} field`}
          items={menuItems(menu.area, menu.name)}
          onClose={() => setMenu(null)}
        />
      )}
      {settings !== null && (
        <ValueFieldSettings
          layout={layout}
          valueKey={settings}
          onCancel={() => setSettings(null)}
          onApply={(name) => {
            setSettings(null);
            dispatch({ type: "setCustomName", measure: settings, name });
          }}
        />
      )}
    </section>
  );
}
