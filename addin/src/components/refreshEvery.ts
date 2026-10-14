// The report menu's "Refresh Every" choice: Excel's Connection Properties > "Refresh
// every N minutes", for every report of the workbook at once.

import { REFRESH_EVERY_CHOICES, minutesLabel } from "../state/autoRefresh";
import type { MenuEntry } from "./Menu";

/** The submenu, with the workbook's interval ticked. `minutes` 0 is Off. */
export function refreshEveryMenu(minutes: number, choose: (minutes: number) => void): MenuEntry {
  const choices: number[] = REFRESH_EVERY_CHOICES.slice();
  // An interval this menu does not offer (set in a newer version) is still shown as chosen.
  if (minutes > 0 && choices.indexOf(minutes) < 0) choices.push(minutes);
  return {
    label: minutes > 0 ? `Refresh Every: ${minutesLabel(minutes)}` : "Refresh Every",
    submenu: [
      { note: "Refreshes every report in this workbook while this panel is open. Saved with the workbook." },
      ...choices.map(
        (choice): MenuEntry => ({
          label: choice === 0 ? "Off" : minutesLabel(choice),
          checked: choice === minutes,
          run: () => choose(choice),
        }),
      ),
    ],
  };
}

/** What the pane says once the interval has been changed. */
export function refreshEveryNote(minutes: number): string {
  return minutes > 0
    ? `While this panel is open, every report in this workbook refreshes every ${minutesLabel(minutes)}.`
    : "Reports will no longer refresh on their own.";
}
