// Insert Chart and Delete Chart in the report menu: the pane's half of the charts that
// follow a report (sheet/charts.ts). The runner does the work; this says what happened.

import { CHART_KINDS, chartKindLabel, type ChartKind, type ReportChart } from "../sheet/charts";
import type { ReportDef } from "../sheet/host";
import type { ReportRunner, RunOutcome, UiError } from "../state/runner";
import type { MenuEntry } from "./Menu";

export interface ChartActions {
  runner: ReportRunner;
  /**
   * Run the report now, as Refresh does. Resolves with null when it cannot be asked (no
   * connection or cube).
   */
  refresh(def: ReportDef): Promise<RunOutcome | null>;
  /** Null: the pane's usual status for whatever is still running. */
  setBusy(text: string | null): void;
  setError(error: UiError | null): void;
  setNote(note: string | null): void;
  /** The definition as now saved, with its charts. */
  saved(def: ReportDef): void;
}

export async function insertChart(actions: ChartActions, target: ReportDef, type: ChartKind): Promise<void> {
  actions.setError(null);
  actions.setNote(null);
  const runtime = actions.runner.runtimes.get(target.id);
  if (!runtime || !runtime.chart) {
    // Not written since the workbook was opened: the chart needs the report's numbers.
    const outcome = await actions.refresh(target);
    if (outcome === null) {
      actions.setNote("Choose a connection for this report first.");
      return;
    }
    if (outcome.kind !== "written") return; // the run has said why
  }
  actions.setBusy("Adding the chart…");
  const result = await actions.runner.insertChart(target, type);
  actions.setBusy(null);
  if ("error" in result) {
    actions.setError(result.error);
    return;
  }
  actions.saved(result.def);
  actions.setNote(`${chartKindLabel(type)} chart added beside the report. It updates whenever the report does.`);
}

export async function deleteChart(actions: ChartActions, target: ReportDef, chart: ReportChart): Promise<void> {
  actions.setError(null);
  actions.setNote(null);
  const result = await actions.runner.deleteChart(target, chart.id);
  if ("error" in result) {
    actions.setError(result.error);
    return;
  }
  actions.saved(result.def);
  actions.setNote("Chart deleted.");
}

/**
 * "Insert Chart ▸" with Windows' PivotChart types, and "Delete Chart" when the report has
 * charts (a list to pick from when it has several, in the order they were added).
 * `busy`: something is running or the report has changes not yet applied; `canInsert`:
 * the report is on a sheet and has fields.
 */
export function chartMenuItems(
  actions: ChartActions,
  target: ReportDef,
  state: { busy: boolean; canInsert: boolean },
): MenuEntry[] {
  const items: MenuEntry[] = [
    {
      label: "Insert Chart",
      disabled: state.busy || !state.canInsert,
      submenu: CHART_KINDS.map(({ kind, label }) => ({
        label,
        run: () => void insertChart(actions, target, kind),
      })),
    },
  ];
  const charts = actions.runner.chartsOf(target);
  if (charts.length === 1) {
    items.push({ label: "Delete Chart", disabled: state.busy, run: () => void deleteChart(actions, target, charts[0]) });
  } else if (charts.length > 1) {
    items.push({
      label: "Delete Chart",
      disabled: state.busy,
      submenu: charts.map((chart, i) => ({
        label: `Chart ${i + 1}: ${chartKindLabel(chart.type)}`,
        run: () => void deleteChart(actions, target, chart),
      })),
    });
  }
  return items;
}
