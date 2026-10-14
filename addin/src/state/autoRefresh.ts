// "Refresh every N minutes", as in Excel's Connection Properties: the pane refreshes
// every report of the workbook this often while it is open. The interval is a workbook
// setting (REFRESH_MINUTES), so it applies to every report and is kept with the file.
//
// The scheduler is plain logic with an injected clock: something calls `tick()` every
// few seconds (useAutoRefresh), and it decides whether a Refresh All is due and may
// start now. It never starts one while the pane is busy (a run, a question, Show
// Details), never two at once, and when Excel was busy (a cell being edited) it tries
// again at the next tick instead of waiting for the next interval.

import type { UiError } from "./runner";

/** The choices offered, in minutes. 0 is Off. */
export const REFRESH_EVERY_CHOICES = [0, 1, 5, 10, 15, 30, 60] as const;

/** How often the pane looks whether a refresh is due. */
export const TICK_MS = 15000;

/** Longest interval accepted from a workbook: a day. */
const MAX_MINUTES = 24 * 60;

/** What a Refresh All run by the scheduler reports back. */
export type AutoRefreshResult =
  /** It ran (whatever each report's outcome): the next one is a whole interval away. */
  | "done"
  /** Excel or the pane was busy and it stopped: try again at the next tick. */
  | "busy";

export type TickResult =
  | "off"
  /** Not due yet. */
  | "waiting"
  /** Due, but the pane is busy: the next tick looks again. */
  | "blocked"
  /** The previous automatic refresh is still running. */
  | "running"
  | "refreshed"
  /** Due, tried, Excel was busy: the next tick tries again. */
  | "retry";

export interface SchedulerDeps {
  /** Milliseconds, like Date.now. */
  now(): number;
  /** False while the pane is busy: a run, a question in a dialog, Show Details, edits waiting to run. */
  canRun(): boolean;
  refresh(): Promise<AutoRefreshResult>;
}

/** A stored interval as minutes: a whole number from 1 to a day, else 0 (Off). */
export function normalizeMinutes(value: unknown): number {
  if (typeof value !== "number" || !isFinite(value)) return 0;
  const minutes = Math.round(value);
  return minutes >= 1 && minutes <= MAX_MINUTES ? minutes : 0;
}

/** "1 minute", "15 minutes". */
export function minutesLabel(minutes: number): string {
  return minutes === 1 ? "1 minute" : `${minutes} minutes`;
}

/** The time of day for the status line, in the user's own format ("14:05" or "2:05 PM"). */
export function clockTime(at: number): string {
  return new Date(at).toLocaleTimeString(undefined, { hour: "numeric", minute: "2-digit" });
}

/**
 * True when an error says that Excel was busy rather than that something is wrong:
 * a cell is being edited (Office's InvalidOperationInCellEditMode, kept in the detail
 * by toSheetError).
 */
export function excelWasBusy(error: UiError): boolean {
  return (error.detail ?? "").indexOf("InvalidOperationInCellEditMode") >= 0;
}

export class AutoRefreshScheduler {
  private minutes = 0;
  /** When the countdown started: the last refresh, or when the interval was set. */
  private since: number;
  private running = false;

  constructor(private readonly deps: SchedulerDeps) {
    this.since = deps.now();
  }

  get interval(): number {
    return this.minutes;
  }

  /** True while an automatic refresh runs. */
  get busy(): boolean {
    return this.running;
  }

  /** Change the interval (0 is Off). A new interval counts from now. */
  setMinutes(value: number): void {
    const minutes = normalizeMinutes(value);
    if (minutes === this.minutes) return;
    this.minutes = minutes;
    this.since = this.deps.now();
  }

  /** Every report was just refreshed (by hand, or when the workbook opened): count from now. */
  noteRefreshed(): void {
    this.since = this.deps.now();
  }

  /** When the next refresh is due, or null when it is Off. */
  dueAt(): number | null {
    return this.minutes > 0 ? this.since + this.minutes * 60000 : null;
  }

  async tick(): Promise<TickResult> {
    const due = this.dueAt();
    if (due === null) return "off";
    if (this.running) return "running";
    if (this.deps.now() < due) return "waiting";
    if (!this.deps.canRun()) return "blocked";
    this.running = true;
    let result: AutoRefreshResult;
    try {
      result = await this.deps.refresh();
    } catch {
      // Whatever failed has been shown; do not try again until the next interval.
      result = "done";
    } finally {
      this.running = false;
    }
    if (result === "busy") return "retry";
    this.since = this.deps.now();
    return "refreshed";
  }
}
