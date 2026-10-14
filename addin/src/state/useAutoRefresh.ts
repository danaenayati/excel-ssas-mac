// The pane's side of "Refresh every N minutes": one AutoRefreshScheduler for the life
// of the pane, ticked every TICK_MS while the interval is on. It runs only while the
// pane is open (Excel opens the pane with the workbook, see ensureAutoShow).

import { useEffect, useRef } from "react";
import { AutoRefreshScheduler, TICK_MS, type AutoRefreshResult } from "./autoRefresh";

interface Options {
  /** False until the pane has started. */
  active: boolean;
  /** The workbook's interval in minutes; 0 is Off. */
  minutes: number;
  /** False while the pane is busy. Read at each tick. */
  canRun: () => boolean;
  /** Refresh All, automatically. */
  refresh: () => Promise<AutoRefreshResult>;
}

export interface AutoRefresh {
  /** Every report was just refreshed some other way: the countdown starts again. */
  noteRefreshed(): void;
}

export function useAutoRefresh({ active, minutes, canRun, refresh }: Options): AutoRefresh {
  const latest = useRef({ canRun, refresh });
  latest.current = { canRun, refresh };

  const scheduler = useRef<AutoRefreshScheduler | null>(null);
  if (scheduler.current === null) {
    scheduler.current = new AutoRefreshScheduler({
      now: () => Date.now(),
      canRun: () => latest.current.canRun(),
      refresh: () => latest.current.refresh(),
    });
  }
  const current = scheduler.current;

  useEffect(() => {
    current.setMinutes(minutes);
  }, [current, minutes]);

  useEffect(() => {
    if (!active || minutes <= 0) return;
    const timer = window.setInterval(() => void current.tick(), TICK_MS);
    return () => window.clearInterval(timer);
  }, [active, current, minutes]);

  return { noteRefreshed: () => current.noteRefreshed() };
}
