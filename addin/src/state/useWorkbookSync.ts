// When does the pane re-read the report definitions (see sync.ts for what it does with
// them): when Excel says the workbook settings changed (ExcelApi 1.4), and when the pane
// gets the focus back, which covers an Excel that does not pass the event on from the
// commands' runtime.
//
// A read that overlaps one of the pane's own runs or saves may return the definitions
// as they were before it, so its result is dropped and the read is made again later.

import { useEffect, useRef } from "react";
import type { ReportDef, SheetHost } from "../sheet/host";

/** Settings events come in bursts (a save, then the document's own bookkeeping): wait for the last. */
export const SETTLE_MS = 300;
/** Look again this long after a read had to be dropped. */
export const RETRY_MS = 1000;
/** Give up after this many reads in a row that Excel refused (a cell being edited); the next event or focus starts again. */
const MAX_FAILURES = 5;

export interface WatchDeps {
  host: Pick<SheetHost, "loadReports" | "onSettingsChanged">;
  /** True while the pane is running a report; its definitions are in flux. */
  paneBusy(): boolean;
  /** When the pane last wrote a definition (`now()` after the write). */
  lastOwnWrite(): number;
  /** Take in what was read. */
  apply(loaded: ReportDef[]): void;
  now(): number;
}

export interface Watch {
  /** Read again now (the pane got the focus). */
  poke(): void;
  stop(): void;
}

/** Re-read the definitions after each settings change and each `poke`, one read at a time. */
export function watchWorkbook(deps: WatchDeps): Watch {
  let disposed = false;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let reading = false;
  let again = false;
  let failures = 0;

  const schedule = (delay: number) => {
    if (disposed) return;
    if (timer !== undefined) clearTimeout(timer);
    timer = setTimeout(() => void read(), delay);
  };

  const read = async () => {
    timer = undefined;
    if (disposed) return;
    if (reading) {
      again = true;
      return;
    }
    if (deps.paneBusy()) {
      schedule(RETRY_MS);
      return;
    }
    reading = true;
    const started = deps.now();
    let retry = false;
    try {
      const loaded = await deps.host.loadReports();
      failures = 0;
      if (disposed) return;
      if (deps.paneBusy() || deps.lastOwnWrite() >= started) retry = true;
      else deps.apply(loaded);
    } catch {
      failures++;
      retry = failures < MAX_FAILURES;
    } finally {
      reading = false;
    }
    if (again || retry) {
      again = false;
      schedule(retry ? RETRY_MS : SETTLE_MS);
    }
  };

  /** Something new happened: read after `delay`, with a fresh allowance of tries. */
  const wake = (delay: number) => {
    failures = 0;
    schedule(delay);
  };
  const unsubscribe = deps.host.onSettingsChanged ? deps.host.onSettingsChanged(() => wake(SETTLE_MS)) : () => undefined;
  return {
    // At once: a click on the pane (Refresh, say) follows its focus within moments, and
    // must find the definitions as the commands left them.
    poke: () => wake(0),
    stop: () => {
      disposed = true;
      if (timer !== undefined) clearTimeout(timer);
      unsubscribe();
    },
  };
}

interface Options {
  host: SheetHost;
  /** False until the pane has started. */
  active: boolean;
  paneBusy: () => boolean;
  /** When the pane last wrote a definition (Date.now() after the write). */
  lastOwnWrite: () => number;
  apply: (loaded: ReportDef[]) => void;
}

export function useWorkbookSync({ host, active, paneBusy, lastOwnWrite, apply }: Options): void {
  const latest = useRef({ paneBusy, lastOwnWrite, apply });
  latest.current = { paneBusy, lastOwnWrite, apply };

  useEffect(() => {
    // Without the event (the browser preview has it; an Excel without ExcelApi 1.4 does
    // not load the add-in at all) the pane works as before.
    if (!active || !host.onSettingsChanged) return;
    const watch = watchWorkbook({
      host,
      paneBusy: () => latest.current.paneBusy(),
      lastOwnWrite: () => latest.current.lastOwnWrite(),
      apply: (loaded) => latest.current.apply(loaded),
      now: () => Date.now(),
    });
    const onFocus = () => watch.poke();
    window.addEventListener("focus", onFocus);
    return () => {
      window.removeEventListener("focus", onFocus);
      watch.stop();
    };
  }, [active, host]);
}
