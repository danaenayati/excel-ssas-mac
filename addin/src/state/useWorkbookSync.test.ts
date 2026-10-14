import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { ReportDef } from "../sheet/host";
import { PreviewHost } from "../sheet/preview";
import { layoutReducer } from "./layout";
import { newReportDef } from "./report";
import { RETRY_MS, SETTLE_MS, watchWorkbook } from "./useWorkbookSync";

let host: PreviewHost;
let applied: ReportDef[][];
let busy: boolean;
let ownWrite: number;
let reads: number;

function watch() {
  const original = host.loadReports.bind(host);
  host.loadReports = async () => {
    reads++;
    return original();
  };
  return watchWorkbook({
    host,
    paneBusy: () => busy,
    lastOwnWrite: () => ownWrite,
    apply: (loaded) => applied.push(loaded),
    now: () => Date.now(),
  });
}

function report(): ReportDef {
  return newReportDef("demo", "Sample", "Rides");
}

beforeEach(() => {
  vi.useFakeTimers();
  host = new PreviewHost({ persist: false });
  applied = [];
  busy = false;
  ownWrite = 0;
  reads = 0;
});

afterEach(() => {
  vi.useRealTimers();
});

describe("watchWorkbook", () => {
  it("reads the definitions once after a burst of settings changes", async () => {
    const w = watch();
    const def = report();
    await host.saveReport(def);
    await host.saveReport({ ...def, layout: layoutReducer(def.layout, { type: "setForm", form: "outline" }) });
    await vi.advanceTimersByTimeAsync(SETTLE_MS + 10);
    expect(reads).toBe(1);
    expect(applied).toHaveLength(1);
    expect(applied[0][0].layout.options.form).toBe("outline");
    w.stop();
  });

  it("reads again when poked (the pane got the focus)", async () => {
    const w = watch();
    w.poke();
    await vi.advanceTimersByTimeAsync(SETTLE_MS + 10);
    expect(applied).toEqual([[]]);
    w.stop();
  });

  it("waits while the pane is running a report, then reads", async () => {
    const w = watch();
    busy = true;
    await host.saveReport(report());
    await vi.advanceTimersByTimeAsync(SETTLE_MS + 10);
    expect(reads).toBe(0);
    await vi.advanceTimersByTimeAsync(RETRY_MS * 3);
    expect(reads).toBe(0);
    busy = false;
    await vi.advanceTimersByTimeAsync(RETRY_MS + 10);
    expect(reads).toBe(1);
    expect(applied).toHaveLength(1);
    w.stop();
  });

  it("drops a read that the pane's own write overtook, and reads again", async () => {
    const w = watch();
    await host.saveReport(report());
    // The pane writes just as the read starts: what was read may predate the write.
    ownWrite = Date.now() + SETTLE_MS;
    await vi.advanceTimersByTimeAsync(SETTLE_MS + 10);
    expect(reads).toBe(1);
    expect(applied).toHaveLength(0);
    await vi.advanceTimersByTimeAsync(RETRY_MS + 10);
    expect(reads).toBe(2);
    expect(applied).toHaveLength(1);
    w.stop();
  });

  it("tries a few times when Excel refuses the read, then waits for the next change", async () => {
    host.loadReports = async () => {
      reads++;
      throw new Error("InvalidOperationInCellEditMode");
    };
    const w = watchWorkbook({
      host,
      paneBusy: () => false,
      lastOwnWrite: () => 0,
      apply: (loaded) => applied.push(loaded),
      now: () => Date.now(),
    });
    w.poke();
    await vi.advanceTimersByTimeAsync(SETTLE_MS + RETRY_MS * 10);
    expect(reads).toBe(5);
    expect(applied).toHaveLength(0);
    // The next focus or change starts over.
    w.poke();
    await vi.advanceTimersByTimeAsync(SETTLE_MS + RETRY_MS * 10);
    expect(reads).toBe(10);
    w.stop();
  });

  it("stops listening when stopped", async () => {
    const w = watch();
    w.stop();
    await host.saveReport(report());
    w.poke();
    await vi.advanceTimersByTimeAsync(SETTLE_MS * 10);
    expect(reads).toBe(0);
  });
});
