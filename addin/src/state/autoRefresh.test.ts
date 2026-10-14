import { describe, expect, it } from "vitest";
import {
  AutoRefreshScheduler,
  REFRESH_EVERY_CHOICES,
  excelWasBusy,
  minutesLabel,
  normalizeMinutes,
  type AutoRefreshResult,
} from "./autoRefresh";

const MINUTE = 60000;

function setup(start = 1_000_000) {
  let now = start;
  let open = true;
  const results: AutoRefreshResult[] = [];
  const calls: number[] = [];
  let release: (() => void) | null = null;
  let hold = false;
  const scheduler = new AutoRefreshScheduler({
    now: () => now,
    canRun: () => open,
    refresh: async () => {
      calls.push(now);
      if (hold) await new Promise<void>((resolve) => (release = resolve));
      return results.shift() ?? "done";
    },
  });
  return {
    scheduler,
    calls,
    advance: (ms: number) => (now += ms),
    get now() {
      return now;
    },
    setOpen: (value: boolean) => (open = value),
    queue: (...r: AutoRefreshResult[]) => results.push(...r),
    hold: () => (hold = true),
    release: () => {
      hold = false;
      const go = release;
      release = null;
      if (go) go();
    },
  };
}

describe("normalizeMinutes", () => {
  it("keeps whole minutes from 1 to a day and turns anything else off", () => {
    expect(normalizeMinutes(5)).toBe(5);
    expect(normalizeMinutes(60)).toBe(60);
    expect(normalizeMinutes(1440)).toBe(1440);
    expect(normalizeMinutes(2.4)).toBe(2);
    expect(normalizeMinutes(0)).toBe(0);
    expect(normalizeMinutes(-5)).toBe(0);
    expect(normalizeMinutes(1441)).toBe(0);
    expect(normalizeMinutes(NaN)).toBe(0);
    expect(normalizeMinutes("5")).toBe(0);
    expect(normalizeMinutes(null)).toBe(0);
  });

  it("offers Excel-like choices, Off first", () => {
    expect(REFRESH_EVERY_CHOICES[0]).toBe(0);
    expect(REFRESH_EVERY_CHOICES.slice(1).every((m) => normalizeMinutes(m) === m)).toBe(true);
    expect(minutesLabel(1)).toBe("1 minute");
    expect(minutesLabel(15)).toBe("15 minutes");
  });
});

describe("AutoRefreshScheduler", () => {
  it("does nothing while it is off", async () => {
    const t = setup();
    expect(t.scheduler.dueAt()).toBeNull();
    t.advance(120 * MINUTE);
    expect(await t.scheduler.tick()).toBe("off");
    expect(t.calls).toEqual([]);
  });

  it("refreshes when the interval has passed, counting from when it was set", async () => {
    const t = setup();
    t.advance(10 * MINUTE); // the pane was open a while before the interval was chosen
    t.scheduler.setMinutes(5);
    expect(t.scheduler.dueAt()).toBe(t.now + 5 * MINUTE);
    t.advance(4 * MINUTE + 59000);
    expect(await t.scheduler.tick()).toBe("waiting");
    t.advance(1000);
    expect(await t.scheduler.tick()).toBe("refreshed");
    expect(t.calls).toHaveLength(1);
    // The next one is a whole interval after this one.
    expect(t.scheduler.dueAt()).toBe(t.now + 5 * MINUTE);
    expect(await t.scheduler.tick()).toBe("waiting");
  });

  it("counts the next interval from the end of a slow refresh", async () => {
    const t = setup();
    t.scheduler.setMinutes(1);
    t.advance(MINUTE);
    t.hold();
    const first = t.scheduler.tick();
    await Promise.resolve();
    t.advance(30000); // the cube takes half a minute
    t.release();
    expect(await first).toBe("refreshed");
    expect(t.scheduler.dueAt()).toBe(t.now + MINUTE);
  });

  it("waits while the pane is busy, and runs at the first tick after", async () => {
    const t = setup();
    t.scheduler.setMinutes(1);
    t.advance(MINUTE);
    t.setOpen(false);
    expect(await t.scheduler.tick()).toBe("blocked");
    t.advance(15000);
    expect(await t.scheduler.tick()).toBe("blocked");
    expect(t.calls).toEqual([]);
    t.setOpen(true);
    t.advance(15000);
    expect(await t.scheduler.tick()).toBe("refreshed");
    expect(t.calls).toHaveLength(1);
  });

  it("never runs two at once", async () => {
    const t = setup();
    t.scheduler.setMinutes(1);
    t.advance(MINUTE);
    t.hold();
    const first = t.scheduler.tick();
    expect(t.scheduler.busy).toBe(true);
    t.advance(15000);
    expect(await t.scheduler.tick()).toBe("running");
    t.advance(15000);
    expect(await t.scheduler.tick()).toBe("running");
    t.release();
    expect(await first).toBe("refreshed");
    expect(t.scheduler.busy).toBe(false);
    expect(t.calls).toHaveLength(1);
  });

  it("tries again at the next tick when Excel was busy", async () => {
    const t = setup();
    t.scheduler.setMinutes(10);
    t.advance(10 * MINUTE);
    t.queue("busy", "busy", "done");
    expect(await t.scheduler.tick()).toBe("retry");
    t.advance(15000);
    expect(await t.scheduler.tick()).toBe("retry");
    t.advance(15000);
    expect(await t.scheduler.tick()).toBe("refreshed");
    expect(t.calls).toHaveLength(3);
    expect(t.scheduler.dueAt()).toBe(t.now + 10 * MINUTE);
  });

  it("does not retry at once after a refresh that failed outright", async () => {
    const t = setup();
    let fail = true;
    const scheduler = new AutoRefreshScheduler({
      now: () => t.now,
      canRun: () => true,
      refresh: async () => {
        if (fail) throw new Error("helper gone");
        return "done";
      },
    });
    scheduler.setMinutes(5);
    t.advance(5 * MINUTE);
    expect(await scheduler.tick()).toBe("refreshed");
    fail = false;
    t.advance(15000);
    expect(await scheduler.tick()).toBe("waiting");
  });

  it("starts counting again after a Refresh All by hand", async () => {
    const t = setup();
    t.scheduler.setMinutes(5);
    t.advance(4 * MINUTE);
    t.scheduler.noteRefreshed();
    t.advance(MINUTE);
    expect(await t.scheduler.tick()).toBe("waiting");
    t.advance(4 * MINUTE);
    expect(await t.scheduler.tick()).toBe("refreshed");
  });

  it("restarts the countdown when the interval changes, and stops when turned off", async () => {
    const t = setup();
    t.scheduler.setMinutes(5);
    t.advance(4 * MINUTE);
    t.scheduler.setMinutes(15);
    expect(t.scheduler.interval).toBe(15);
    t.advance(2 * MINUTE);
    expect(await t.scheduler.tick()).toBe("waiting");
    // Setting the same interval again changes nothing.
    t.scheduler.setMinutes(15);
    t.advance(13 * MINUTE);
    expect(await t.scheduler.tick()).toBe("refreshed");
    t.scheduler.setMinutes(0);
    t.advance(60 * MINUTE);
    expect(await t.scheduler.tick()).toBe("off");
    expect(t.calls).toHaveLength(1);
  });

  it("refreshes once after the Mac wakes from a long sleep", async () => {
    const t = setup();
    t.scheduler.setMinutes(1);
    t.advance(8 * 60 * MINUTE);
    expect(await t.scheduler.tick()).toBe("refreshed");
    t.advance(15000);
    expect(await t.scheduler.tick()).toBe("waiting");
    expect(t.calls).toHaveLength(1);
  });
});

describe("excelWasBusy", () => {
  it("recognises a cell being edited, and nothing else", () => {
    expect(
      excelWasBusy({
        message: "Excel is waiting for you to finish editing a cell. Press Return, then try again.",
        detail: "InvalidOperationInCellEditMode · The operation is not allowed in cell edit mode.",
      }),
    ).toBe(true);
    expect(excelWasBusy({ message: "The helper is not running.", helperDown: true })).toBe(false);
    expect(excelWasBusy({ message: "Excel could not complete this.", detail: "GeneralException" })).toBe(false);
  });
});
