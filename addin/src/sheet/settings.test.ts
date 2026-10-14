import { describe, expect, it } from "vitest";
import { emptyLayout } from "../state/layout";
import { newReportDef } from "../state/report";
import { REFRESH_MINUTES, REFRESH_ON_OPEN } from "./host";
import { PreviewHost } from "./preview";

const settle = () => new Promise<void>((resolve) => setTimeout(resolve, 0));

describe("PreviewHost workbook settings", () => {
  it("keeps a number, and forgets it when set to null", async () => {
    const host = new PreviewHost({ persist: false });
    expect(await host.getNumber(REFRESH_MINUTES)).toBeNull();
    await host.setNumber(REFRESH_MINUTES, 15);
    expect(await host.getNumber(REFRESH_MINUTES)).toBe(15);
    await host.setNumber(REFRESH_MINUTES, null);
    expect(await host.getNumber(REFRESH_MINUTES)).toBeNull();
    // A flag is not a number.
    await host.setFlag(REFRESH_ON_OPEN, true);
    expect(await host.getNumber(REFRESH_ON_OPEN)).toBeNull();
  });

  it("tells listeners after a definition, flag or number is saved, and stops when asked", async () => {
    const host = new PreviewHost({ persist: false });
    let heard = 0;
    const stop = host.onSettingsChanged(() => heard++);
    const def = { ...newReportDef("demo", "Sample", "Rides"), layout: emptyLayout() };

    await host.saveReport(def);
    expect(heard).toBe(0); // never during the change itself
    await settle();
    expect(heard).toBe(1);

    await host.setFlag(REFRESH_ON_OPEN, true);
    await host.setNumber(REFRESH_MINUTES, 5);
    await host.deleteReport(def.id);
    await settle();
    expect(heard).toBe(4);

    // Moving the selection is not a settings change.
    host.select(4, 4);
    await settle();
    expect(heard).toBe(4);

    stop();
    await host.setNumber(REFRESH_MINUTES, 10);
    await settle();
    expect(heard).toBe(4);
  });
});
