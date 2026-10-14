// Keeping the pane in step with report definitions changed outside it. The right-click
// commands (src/commands/) run in a runtime of their own and save definitions straight
// to the workbook (Expand, Collapse, ...). The pane must then take those definitions as
// its own, or its next Refresh, Refresh All or "Refresh Every" would write its older
// copy back over them.

import type { ReportDef } from "../sheet/host";
import { parseReportDef, serializeReportDef } from "./report";

/** A definition as stored, so that two copies compare equal whatever their key order. */
export function definitionKey(def: ReportDef): string {
  const text = serializeReportDef(def);
  const parsed = parseReportDef(text, def.id);
  return parsed ? serializeReportDef(parsed) : text;
}

export interface SyncInput {
  /** The definitions just read from the workbook. */
  loaded: ReportDef[];
  /** The pane's list of saved definitions. */
  known: ReportDef[];
  /** The saved report being edited in the pane, as the pane has it; null for a new one. */
  current: ReportDef | null;
  /** The pane has a change of its own to the current report that is not on the sheet yet, or a run in progress. */
  pending: boolean;
}

export interface SyncResult {
  /** The pane's new list. */
  reports: ReportDef[];
  /** Reports changed outside the pane: what the pane remembers of their cells is stale. */
  changed: string[];
  /** Open this definition in place of the current one: it changed outside the pane. */
  reopen: ReportDef | null;
}

/**
 * What the pane should do with definitions read from the workbook. Only reports the
 * pane already knows are updated: the commands never add or delete one, and a read that
 * crossed the pane's own deletion must not bring a report back.
 */
export function syncReports({ loaded, known, current, pending }: SyncInput): SyncResult {
  const stored = new Map<string, ReportDef>();
  for (const def of loaded) stored.set(def.id, def);
  const currentKey = current ? definitionKey(current) : null;

  const changed: string[] = [];
  let updated = false;
  const reports = known.map((def) => {
    const fresh = stored.get(def.id);
    if (!fresh) return def;
    const freshKey = definitionKey(fresh);
    if (freshKey === definitionKey(def)) return def;
    updated = true;
    // The pane's own save, which its list has not caught up with yet.
    if (current && current.id === def.id && freshKey === currentKey) return fresh;
    changed.push(def.id);
    return fresh;
  });

  const reopen =
    current && !pending && changed.indexOf(current.id) >= 0 ? stored.get(current.id) ?? null : null;
  return { reports: updated ? reports : known, changed, reopen };
}
