// Excel calls a custom function once per cell. The batcher collects the calls of a short
// window, then asks the helper once per cube: one request for all CUBE.VALUE cells of
// that cube, one for all other cube functions (larger batches are split in requests of
// at most MAX_PER_REQUEST). Identical calls in one window are asked once. Nothing is
// kept after a window: the next recalculation asks the cube again, as it must.

import type { CellValue, FormulaLookupItem, FormulaResult } from "../types";
import {
  fromApiError,
  notAvailable,
  resolveConnection,
  toFormulaError,
  type FormulaError,
  type KnownConnection,
  type Target,
} from "./args";

/** How long calls are collected before they are sent. */
export const BATCH_MS = 30;
/** The helper's own limit (formulas.MAX_ITEMS). */
export const MAX_PER_REQUEST = 5000;
/** Keeps a request under the helper's 5 MB body limit, even when every character takes 3 bytes in UTF-8. */
export const MAX_REQUEST_CHARS = 1500000;

/** What one cell asks: a CUBE.VALUE tuple, or one of the lookups. */
export type FormulaRequest = { kind: "value"; tuple: string[] } | FormulaLookupItem;

export interface FormulaTransport {
  connections(): Promise<KnownConnection[]>;
  cells(connectionId: string, request: { catalog?: string; cube?: string; tuples: string[][] }): Promise<FormulaResult[]>;
  lookup(
    connectionId: string,
    request: { catalog?: string; cube?: string; items: FormulaLookupItem[] },
  ): Promise<FormulaResult[]>;
}

interface Waiter {
  resolve(value: CellValue): void;
  reject(error: FormulaError): void;
}

interface Entry {
  connection: string;
  request: FormulaRequest;
  waiters: Waiter[];
}

interface Group {
  target: Target;
  values: Entry[];
  lookups: Entry[];
}

function resolveAll(entry: Entry, value: CellValue): void {
  for (const waiter of entry.waiters) waiter.resolve(value);
}

function rejectAll(entry: Entry, error: FormulaError): void {
  for (const waiter of entry.waiters) waiter.reject(error);
}

/** Rough size of an entry in a request body. */
function size(entry: Entry): number {
  return JSON.stringify(entry.request).length + 2;
}

/** Split entries into requests of at most MAX_PER_REQUEST entries and MAX_REQUEST_CHARS. */
export function chunkEntries<T>(entries: T[], weight: (entry: T) => number): T[][] {
  const out: T[][] = [];
  let current: T[] = [];
  let chars = 0;
  for (const entry of entries) {
    const cost = weight(entry);
    if (current.length > 0 && (current.length >= MAX_PER_REQUEST || chars + cost > MAX_REQUEST_CHARS)) {
      out.push(current);
      current = [];
      chars = 0;
    }
    current.push(entry);
    chars += cost;
  }
  if (current.length > 0) out.push(current);
  return out;
}

export class FormulaBatcher {
  private pending = new Map<string, Entry>();
  private scheduled = false;

  /**
   * `schedule` runs the flush after the window; tests pass their own to flush by hand.
   */
  constructor(
    private readonly transport: FormulaTransport,
    private readonly schedule: (flush: () => void) => void = (flush) => {
      setTimeout(flush, BATCH_MS);
    },
  ) {}

  /** The helper's answer for one cell. Rejects with a FormulaError. */
  ask(connection: string, request: FormulaRequest): Promise<CellValue> {
    const key = JSON.stringify([connection.trim().toLowerCase(), request]);
    let entry = this.pending.get(key);
    if (!entry) {
      entry = { connection, request, waiters: [] };
      this.pending.set(key, entry);
    }
    const target = entry;
    const answer = new Promise<CellValue>((resolve, reject) => target.waiters.push({ resolve, reject }));
    if (!this.scheduled) {
      this.scheduled = true;
      this.schedule(() => void this.flush());
    }
    return answer;
  }

  /** Send everything collected so far. Resolves when every call has its answer. */
  async flush(): Promise<void> {
    const batch = Array.from(this.pending.values());
    this.pending.clear();
    this.scheduled = false;
    if (batch.length === 0) return;

    let known: KnownConnection[];
    try {
      known = await this.transport.connections();
    } catch (error) {
      const failure = toFormulaError(error);
      for (const entry of batch) rejectAll(entry, failure);
      return;
    }

    const groups = new Map<string, Group>();
    for (const entry of batch) {
      let target: Target;
      try {
        target = resolveConnection(entry.connection, known);
      } catch (error) {
        rejectAll(entry, toFormulaError(error));
        continue;
      }
      const key = [target.connectionId, target.catalog ?? "", target.cube ?? ""].join("\u0000").toLowerCase();
      let group = groups.get(key);
      if (!group) {
        group = { target, values: [], lookups: [] };
        groups.set(key, group);
      }
      (entry.request.kind === "value" ? group.values : group.lookups).push(entry);
    }

    const jobs: Promise<void>[] = [];
    groups.forEach((group) => {
      const { connectionId, catalog, cube } = group.target;
      const where = catalog !== null && cube !== null ? { catalog, cube } : {};
      for (const chunk of chunkEntries(group.values, size)) {
        const tuples = chunk.map((entry) => (entry.request as { tuple: string[] }).tuple);
        jobs.push(this.answer(chunk, () => this.transport.cells(connectionId, { ...where, tuples })));
      }
      for (const chunk of chunkEntries(group.lookups, size)) {
        const items = chunk.map((entry) => entry.request as FormulaLookupItem);
        jobs.push(this.answer(chunk, () => this.transport.lookup(connectionId, { ...where, items })));
      }
    });
    await Promise.all(jobs);
  }

  private async answer(chunk: Entry[], call: () => Promise<FormulaResult[]>): Promise<void> {
    let results: FormulaResult[];
    try {
      results = await call();
    } catch (error) {
      const failure = toFormulaError(error);
      for (const entry of chunk) rejectAll(entry, failure);
      return;
    }
    chunk.forEach((entry, i) => {
      const result = Array.isArray(results) ? results[i] : undefined;
      if (!result || typeof result !== "object") {
        rejectAll(entry, notAvailable("The helper sent no answer for this formula. Recalculate to try again."));
      } else if ("error" in result) {
        rejectAll(entry, fromApiError(result.error));
      } else {
        resolveAll(entry, result.value);
      }
    });
  }
}
