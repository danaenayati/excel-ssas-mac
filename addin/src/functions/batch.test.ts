import { describe, expect, it } from "vitest";
import { ApiFailure, HelperDown } from "../apiErrors";
import type { FormulaLookupItem, FormulaResult } from "../types";
import { FormulaError } from "./args";
import { FormulaBatcher, MAX_PER_REQUEST, MAX_REQUEST_CHARS, chunkEntries, type FormulaTransport } from "./batch";

const KNOWN = [
  { id: "demo", name: "Sample cube" },
  { id: "c1", name: "Sales Cube" },
];

interface Sent {
  kind: "cells" | "lookup";
  connectionId: string;
  catalog?: string;
  cube?: string;
  count: number;
}

/** A transport that answers from a function, and records what was sent. */
function fakeTransport(answer: (tuple: string[]) => FormulaResult = (t) => ({ value: t.length })) {
  const sent: Sent[] = [];
  let connectionCalls = 0;
  const transport: FormulaTransport & { sent: Sent[]; connectionCalls(): number; fail?: unknown } = {
    sent,
    connectionCalls: () => connectionCalls,
    async connections() {
      connectionCalls++;
      if (transport.fail) throw transport.fail;
      return KNOWN;
    },
    async cells(connectionId, request) {
      sent.push({ kind: "cells", connectionId, catalog: request.catalog, cube: request.cube, count: request.tuples.length });
      return request.tuples.map(answer);
    },
    async lookup(connectionId, request) {
      sent.push({ kind: "lookup", connectionId, catalog: request.catalog, cube: request.cube, count: request.items.length });
      return request.items.map((item: FormulaLookupItem) =>
        item.kind === "member" ? { value: item.members[item.members.length - 1].slice(1, 4) } : { value: 7 },
      );
    },
  };
  return transport;
}

/** A batcher whose window is closed by hand. */
function manual(transport: FormulaTransport) {
  let scheduled = 0;
  const batcher = new FormulaBatcher(transport, () => {
    scheduled++;
  });
  return {
    batcher,
    async close() {
      // One window was opened by the first call of the batch, and only one.
      expect(scheduled).toBe(1);
      scheduled = 0;
      await batcher.flush();
    },
  };
}

async function outcome(promise: Promise<unknown>): Promise<unknown> {
  try {
    return await promise;
  } catch (error) {
    return error instanceof FormulaError ? `${error.code} ${error.message}` : error;
  }
}

describe("FormulaBatcher", () => {
  it("asks once per cube for all the cells of one window", async () => {
    const transport = fakeTransport();
    const { batcher, close } = manual(transport);
    const answers = [
      batcher.ask("Sample cube", { kind: "value", tuple: ["[A].&[1]"] }),
      batcher.ask("sample cube", { kind: "value", tuple: ["[A].&[1]", "[B].&[2]"] }),
      batcher.ask("Sales Cube/BI/Model", { kind: "value", tuple: [] }),
      batcher.ask("Sales Cube/bi/model", { kind: "member", members: ["[Geo].&[1]"] }),
      batcher.ask("Sales Cube", { kind: "count", set: "[Geo].Children" }),
    ];
    await close();
    expect(await Promise.all(answers)).toEqual([1, 2, 0, "Geo", 7]);
    expect(transport.connectionCalls()).toBe(1);
    expect(transport.sent).toEqual([
      { kind: "cells", connectionId: "demo", catalog: undefined, cube: undefined, count: 2 },
      { kind: "cells", connectionId: "c1", catalog: "BI", cube: "Model", count: 1 },
      { kind: "lookup", connectionId: "c1", catalog: "BI", cube: "Model", count: 1 },
      { kind: "lookup", connectionId: "c1", catalog: undefined, cube: undefined, count: 1 },
    ]);
  });

  it("asks identical calls of one window once", async () => {
    const transport = fakeTransport();
    const { batcher, close } = manual(transport);
    const a = batcher.ask("Sample cube", { kind: "value", tuple: ["[A].&[1]"] });
    const b = batcher.ask(" SAMPLE CUBE", { kind: "value", tuple: ["[A].&[1]"] });
    await close();
    expect([await a, await b]).toEqual([1, 1]);
    expect(transport.sent[0].count).toBe(1);
  });

  it("asks again in the next window: nothing is kept", async () => {
    const transport = fakeTransport();
    const { batcher, close } = manual(transport);
    const first = batcher.ask("Sample cube", { kind: "value", tuple: ["[A].&[1]"] });
    await close();
    const second = batcher.ask("Sample cube", { kind: "value", tuple: ["[A].&[1]"] });
    await close();
    expect([await first, await second]).toEqual([1, 1]);
    expect(transport.sent).toHaveLength(2);
    expect(transport.connectionCalls()).toBe(2);
  });

  it("gives each cell its own answer or error", async () => {
    const transport = fakeTransport((tuple) =>
      tuple[0] === "[Bad].&[1]"
        ? { error: { category: "not-found", message: "The cube has no such member." } }
        : tuple[0] === "[Odd]"
          ? { error: { category: "bad-request", message: "Not a name." } }
          : { value: null },
    );
    const { batcher, close } = manual(transport);
    const good = batcher.ask("Sample cube", { kind: "value", tuple: ["[A].&[1]"] });
    const bad = batcher.ask("Sample cube", { kind: "value", tuple: ["[Bad].&[1]"] });
    const odd = batcher.ask("Sample cube", { kind: "value", tuple: ["[Odd]"] });
    const lost = batcher.ask("Nope", { kind: "value", tuple: [] });
    await close();
    expect(await outcome(good)).toBeNull();
    expect(await outcome(bad)).toBe("#N/A The cube has no such member.");
    expect(await outcome(odd)).toBe("#VALUE! Not a name.");
    expect(await outcome(lost)).toMatch(/^#N\/A Cube Explorer has no connection called "Nope"/);
    expect(transport.sent).toHaveLength(1);
  });

  it("fails every cell of a request that failed, with its reason", async () => {
    const transport = fakeTransport();
    transport.cells = async () => {
      throw new ApiFailure({ category: "bad-request", message: '"Sales Cube" has more than one database.' }, 400);
    };
    const { batcher, close } = manual(transport);
    const a = batcher.ask("Sales Cube", { kind: "value", tuple: [] });
    const b = batcher.ask("Sales Cube", { kind: "value", tuple: ["[A].&[1]"] });
    const other = batcher.ask("Sales Cube", { kind: "member", members: ["[Geo].&[1]"] });
    await close();
    expect(await outcome(a)).toBe('#VALUE! "Sales Cube" has more than one database.');
    expect(await outcome(b)).toBe('#VALUE! "Sales Cube" has more than one database.');
    expect(await outcome(other)).toBe("Geo");
  });

  it("says the helper is not running when it cannot be reached", async () => {
    const transport = fakeTransport();
    transport.fail = new HelperDown("connection refused");
    const { batcher, close } = manual(transport);
    const a = batcher.ask("Sample cube", { kind: "value", tuple: [] });
    await close();
    expect(await outcome(a)).toMatch(/^#N\/A The Cube Explorer helper is not running/);
  });

  it("treats a short answer as no answer", async () => {
    const transport = fakeTransport();
    transport.cells = async () => [{ value: 1 }];
    const { batcher, close } = manual(transport);
    const a = batcher.ask("Sample cube", { kind: "value", tuple: ["[A].&[1]"] });
    const b = batcher.ask("Sample cube", { kind: "value", tuple: ["[A].&[2]"] });
    await close();
    expect(await outcome(a)).toBe(1);
    expect(await outcome(b)).toMatch(/^#N\/A The helper sent no answer/);
  });

  it("waits for the window before sending", async () => {
    const transport = fakeTransport();
    const batcher = new FormulaBatcher(transport);
    const a = batcher.ask("Sample cube", { kind: "value", tuple: ["[A].&[1]"] });
    expect(transport.sent).toHaveLength(0);
    const b = batcher.ask("Sample cube", { kind: "value", tuple: ["[A].&[2]"] });
    expect(await Promise.all([a, b])).toEqual([1, 1]);
    expect(transport.sent).toEqual([{ kind: "cells", connectionId: "demo", catalog: undefined, cube: undefined, count: 2 }]);
  });
});

describe("chunkEntries", () => {
  it("splits by count and by size", () => {
    const many = Array.from({ length: MAX_PER_REQUEST + 1 }, (_, i) => i);
    expect(chunkEntries(many, () => 1).map((c) => c.length)).toEqual([MAX_PER_REQUEST, 1]);
    expect(chunkEntries([1, 2, 3], () => MAX_REQUEST_CHARS / 2 + 1).map((c) => c.length)).toEqual([1, 1, 1]);
    expect(chunkEntries([1, 2, 3], () => MAX_REQUEST_CHARS / 3).map((c) => c.length)).toEqual([3]);
    expect(chunkEntries([], () => 1)).toEqual([]);
  });
});
