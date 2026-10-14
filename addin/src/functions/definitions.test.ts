import { describe, expect, it } from "vitest";
import metadata from "./functions.json";
import { FormulaError } from "./args";
import type { FormulaBatcher, FormulaRequest } from "./batch";
import { cubeFunctions } from "./definitions";

/** Records what each function asks, and answers from a table. */
function fakeBatcher(answer: (request: FormulaRequest) => unknown = () => 42) {
  const asked: { connection: string; request: FormulaRequest }[] = [];
  const batcher = {
    ask: async (connection: string, request: FormulaRequest) => {
      asked.push({ connection, request });
      const value = answer(request);
      if (value instanceof Error) throw value;
      return value;
    },
  } as unknown as FormulaBatcher;
  return { batcher, asked };
}

async function outcome(promise: Promise<unknown>): Promise<unknown> {
  try {
    return await promise;
  } catch (error) {
    return error instanceof FormulaError ? error.code : error;
  }
}

const INVOCATION = { address: "Sheet1!B2", functionName: "CUBE.VALUE" };

describe("cube functions", () => {
  it("CUBE.VALUE asks for the tuple of all its member arguments", async () => {
    const { batcher, asked } = fakeBatcher();
    const f = cubeFunctions(batcher);
    // As Excel passes a repeating range argument: a list of 2-D arrays, then the invocation.
    const members = [[["[Date].[Year].&[2025]"]], [["[Geo].[City].&[1]", ""], ["[Measures].[Rides]", null]]];
    expect(await f.VALUE("Sample cube", members, INVOCATION)).toBe(42);
    expect(asked).toEqual([
      {
        connection: "Sample cube",
        request: { kind: "value", tuple: ["[Date].[Year].&[2025]", "[Geo].[City].&[1]", "[Measures].[Rides]"] },
      },
    ]);
    expect(await f.VALUE("Sample cube", [], INVOCATION)).toBe(42);
    expect(asked[1].request).toEqual({ kind: "value", tuple: [] });
  });

  it("CUBE.VALUE shows an empty cube cell as empty", async () => {
    const f = cubeFunctions(fakeBatcher(() => null).batcher);
    expect(await f.VALUE("Sample cube", [[["[Measures].[Rides]"]]], INVOCATION)).toBe("");
  });

  it("checks arguments before asking anything", async () => {
    const { batcher, asked } = fakeBatcher();
    const f = cubeFunctions(batcher);
    expect(await outcome(f.VALUE("", [], INVOCATION))).toBe("#VALUE!");
    expect(await outcome(f.VALUE("Sample cube", [[["Tehran"]]], INVOCATION))).toBe("#VALUE!");
    expect(await outcome(f.MEMBER("Sample cube", [[""]], null, INVOCATION))).toBe("#VALUE!");
    expect(await outcome(f.RANKEDMEMBER("Sample cube", "[Geo].Children", 0, null, INVOCATION))).toBe("#VALUE!");
    expect(await outcome(f.SETCOUNT("Sample cube", "", INVOCATION))).toBe("#VALUE!");
    expect(await outcome(f.MEMBERPROPERTY("Sample cube", "[Geo].&[1]", "", INVOCATION))).toBe("#VALUE!");
    expect(asked).toEqual([]);
  });

  it("CUBE.MEMBER shows the cube's caption, or the one given", async () => {
    const { batcher, asked } = fakeBatcher(() => "Tehran");
    const f = cubeFunctions(batcher);
    expect(await f.MEMBER("Sample cube", [["[Geo].[City].&[1]"]], null, INVOCATION)).toBe("Tehran");
    expect(await f.MEMBER("Sample cube", [["[Geo].[City].&[1]"]], "Capital", INVOCATION)).toBe("Capital");
    expect(await f.MEMBER("Sample cube", [["[Date].[Year].&[2025]", "[Geo].[City].&[1]"]], null, INVOCATION)).toBe("Tehran");
    expect(asked[2].request).toEqual({ kind: "member", members: ["[Date].[Year].&[2025]", "[Geo].[City].&[1]"] });
  });

  it("CUBE.MEMBER with a caption still fails when the member does not exist", async () => {
    const f = cubeFunctions(fakeBatcher(() => new FormulaError("#N/A", "No such member.")).batcher);
    expect(await outcome(f.MEMBER("Sample cube", [["[Geo].[City].&[9]"]], "Capital", INVOCATION))).toBe("#N/A");
  });

  it("the other functions ask for what they name", async () => {
    const { batcher, asked } = fakeBatcher((r) => (r.kind === "count" ? 9 : r.kind === "ranked" ? "Shiraz" : "1"));
    const f = cubeFunctions(batcher);
    expect(await f.MEMBERPROPERTY("Sample cube", "[Geo].[City].&[1]", "Key", INVOCATION)).toBe("1");
    expect(await f.RANKEDMEMBER("Sample cube", "[Geo].[City].Children", 2, null, INVOCATION)).toBe("Shiraz");
    expect(await f.RANKEDMEMBER("Sample cube", "[Geo].[City].Children", 2.7, "Second", INVOCATION)).toBe("Second");
    expect(await f.SETCOUNT("Sample cube", "[Geo].[City].Children", INVOCATION)).toBe(9);
    expect(asked.map((a) => a.request)).toEqual([
      { kind: "property", member: "[Geo].[City].&[1]", property: "Key" },
      { kind: "ranked", set: "[Geo].[City].Children", rank: 2 },
      { kind: "ranked", set: "[Geo].[City].Children", rank: 2 },
      { kind: "count", set: "[Geo].[City].Children" },
    ]);
  });
});

describe("functions.json", () => {
  const functions = cubeFunctions(fakeBatcher().batcher);

  it("describes exactly the functions that are registered", () => {
    expect(metadata.functions.map((f) => f.id).sort()).toEqual(Object.keys(functions).sort());
    for (const f of metadata.functions) {
      expect(f.id).toMatch(/^[A-Z0-9.]+$/);
      expect(f.name).toBe(f.id);
    }
  });

  it("has a repeating member list last, and optional captions last", () => {
    for (const f of metadata.functions) {
      const params = f.parameters as { name: string; repeating?: boolean; optional?: boolean; dimensionality?: string }[];
      expect(params[0].name).toBe("connection");
      params.forEach((p, i) => {
        if (p.repeating) {
          expect(i).toBe(params.length - 1);
          expect(p.dimensionality).toBe("matrix");
        }
        if (p.optional) expect(params.slice(i).every((q) => q.optional || q.repeating)).toBe(true);
      });
    }
    const value = metadata.functions.find((f) => f.id === "VALUE")!;
    expect(value.parameters[1]).toMatchObject({ name: "member_expression1", repeating: true });
  });
});
