import { describe, expect, it } from "vitest";
import { ApiFailure, HelperDown } from "../apiErrors";
import { HELPER, createFormulaTransport } from "./client";

interface Call {
  url: string;
  init: { method: string; headers: Record<string, string>; body: string };
}

function fakeFetch(status: number, body: string) {
  const calls: Call[] = [];
  const fetchImpl = async (url: string, init: Call["init"]) => {
    calls.push({ url, init });
    return { ok: status >= 200 && status < 300, status, text: async () => body };
  };
  return { calls, fetchImpl };
}

describe("formula transport", () => {
  it("sends simple requests: POST, text/plain, the token in the body", async () => {
    const { calls, fetchImpl } = fakeFetch(200, JSON.stringify({ results: [{ value: 5 }] }));
    const transport = createFormulaTransport("tok-1", fetchImpl);
    const results = await transport.cells("c 1", { catalog: "BI", cube: "Model", tuples: [["[A].&[1]"]] });
    expect(results).toEqual([{ value: 5 }]);
    expect(calls).toHaveLength(1);
    expect(calls[0].url).toBe(`${HELPER}/api/connections/c%201/cells`);
    expect(calls[0].init.method).toBe("POST");
    // No custom header: anything else would need a preflight the runtime does not do.
    expect(calls[0].init.headers).toEqual({ "Content-Type": "text/plain;charset=UTF-8" });
    expect(JSON.parse(calls[0].init.body)).toEqual({ catalog: "BI", cube: "Model", tuples: [["[A].&[1]"]], token: "tok-1" });
  });

  it("reads the connection list and lookups", async () => {
    const list = fakeFetch(200, JSON.stringify({ connections: [{ id: "demo", name: "Sample cube" }] }));
    expect(await createFormulaTransport("t", list.fetchImpl).connections()).toEqual([{ id: "demo", name: "Sample cube" }]);
    expect(list.calls[0].url).toBe(`${HELPER}/api/formulas/connections`);
    const lookup = fakeFetch(200, JSON.stringify({ results: [{ value: "Tehran" }] }));
    expect(await createFormulaTransport("t", lookup.fetchImpl).lookup("demo", { items: [{ kind: "count", set: "x" }] })).toEqual([
      { value: "Tehran" },
    ]);
    expect(lookup.calls[0].url).toBe(`${HELPER}/api/connections/demo/lookup`);
  });

  it("turns the helper's error into an ApiFailure", async () => {
    const { fetchImpl } = fakeFetch(400, JSON.stringify({ error: { category: "bad-request", message: "Two databases." } }));
    const failure = await createFormulaTransport("t", fetchImpl)
      .cells("demo", { tuples: [] })
      .catch((error: unknown) => error);
    expect(failure).toBeInstanceOf(ApiFailure);
    expect((failure as ApiFailure).error).toEqual({ category: "bad-request", message: "Two databases." });
  });

  it("an answer that is not the helper's JSON is still an ApiFailure", async () => {
    for (const [status, body, category] of [
      [403, "Forbidden", "forbidden"],
      [413, "<html>too big</html>", "internal"],
      [200, "not json", "internal"],
    ] as const) {
      const failure = await createFormulaTransport("t", fakeFetch(status, body).fetchImpl)
        .connections()
        .catch((error: unknown) => error);
      expect(failure).toBeInstanceOf(ApiFailure);
      expect((failure as ApiFailure).error.category).toBe(category);
    }
  });

  it("a request that never reaches the helper is HelperDown", async () => {
    const failing = async () => {
      throw new TypeError("Load failed");
    };
    const failure = await createFormulaTransport("t", failing).connections().catch((error: unknown) => error);
    expect(failure).toBeInstanceOf(HelperDown);
  });
});
