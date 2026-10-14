import { describe, expect, it } from "vitest";
import { ApiFailure, HelperDown } from "../apiErrors";
import {
  FormulaError,
  MAX_MESSAGE,
  captionArg,
  connectionArg,
  flattenValues,
  fromApiError,
  isUniqueName,
  memberNames,
  propertyArg,
  rankArg,
  resolveConnection,
  setArg,
  singleMember,
  splitTuple,
  toFormulaError,
} from "./args";

const KNOWN = [
  { id: "demo", name: "Sample cube" },
  { id: "c1", name: "Sales Cube" },
  { id: "c2", name: "BI/Old" },
];

function failure(run: () => unknown): FormulaError {
  try {
    run();
  } catch (error) {
    expect(error).toBeInstanceOf(FormulaError);
    return error as FormulaError;
  }
  throw new Error("did not throw");
}

describe("connection", () => {
  it("matches the name alone, without regard to case and spaces", () => {
    expect(resolveConnection("Sales Cube", KNOWN)).toEqual({ connectionId: "c1", catalog: null, cube: null });
    expect(resolveConnection("  sales cube ", KNOWN)).toEqual({ connectionId: "c1", catalog: null, cube: null });
  });

  it("splits a database and a cube from the right", () => {
    expect(resolveConnection("Sales Cube/BI/Model", KNOWN)).toEqual({ connectionId: "c1", catalog: "BI", cube: "Model" });
    expect(resolveConnection("Sales Cube / BI / Ride Sales", KNOWN)).toEqual({
      connectionId: "c1",
      catalog: "BI",
      cube: "Ride Sales",
    });
  });

  it("works for a connection name that contains a slash", () => {
    expect(resolveConnection("BI/Old", KNOWN)).toEqual({ connectionId: "c2", catalog: null, cube: null });
    expect(resolveConnection("BI/Old/Sales/Model", KNOWN)).toEqual({ connectionId: "c2", catalog: "Sales", cube: "Model" });
  });

  it("asks for both the database and the cube", () => {
    const half = failure(() => resolveConnection("Sales Cube/BI", KNOWN));
    expect(half.code).toBe("#VALUE!");
    expect(half.message).toContain('"Sales Cube/Database/Cube"');
    expect(failure(() => resolveConnection("Sales Cube//Model", KNOWN)).code).toBe("#VALUE!");
  });

  it("says plainly when there is no such connection", () => {
    const missing = failure(() => resolveConnection("Sales", KNOWN));
    expect(missing.code).toBe("#N/A");
    expect(missing.message).toContain('no connection called "Sales"');
    expect(failure(() => resolveConnection("Nope/BI/Model", KNOWN)).code).toBe("#N/A");
  });

  it("refuses a name two connections share", () => {
    const twice = [...KNOWN, { id: "c3", name: "sales cube" }];
    expect(failure(() => resolveConnection("Sales Cube", twice)).message).toContain("Several");
    expect(failure(() => resolveConnection("Sales Cube/BI/Model", twice)).message).toContain("Several");
  });

  it("takes the first argument as text", () => {
    expect(connectionArg(" Sample cube ")).toBe("Sample cube");
    expect(connectionArg(2025)).toBe("2025");
    expect(failure(() => connectionArg("")).code).toBe("#VALUE!");
    expect(failure(() => connectionArg(null)).code).toBe("#VALUE!");
  });
});

describe("members", () => {
  it("accepts unique names the way the helper does", () => {
    expect(isUniqueName("[Date].[Year].&[2025]")).toBe(true);
    expect(isUniqueName("[Date].[Calendar].[Month].&[2025]&[Q1]&[1]")).toBe(true);
    expect(isUniqueName("[Odd]]Name].[x]")).toBe(true);
    expect(isUniqueName("[Measures].[Rides]")).toBe(true);
    expect(isUniqueName("2025")).toBe(false);
    expect(isUniqueName("[Date].[Year].&[2025]; DROP")).toBe(false);
    expect(isUniqueName("[Date].[Year].Children")).toBe(false);
    expect(isUniqueName("[a" + "x".repeat(1000) + "]")).toBe(false);
  });

  it("flattens values, ranges and repeated arguments, skipping empty cells", () => {
    expect(flattenValues([[["a", ""]], [[null, "b"]], "c", [[" "]]])).toEqual(["a", "b", "c"]);
    expect(flattenValues(undefined)).toEqual([]);
    // The invocation object Excel passes last is not a value.
    expect(flattenValues({ address: "Sheet1!A1" })).toEqual([]);
  });

  it("splits a tuple written out in one cell", () => {
    expect(splitTuple("([Geo].[City].&[a,b], [Measures].[Rides])")).toEqual(["[Geo].[City].&[a,b]", "[Measures].[Rides]"]);
    expect(splitTuple("([A].&[x]]),y], [B].&[1])")).toEqual(["[A].&[x]]),y]", "[B].&[1]"]);
    expect(splitTuple(" [Date].[Year].&[2025] ")).toEqual(["[Date].[Year].&[2025]"]);
  });

  it("collects the member names of all arguments in order", () => {
    const args = [[["[Date].[Year].&[2025]"]], [["([Geo].[City].&[1], [Measures].[Rides])", ""]]];
    expect(memberNames(args)).toEqual(["[Date].[Year].&[2025]", "[Geo].[City].&[1]", "[Measures].[Rides]"]);
    expect(memberNames([])).toEqual([]);
  });

  it("explains that a caption is not a member", () => {
    const caption = failure(() => memberNames([[["2025"]]]));
    expect(caption.code).toBe("#VALUE!");
    expect(caption.message).toContain("CUBE.MEMBER cell shows a caption");
    expect(failure(() => memberNames([[[2025]]])).code).toBe("#VALUE!");
    expect(failure(() => memberNames([[["[Date].[Year].Children"]]])).message).not.toContain("caption");
  });

  it("allows at most 64 members", () => {
    const names = Array.from({ length: 65 }, (_, i) => `[D${i}].[H].&[1]`);
    expect(memberNames([[names.slice(0, 64)]])).toHaveLength(64);
    expect(failure(() => memberNames([[names]])).message).toContain("at most 64");
  });

  it("takes exactly one member where one is wanted", () => {
    expect(singleMember("[Geo].[City].&[1]")).toBe("[Geo].[City].&[1]");
    expect(failure(() => singleMember("([A].&[1], [B].&[2])")).code).toBe("#VALUE!");
    expect(failure(() => singleMember("")).code).toBe("#VALUE!");
  });
});

describe("other arguments", () => {
  it("rank", () => {
    expect(rankArg(1)).toBe(1);
    expect(rankArg(2.9)).toBe(2);
    expect(rankArg(" 3 ")).toBe(3);
    for (const bad of [0, -1, 10001, "x", null, NaN, Infinity]) expect(failure(() => rankArg(bad)).code).toBe("#VALUE!");
  });

  it("set, property and caption", () => {
    expect(setArg(" [Geo].[City].Children ")).toBe("[Geo].[City].Children");
    expect(failure(() => setArg("")).code).toBe("#VALUE!");
    expect(failure(() => setArg("x".repeat(2001))).code).toBe("#VALUE!");
    expect(propertyArg("Key")).toBe("Key");
    expect(failure(() => propertyArg(" ")).code).toBe("#VALUE!");
    expect(captionArg(null)).toBeNull();
    expect(captionArg("")).toBeNull();
    expect(captionArg("Total")).toBe("Total");
    expect(captionArg(2025)).toBe("2025");
    expect(captionArg({ address: "A1" })).toBeNull();
  });
});

describe("errors", () => {
  it("wrong input is #VALUE!, everything else #N/A, with the server's words", () => {
    expect(fromApiError({ category: "bad-request", message: "Bad." }).code).toBe("#VALUE!");
    const missing = fromApiError({ category: "not-found", message: "No such member.", detail: "member [x] not found" });
    expect(missing.code).toBe("#N/A");
    expect(missing.message).toBe("No such member. (member [x] not found)");
  });

  it("maps what a request can throw", () => {
    expect(toFormulaError(new HelperDown("refused")).message).toContain("helper is not running");
    expect(toFormulaError(new ApiFailure({ category: "forbidden", message: "x" }, 403)).message).toContain("Quit Excel");
    const server = toFormulaError(new ApiFailure({ category: "bad-request", message: "Has two databases." }, 400));
    expect(server.code).toBe("#VALUE!");
    expect(toFormulaError(new TypeError("boom")).code).toBe("#N/A");
  });

  it("keeps messages short enough for Excel", () => {
    expect(fromApiError({ category: "server", message: "x".repeat(400) }).message.length).toBe(MAX_MESSAGE);
  });
});
