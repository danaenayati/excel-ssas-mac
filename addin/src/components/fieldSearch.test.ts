// The field list's search: field names, and the names of the tables, measure groups and
// folders that hold them (as in Excel, a table found by its name is listed whole).

import { beforeAll, describe, expect, it } from "vitest";
import { createMockApi } from "../dev/mockApi";
import type { Fields } from "../types";
import { buildFieldTree, filterFieldTree } from "./FieldList";

let fields: Fields;

beforeAll(async () => {
  fields = await createMockApi({ delayMs: 0 }).fields("demo", "Sample", "Rides");
});

type Node = { type: string; label?: string; children?: Node[]; hierarchy?: { caption: string }; measure?: { caption: string } };

function names(entries: Node[]): string[] {
  const out: string[] = [];
  for (const e of entries) {
    if (e.type === "branch") out.push(`${e.label}: [${names(e.children ?? []).join(", ")}]`);
    else out.push(e.hierarchy?.caption ?? e.measure?.caption ?? e.type);
  }
  return out;
}

function search(text: string): string[] {
  return names(filterFieldTree(buildFieldTree(fields), text.trim().toLowerCase()) as Node[]);
}

describe("searching the field list", () => {
  it("finds a table by its name and lists all its fields", () => {
    // No field of Customer has "custom" in its name: the table's own name is what matches.
    const all = names(buildFieldTree(fields) as Node[]).find((n) => n.startsWith("Customer:"));
    expect(all).toBeDefined();
    expect(search("custom")).toEqual([all]);
    expect(search("CUSTOM")).toEqual([all]);
  });

  it("still finds fields by their own names, inside tables whose names do not match", () => {
    expect(search("segment")).toEqual(["Customer: [Segment]"]);
  });

  it("finds a measure group and a display folder by their names", () => {
    expect(search("trips")).toEqual(["Trips: [Distance km, Avg Distance km]"]);
    expect(search("money")).toEqual(["Rides: [Money: [Fare, Avg Fare]]"]);
  });

  it("finds nothing for a name that is nowhere", () => {
    expect(search("zzz")).toEqual([]);
  });
});
