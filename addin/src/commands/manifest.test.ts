import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { COMMANDS, FUNCTION_NAMES } from "./actions";

const manifest = readFileSync(new URL("../../manifest/manifest.xml", import.meta.url), "utf8");

function all(pattern: RegExp): string[] {
  const out: string[] = [];
  const re = new RegExp(pattern.source, "g");
  let match: RegExpExecArray | null;
  while ((match = re.exec(manifest)) !== null) out.push(match[1]);
  return out;
}

describe("the manifest's right-click menu", () => {
  it("names every command's function, once each", () => {
    expect(all(/<FunctionName>([^<]+)<\/FunctionName>/).sort()).toEqual(COMMANDS.map((c) => FUNCTION_NAMES[c]).sort());
  });

  it("is on the cell menu and runs commands.html, which the build makes", () => {
    expect(manifest).toContain('<OfficeMenu id="ContextMenuCell">');
    expect(manifest).toContain('<FunctionFile resid="Commands.Url"/>');
    expect(manifest).toContain('<bt:Url id="Commands.Url" DefaultValue="https://localhost:47613/commands.html"/>');
    const built = readFileSync(new URL("../../vite.config.ts", import.meta.url), "utf8");
    expect(built).toContain('"commands.html"');
  });

  it("defines every string and URL it refers to", () => {
    const defined = new Set(all(/<bt:(?:String|Url|Image) id="([^"]+)"/));
    const used = all(/resid="([^"]+)"/);
    expect(used.filter((id) => !defined.has(id))).toEqual([]);
  });

  it("uses Windows' words", () => {
    const labels = all(/<bt:String id="Cell\.\w+\.Label" DefaultValue="([^"]+)"/);
    expect(labels).toEqual([
      "Show Details",
      "Refresh",
      "Refresh All",
      "Expand",
      "Collapse",
      "Expand Entire Field",
      "Collapse Entire Field",
    ]);
  });
});
