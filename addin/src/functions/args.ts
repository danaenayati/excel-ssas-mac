// The cube functions' arguments, as Excel passes them, turned into what the helper is
// asked; and the errors a formula shows. Pure, so it is tested without Excel.
//
// Arguments mirror Windows Excel's CUBEVALUE and its siblings, with one difference: a
// cell can only hold a value, so a CUBE.MEMBER cell holds a caption, not a member.
// CUBE.VALUE therefore takes members by unique name ("[Date].[Year].&[2025]"), as text
// or from a cell that holds that text, not from a CUBE.MEMBER cell.

import { ApiFailure, HelperDown } from "../apiErrors";
import type { ApiError } from "../types";

export type ExcelErrorCode = "#VALUE!" | "#N/A";

/** A failure the formula shows as an Excel error, with a sentence in its tooltip. */
export class FormulaError extends Error {
  readonly code: ExcelErrorCode;
  constructor(code: ExcelErrorCode, message: string) {
    super(clip(message));
    this.name = "FormulaError";
    this.code = code;
  }
}

/** A wrong argument: #VALUE!, as Excel's own functions show for one. */
export function valueError(message: string): FormulaError {
  return new FormulaError("#VALUE!", message);
}

/** The cube could not give an answer: #N/A, as CUBEVALUE shows for a missing member. */
export function notAvailable(message: string): FormulaError {
  return new FormulaError("#N/A", message);
}

/** Longest message passed to Excel; its error tooltips are short. */
export const MAX_MESSAGE = 250;
export const MAX_TUPLE_MEMBERS = 64;
export const MAX_RANK = 10000;
export const MAX_SET = 2000;
export const MAX_PROPERTY = 200;

function clip(text: string): string {
  return text.length <= MAX_MESSAGE ? text : text.slice(0, MAX_MESSAGE - 1) + "…";
}

function shorten(text: string, length = 60): string {
  return text.length <= length ? text : text.slice(0, length - 1) + "…";
}

// -- errors from the helper ----------------------------------------------------------

/** An answer's or a request's error: wrong input is #VALUE!, everything else #N/A. */
export function fromApiError(error: ApiError): FormulaError {
  const detail = error.detail ? ` (${shorten(error.detail, 160)})` : "";
  return error.category === "bad-request"
    ? valueError(error.message + detail)
    : notAvailable(error.message + detail);
}

/** Whatever a request threw, as the error its formulas show. */
export function toFormulaError(error: unknown): FormulaError {
  if (error instanceof FormulaError) return error;
  if (error instanceof HelperDown) {
    return notAvailable(
      "The Cube Explorer helper is not running, or Excel does not trust it. Open the Cube Explorer panel to see what to do.",
    );
  }
  if (error instanceof ApiFailure) {
    if (error.error.category === "forbidden") {
      return notAvailable(
        "The Cube Explorer helper did not accept this formula. Quit Excel and open it again; if that does not help, run install.sh again.",
      );
    }
    return fromApiError(error.error);
  }
  return notAvailable("Something went wrong while asking the cube. Recalculate to try again.");
}

// -- the connection ------------------------------------------------------------------

export interface KnownConnection {
  id: string;
  name: string;
}

/** What a formula's first argument names: a connection, and its database and cube if given. */
export interface Target {
  connectionId: string;
  /** Null: the connection's only database (the helper finds it, or says to name one). */
  catalog: string | null;
  cube: string | null;
}

/** The first argument as text. */
export function connectionArg(value: unknown): string {
  const text = typeof value === "number" ? String(value) : typeof value === "string" ? value.trim() : "";
  if (text === "") {
    throw valueError('The first argument is the Cube Explorer connection, such as "Sample cube".');
  }
  return text;
}

function sameName(a: string, b: string): boolean {
  return a.trim().toLowerCase() === b.trim().toLowerCase();
}

/**
 * Match "Name" or "Name/Database/Cube" against the connections in Cube Explorer. The
 * name is tried whole first, then the last two parts are split off from the right, so a
 * connection name that itself contains "/" still works. Case does not matter.
 */
export function resolveConnection(text: string, known: KnownConnection[]): Target {
  const wanted = text.trim();
  const named = (name: string) => known.filter((c) => sameName(c.name, name));
  const several = (name: string) =>
    notAvailable(`Several Cube Explorer connections are called "${shorten(name)}". Rename one of them in the Cube Explorer panel.`);

  const whole = named(wanted);
  if (whole.length === 1) return { connectionId: whole[0].id, catalog: null, cube: null };
  if (whole.length > 1) throw several(wanted);

  const parts = wanted.split("/");
  if (parts.length >= 3) {
    const name = parts.slice(0, -2).join("/").trim();
    const catalog = parts[parts.length - 2].trim();
    const cube = parts[parts.length - 1].trim();
    const found = named(name);
    if (found.length > 1) throw several(name);
    if (found.length === 1) {
      if (catalog === "" || cube === "") {
        throw valueError(`Write the connection as "${shorten(name)}/Database/Cube", with both names.`);
      }
      return { connectionId: found[0].id, catalog, cube };
    }
  }
  if (parts.length >= 2) {
    const name = parts.slice(0, -1).join("/").trim();
    if (named(name).length > 0) {
      throw valueError(`Write the connection as "${shorten(name)}/Database/Cube": after the name, the database and then the cube.`);
    }
  }
  const example = known.length > 0 ? known[known.length - 1].name : "Sample cube";
  throw notAvailable(
    `Cube Explorer has no connection called "${shorten(wanted)}". Use a name from its panel, such as "${shorten(example, 40)}".`,
  );
}

// -- members -------------------------------------------------------------------------

const PART = String.raw`\[(?:[^\]\x00-\x1f]|\]\])*\]`;
const UNIQUE_NAME = new RegExp(`^${PART}(?:\\.${PART}|\\.?&${PART})*$`);
const MAX_NAME = 1000;

/** True for a unique name in the shape the cube writes one (the helper's own check). */
export function isUniqueName(text: string): boolean {
  return text.length > 0 && text.length <= MAX_NAME && UNIQUE_NAME.test(text);
}

/**
 * Every value in what Excel passes: one value, a range (2-D array), or a list of those.
 * Empty cells are left out, and so is anything that is not a plain value (such as the
 * invocation object Excel passes after the last argument).
 */
export function flattenValues(value: unknown, out: unknown[] = []): unknown[] {
  if (Array.isArray(value)) {
    for (const item of value) flattenValues(item, out);
  } else if (typeof value === "number" || typeof value === "boolean" || (typeof value === "string" && value.trim() !== "")) {
    out.push(value);
  }
  return out;
}

/** Split on the commas that are outside brackets. */
function splitOutsideBrackets(text: string): string[] {
  const parts: string[] = [];
  let start = 0;
  let inside = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text.charAt(i);
    if (inside) {
      if (ch === "]") {
        if (text.charAt(i + 1) === "]") i++;
        else inside = false;
      }
    } else if (ch === "[") {
      inside = true;
    } else if (ch === ",") {
      parts.push(text.slice(start, i));
      start = i + 1;
    }
  }
  parts.push(text.slice(start));
  return parts.map((p) => p.trim()).filter((p) => p !== "");
}

/** "([A].&[1], [B].&[2])", a tuple written out as Windows Excel allows, is its members; other text is itself. */
export function splitTuple(text: string): string[] {
  const trimmed = text.trim();
  if (trimmed.length >= 2 && trimmed.charAt(0) === "(" && trimmed.charAt(trimmed.length - 1) === ")") {
    return splitOutsideBrackets(trimmed.slice(1, -1));
  }
  return [trimmed];
}

function notAName(value: unknown): FormulaError {
  const text = typeof value === "string" ? value.trim() : String(value);
  const caption = text.indexOf("[") < 0;
  return valueError(
    `"${shorten(text, 50)}" is not a member's unique name, such as [Date].[Year].&[2025].` +
      (caption ? " A CUBE.MEMBER cell shows a caption: use the unique name it was given instead." : ""),
  );
}

/** The member unique names in member arguments, in order. */
export function memberNames(values: unknown): string[] {
  const names: string[] = [];
  for (const value of flattenValues(values)) {
    if (typeof value !== "string") throw notAName(value);
    for (const part of splitTuple(value)) {
      if (!isUniqueName(part)) throw notAName(part);
      names.push(part);
    }
  }
  if (names.length > MAX_TUPLE_MEMBERS) {
    throw valueError(`One formula can name at most ${MAX_TUPLE_MEMBERS} members.`);
  }
  return names;
}

/** One member, for CUBE.MEMBERPROPERTY. */
export function singleMember(value: unknown): string {
  const names = memberNames(value);
  if (names.length !== 1) throw valueError("Give one member's unique name, such as [Geography].[City].&[Tehran].");
  return names[0];
}

// -- the other arguments -------------------------------------------------------------

/** An MDX set expression, such as "[Geography].[City].Children" or "TopCount(…)". */
export function setArg(value: unknown): string {
  const text = typeof value === "string" ? value.trim() : "";
  if (text === "") throw valueError("Give the set as an MDX expression, such as [Geography].[City].Children.");
  if (text.length > MAX_SET) throw valueError(`The set expression is longer than ${MAX_SET} characters.`);
  return text;
}

export function propertyArg(value: unknown): string {
  const text = typeof value === "string" ? value.trim() : "";
  if (text === "" || text.length > MAX_PROPERTY) {
    throw valueError("Give the name of a member property, such as Key or Caption.");
  }
  return text;
}

/** A rank from 1; decimals are dropped, as Excel's own functions do. */
export function rankArg(value: unknown): number {
  const n = typeof value === "number" ? value : typeof value === "string" && value.trim() !== "" ? Number(value) : NaN;
  if (!isFinite(n)) throw valueError("The rank must be a number, such as 1 for the first item.");
  const rank = Math.trunc(n);
  if (rank < 1 || rank > MAX_RANK) throw valueError(`The rank must be from 1 to ${MAX_RANK.toLocaleString("en-US")}.`);
  return rank;
}

/** An optional caption: null when not given. */
export function captionArg(value: unknown): string | null {
  if (typeof value === "string") return value === "" ? null : value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  return null;
}
