// What each cube function does with its arguments. Excel passes them in the order of
// functions.json, then an invocation object that these functions do not use. Failures
// are thrown as FormulaError; functions.ts turns them into Excel errors.

import type { CellValue } from "../types";
import {
  captionArg,
  connectionArg,
  memberNames,
  propertyArg,
  rankArg,
  setArg,
  singleMember,
  valueError,
} from "./args";
import type { FormulaBatcher } from "./batch";

export type CubeFunction = (...args: unknown[]) => Promise<CellValue>;

/** An empty cube cell shows empty, as CUBEVALUE's does. */
function shown(value: CellValue): CellValue {
  return value === null || value === undefined ? "" : value;
}

function text(value: CellValue): string {
  return value === null || value === undefined ? "" : String(value);
}

/**
 * The functions by id (functions.json). Arguments are checked before anything is asked,
 * so a wrong one fails at once with #VALUE!.
 */
export function cubeFunctions(batcher: FormulaBatcher): Record<string, CubeFunction> {
  return {
    /** CUBE.VALUE(connection, [member_expression1], …): the cube's value for those members. */
    VALUE: async (connection, members) => {
      const conn = connectionArg(connection);
      const tuple = memberNames(members);
      return shown(await batcher.ask(conn, { kind: "value", tuple }));
    },

    /** CUBE.MEMBER(connection, member_expression, [caption]): the member's caption, once it is known to exist. */
    MEMBER: async (connection, member, caption) => {
      const conn = connectionArg(connection);
      const members = memberNames(member);
      if (members.length === 0) throw valueError("Give a member's unique name, such as [Date].[Year].&[2025].");
      const own = captionArg(caption);
      const found = await batcher.ask(conn, { kind: "member", members });
      return own !== null ? own : text(found);
    },

    /** CUBE.MEMBERPROPERTY(connection, member_expression, property). */
    MEMBERPROPERTY: async (connection, member, property) => {
      const conn = connectionArg(connection);
      const name = singleMember(member);
      const prop = propertyArg(property);
      return shown(await batcher.ask(conn, { kind: "property", member: name, property: prop }));
    },

    /** CUBE.RANKEDMEMBER(connection, set_expression, rank, [caption]): the rank-th item's caption. */
    RANKEDMEMBER: async (connection, set, rank, caption) => {
      const conn = connectionArg(connection);
      const expression = setArg(set);
      const position = rankArg(rank);
      const own = captionArg(caption);
      const found = await batcher.ask(conn, { kind: "ranked", set: expression, rank: position });
      return own !== null ? own : text(found);
    },

    /** CUBE.SETCOUNT(connection, set_expression): how many items the set has. */
    SETCOUNT: async (connection, set) => {
      const conn = connectionArg(connection);
      const expression = setArg(set);
      return shown(await batcher.ask(conn, { kind: "count", set: expression }));
    },
  };
}
