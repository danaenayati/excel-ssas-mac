// Entry point of the cube functions: CUBE.VALUE, CUBE.MEMBER, CUBE.MEMBERPROPERTY,
// CUBE.RANKEDMEMBER and CUBE.SETCOUNT ("CUBE" is the namespace in the manifest).
//
// Built on its own into dist/functions.js, a classic script (vite.config.ts, mode
// "functions"), because Excel runs custom functions in a JavaScript-only runtime of
// their own: no page, no modules, no Office.js, and nothing shared with the pane.
// Excel loads the script with the metadata in functions.json; on the web it loads
// functions.html, which loads the script.

import type { CellValue } from "../types";
import { toFormulaError } from "./args";
import { FormulaBatcher } from "./batch";
import { createFormulaTransport } from "./client";
import { cubeFunctions } from "./definitions";

// The parts of the functions runtime used here (it has no package of types in this project).
declare const CustomFunctions: {
  associate(id: string, fn: (...args: unknown[]) => unknown): void;
  Error: new (code: string, message?: string) => Error;
  ErrorCode: { invalidValue: string; notAvailable: string };
};

/**
 * The install token. The helper replaces the placeholder when it serves this script,
 * as it does in index.html. It cannot be read from the page: there is none.
 */
const TOKEN = "__CUBE_TOKEN__";

function excelError(error: unknown): Error {
  const failure = toFormulaError(error);
  const code = failure.code === "#VALUE!" ? CustomFunctions.ErrorCode.invalidValue : CustomFunctions.ErrorCode.notAvailable;
  return new CustomFunctions.Error(code, failure.message);
}

if (typeof CustomFunctions !== "undefined") {
  const functions = cubeFunctions(new FormulaBatcher(createFormulaTransport(TOKEN)));
  for (const id of Object.keys(functions)) {
    const run = functions[id];
    CustomFunctions.associate(id, (...args: unknown[]): Promise<CellValue> =>
      run(...args).catch((error: unknown) => {
        throw excelError(error);
      }),
    );
  }
}
