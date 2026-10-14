// Entry point of the right-click commands (manifest: ContextMenuCell > "Cube Explorer").
//
// Built on its own into dist/commands.js, a classic script (vite.config.ts, mode
// "commands"), and loaded by commands.html, the manifest's FunctionFile. Excel runs that
// page unseen, in a runtime of its own, when a command is chosen: nothing is shared with
// the pane but the workbook, whose settings the pane watches (useWorkbookSync).
//
// Each command is registered under its manifest FunctionName with Office.actions.associate
// and, for Excels that predate it, as a global function. It always calls
// event.completed(); when it has something to say, it first shows a small dialog
// (message.html), which closes on OK or after a few seconds.
//
// Requirement sets: AddinCommands 1.1 (the menu itself; an Excel without it shows no
// menu and still loads the add-in), ExcelApi 1.4 (as the pane), DialogApi 1.1 (guarded:
// without it the message is not shown).

import { createHttpApi, readToken } from "../api";
import { rectA1, type Rect } from "../sheet/geometry";
import { ExcelHost } from "../sheet/reports";
import { runExcel, supportsExcelApi } from "../sheet/writer";
import { COMMANDS, FUNCTION_NAMES, runCommand, type CommandDeps, type CommandName } from "./actions";

const MESSAGE_PAGE = "message.html";
/** The message closes on its own after this long. */
const MESSAGE_MS = 15000;
/** Its size, in percent of the screen. */
const MESSAGE_SIZE = { height: 24, width: 30 };

const TOO_OLD = "This version of Excel is too old for Cube Explorer. Update Excel and try again.";
const FAILED = "Cube Explorer could not do this. Open the Cube Explorer panel and try from there.";

let host: ExcelHost | null = null;

function readValues(sheetId: string, rect: Rect): Promise<unknown[][]> {
  return runExcel(async (context) => {
    const range = context.workbook.worksheets.getItem(sheetId).getRange(rectA1(rect));
    range.load("values");
    await context.sync();
    return range.values;
  });
}

function deps(): CommandDeps {
  if (!host) host = new ExcelHost();
  return { api: createHttpApi(readToken()), host, readValues };
}

function dialogSupported(): boolean {
  try {
    return Office.context.requirements.isSetSupported("DialogApi", "1.1");
  } catch {
    return false;
  }
}

/** Show `text` in a small dialog. Resolves when it has closed (or could not open). */
function showMessage(text: string): Promise<void> {
  return new Promise<void>((resolve) => {
    if (!dialogSupported()) {
      resolve();
      return;
    }
    const url = new URL(`${MESSAGE_PAGE}?text=${encodeURIComponent(text)}`, window.location.href).toString();
    try {
      Office.context.ui.displayDialogAsync(url, { ...MESSAGE_SIZE, displayInIframe: true }, (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }
        const dialog = result.value;
        let open = true;
        const finish = () => {
          if (!open) return;
          open = false;
          clearTimeout(timer);
          try {
            dialog.close();
          } catch {
            // Already closed by the user.
          }
          resolve();
        };
        const timer = setTimeout(finish, MESSAGE_MS);
        // OK in the page, or the window's own close button (DialogEventReceived).
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, finish);
        dialog.addEventHandler(Office.EventType.DialogEventReceived, finish);
      });
    } catch {
      resolve();
    }
  });
}

function complete(event: Office.AddinCommands.Event | undefined): void {
  try {
    if (event) event.completed();
  } catch {
    // Nothing more can be done; Excel ends the runtime after its own timeout.
  }
}

function handler(name: CommandName): (event?: Office.AddinCommands.Event) => void {
  return (event) => {
    void (async () => {
      let message: string | null;
      try {
        await Office.onReady();
        message = supportsExcelApi("1.4") ? await runCommand(name, deps()) : TOO_OLD;
      } catch {
        message = FAILED;
      }
      try {
        if (message) await showMessage(message);
      } finally {
        complete(event);
      }
    })();
  };
}

if (typeof Office !== "undefined") {
  const scope = window as unknown as Record<string, unknown>;
  for (const name of COMMANDS) {
    const run = handler(name);
    scope[FUNCTION_NAMES[name]] = run;
    try {
      if (Office.actions && typeof Office.actions.associate === "function") {
        Office.actions.associate(FUNCTION_NAMES[name], run);
      }
    } catch {
      // An Excel without Office.actions calls the global function instead.
    }
  }
  void Office.onReady();
}
