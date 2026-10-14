// The right-click commands on report cells (manifest: ContextMenuCell > "Cube Explorer"),
// in Windows' words: Show Details, Refresh, Refresh All, Expand, Collapse, Expand Entire
// Field, Collapse Entire Field.
//
// They run in Excel's hidden runtime for add-in commands, apart from the pane and with
// no memory of their own: each one finds the report around the active cell from the
// workbook names, loads its saved definition, and does the job with the pane's own
// modules (ReportRunner, buildBlock). To know which member a cell stands for, the report's
// query is asked again and the cell is looked up in a block built from that result, not
// written. That is trusted only when the block has the report's size and the same labels
// along the cell's row and column; otherwise the report has changed since it was last
// refreshed, and the user is asked to refresh first.
//
// Plain logic with the host and the API injected, so it is tested with PreviewHost and
// the mock API (actions.test.ts). commands.ts is the Office side.

import type { Api } from "../api";
import { offsetRect, type Rect } from "../sheet/geometry";
import type { CellAddress, ReportDef, SheetHost, SheetSelection } from "../sheet/host";
import {
  blockFields,
  buildBlock,
  expandableMembers,
  expandedDescendants,
  fieldCellAt,
  memberAt,
  valueCellMembers,
  valueFieldIndex,
  type Block,
  type BlockShape,
  type GridHeaders,
} from "../sheet/layout";
import { findField, isEmptyLayout, layoutReducer, type LayoutAction } from "../state/layout";
import { ReportRunner, toUiError, type RunOutcome } from "../state/runner";

/** The commands, by the name each has in the manifest's FunctionName (see FUNCTION_NAMES). */
export const COMMANDS = [
  "showDetails",
  "refresh",
  "refreshAll",
  "expand",
  "collapse",
  "expandField",
  "collapseField",
] as const;
export type CommandName = (typeof COMMANDS)[number];

/** `<FunctionName>` of each command in manifest.xml. */
export const FUNCTION_NAMES: Record<CommandName, string> = {
  showDetails: "cubeShowDetails",
  refresh: "cubeRefresh",
  refreshAll: "cubeRefreshAll",
  expand: "cubeExpand",
  collapse: "cubeCollapse",
  expandField: "cubeExpandField",
  collapseField: "cubeCollapseField",
};

/** Rows asked for by Show Details: Excel's own default for cube drill-through, as in the pane. */
export const DETAILS_MAX_ROWS = 1000;

export interface CommandDeps {
  api: Api;
  host: SheetHost;
  /** The values of `rect` (absolute positions) on a sheet, as the sheet holds them now. */
  readValues(sheetId: string, rect: Rect): Promise<unknown[][]>;
}

export const MESSAGES = {
  notInReport: "This cell is not part of a Cube Explorer report. Right-click a cell of a report.",
  noDefinition: "Cube Explorer cannot find this report's settings in the workbook.",
  noFields: "This report has no fields yet. Tick fields in the Cube Explorer panel to build it.",
  noConnection: "This report needs a connection. Open the Cube Explorer panel and choose one for it.",
  stale: "This report has changed since it was last refreshed. Choose Refresh, then try again.",
  notValue: "Show Details works on a number of the report. Right-click a value cell.",
  expandWhere: "Right-click an item of a row or column field (such as a year) to expand it.",
  collapseWhere: "Right-click an item of a row or column field (such as a year) to collapse it.",
  fieldWhere: "Right-click an item or the name of a row or column field.",
  fieldWhich: "This cell stands for several fields. Right-click an item of the field you mean.",
  declined:
    "Nothing was changed: the report would need more room, and some of those cells hold other things. " +
    "Use the Cube Explorer panel to decide whether to replace them.",
  notOnSheet: "This report is no longer on a sheet. Open the Cube Explorer panel to place it again.",
  helperDown: "The Cube Explorer helper is not running. Open the Cube Explorer panel for what to do.",
};

const quoted = (text: string) => `“${text}”`;

/** A failure as one sentence for the user. */
function failureMessage(error: unknown): string {
  const ui = toUiError(error);
  return ui.helperDown ? MESSAGES.helperDown : ui.message;
}

/** What to tell the user about a run, or null when the sheet says it all. */
function outcomeMessage(outcome: RunOutcome): string | null {
  switch (outcome.kind) {
    case "declined":
      return MESSAGES.declined;
    case "needs-anchor":
      return MESSAGES.notOnSheet;
    case "error":
      return outcome.error.helperDown ? MESSAGES.helperDown : outcome.error.message;
    default:
      return null;
  }
}

/**
 * A runner that never covers other cells: from a right-click there is no way to ask
 * first, so a report that would need them is left as it is ("declined").
 */
function newRunner(deps: CommandDeps): ReportRunner {
  return new ReportRunner({ api: deps.api, host: deps.host, confirmOverwrite: async () => false });
}

/**
 * Runs one command on the active cell. Resolves with a sentence for the user, or null
 * when the sheet shows what happened. Never rejects.
 */
export async function runCommand(name: CommandName, deps: CommandDeps): Promise<string | null> {
  try {
    if (name === "refreshAll") return await refreshAll(deps);
    const found = await locate(deps);
    if (typeof found === "string") return found;
    const { def } = found;
    if (name === "refresh") {
      // A report without fields has nothing to refresh; its placeholder stays.
      if (isEmptyLayout(def.layout)) return null;
      const missing = await missingConnection(deps, def);
      if (missing) return missing;
      return outcomeMessage(await newRunner(deps).run(def, { anchor: null, useSelection: false }));
    }
    if (isEmptyLayout(def.layout)) return MESSAGES.noFields;
    const missing = await missingConnection(deps, def);
    if (missing) return missing;
    const block = await blockAtCell(deps, found);
    if (typeof block === "string") return block;
    if (name === "showDetails") return await showDetails(deps, found, block);
    const action = drillAction(name, found, block);
    if (typeof action === "string") return action;
    const layout = layoutReducer(def.layout, action);
    if (layout === def.layout) return null;
    return outcomeMessage(await newRunner(deps).run({ ...def, layout }, { anchor: null, useSelection: false }));
  } catch (error) {
    return failureMessage(error);
  }
}

interface Located {
  def: ReportDef;
  /** The active cell's place in the report (offsets) and the report's size now. */
  at: NonNullable<SheetSelection["report"]>;
  cell: CellAddress;
}

/** The report around the active cell, with its saved definition; or why there is none. */
async function locate(deps: CommandDeps): Promise<Located | string> {
  const selection = await deps.host.getSelection();
  const at = selection.report;
  if (!at) return MESSAGES.notInReport;
  const defs = await deps.host.loadReports();
  const def = defs.find((d) => d.id === at.id);
  if (!def) return MESSAGES.noDefinition;
  return { def, at, cell: selection.cell };
}

async function missingConnection(deps: CommandDeps, def: ReportDef): Promise<string | null> {
  if (def.catalog === "" || def.cube === "") return MESSAGES.noConnection;
  const connections = await deps.api.connections();
  return connections.some((c) => c.id === def.connectionId) ? null : MESSAGES.noConnection;
}

/**
 * The cells to compare to tell whether the sheet still shows the result just asked for:
 * the labels along the cell's row and column (positions in the block). Value cells are
 * never compared: their numbers may have changed without moving anything.
 */
export function labelCross(shape: BlockShape, cols: number, row: number, col: number): Rect[] {
  const rects: Rect[] = [];
  // Along the row: all of it above the data, the row labels beside it.
  const width = row < shape.dataTop ? cols : Math.min(shape.headerCols, cols);
  if (width > 0) rects.push({ row, col: 0, rows: 1, cols: width });
  // Along the column: the headers above a value column; a label column down to the cell.
  const height = col < shape.headerCols ? row + 1 : shape.dataTop;
  if (height > 0) rects.push({ row: 0, col, rows: height, cols: 1 });
  return rects;
}

/** Cell text as compared: blanks alike, numbers as text, indent spaces ignored. */
function cellText(value: unknown): string {
  return value === null || value === undefined ? "" : String(value).trim();
}

/** The report's result asked again and built as a block, if the sheet still shows it. */
async function blockAtCell(deps: CommandDeps, found: Located): Promise<Block | string> {
  const { def, at, cell } = found;
  const grid = await deps.api.query(def.connectionId, { catalog: def.catalog, cube: def.cube, layout: def.layout });
  const block = buildBlock(grid, blockFields(def.layout));
  if (block.rows !== at.rows || block.cols !== at.cols) return MESSAGES.stale;
  const top = cell.row - at.row;
  const left = cell.col - at.col;
  for (const rect of labelCross(block.shape, block.cols, at.row, at.col)) {
    const sheet = await deps.readValues(cell.sheetId, offsetRect(rect, top, left));
    for (let r = 0; r < rect.rows; r++) {
      for (let c = 0; c < rect.cols; c++) {
        const shown = sheet[r] ? sheet[r][c] : undefined;
        if (cellText(shown) !== cellText(block.values[rect.row + r][rect.col + c])) return MESSAGES.stale;
      }
    }
  }
  return block;
}

/** Show Details: the rows behind the value cell, on a new sheet, as the pane does. */
async function showDetails(deps: CommandDeps, found: Located, block: Block): Promise<string | null> {
  const { def, at } = found;
  const shape = block.shape;
  const line = at.row - shape.dataTop;
  const column = at.col - shape.headerCols;
  if (line < 0 || line >= shape.dataRows || column < 0 || column >= shape.dataCols) return MESSAGES.notValue;
  const index = valueFieldIndex(block.grid.colValueIndex, column, def.layout.values.length);
  const members = index === null ? null : valueCellMembers(block, line, column);
  if (index === null || members === null) return MESSAGES.notValue;
  const response = await deps.api.details(def.connectionId, {
    catalog: def.catalog,
    cube: def.cube,
    layout: def.layout,
    members,
    measure: def.layout.values[index].measure,
    maxRows: DETAILS_MAX_ROWS,
  });
  await deps.host.addDetailsSheet(response.columns, response.rows);
  // The new sheet is in front, as in Excel; like Excel, nothing more to say.
  return null;
}

function headersOf(block: Block): GridHeaders {
  return { rowHeaders: block.grid.rowHeaders, colHeaderRows: block.grid.colHeaderRows };
}

/** The layout change Expand, Collapse or an Entire Field command makes for the cell; or why none. */
function drillAction(name: CommandName, found: Located, block: Block): LayoutAction | string {
  const { def, at } = found;
  const hit = memberAt(block, at.row, at.col);

  if (name === "expand" || name === "collapse") {
    if (!hit) return name === "expand" ? MESSAGES.expandWhere : MESSAGES.collapseWhere;
    const member = hit.member;
    if (name === "expand") {
      if (!member.hasChildren) return `${quoted(member.caption)} has no items under it.`;
      if (member.expanded) return `${quoted(member.caption)} is already expanded.`;
      return { type: "expand", hierarchy: member.hierarchy, member: member.uniqueName };
    }
    if (!member.expanded) return `${quoted(member.caption)} is not expanded.`;
    return {
      type: "collapse",
      hierarchy: member.hierarchy,
      member: member.uniqueName,
      descendants: expandedDescendants(headersOf(block), hit.axis, member),
    };
  }

  // Expand / Collapse Entire Field: the field of an item, or of a field's name cell.
  let hierarchy: string | null = hit ? hit.member.hierarchy : null;
  if (!hierarchy) {
    const named = fieldCellAt(block, at.row, at.col);
    if (!named || named.area === "filters") return MESSAGES.fieldWhere;
    const fields = def.layout[named.area];
    if (named.labels && fields.length > 1) return MESSAGES.fieldWhich;
    hierarchy = fields[named.index] ? fields[named.index].hierarchy : null;
  }
  const field = hierarchy ? findField(def.layout, hierarchy) : null;
  if (!field || !hierarchy) return MESSAGES.fieldWhere;
  if (name === "expandField") {
    const open = expandableMembers(headersOf(block), hierarchy);
    if (!open.anyChildren) return `The items of ${quoted(field.caption)} have no items under them.`;
    if (open.members.length === 0) return `${quoted(field.caption)} is already expanded all the way.`;
    return { type: "expandField", hierarchy, members: open.members };
  }
  if (field.expanded.length === 0) return `Nothing in ${quoted(field.caption)} is expanded.`;
  return { type: "collapseField", hierarchy };
}

/** Refresh All, as in the pane: every report of the workbook, then the cube formulas. */
async function refreshAll(deps: CommandDeps): Promise<string | null> {
  const defs = (await deps.host.loadReports()).filter((d) => !isEmptyLayout(d.layout));
  const runner = newRunner(deps);
  let done = 0;
  let unbound = 0;
  let lost = 0;
  let skipped = 0;
  let failed = 0;
  let firstError: string | null = null;
  if (defs.length > 0) {
    const connections = await deps.api.connections();
    for (const def of defs) {
      if (def.catalog === "" || def.cube === "" || !connections.some((c) => c.id === def.connectionId)) {
        unbound++;
        continue;
      }
      const outcome = await runner.run(def, { anchor: null, useSelection: false });
      if (outcome.kind === "written" || outcome.kind === "empty") done++;
      else if (outcome.kind === "needs-anchor") lost++;
      else if (outcome.kind === "declined") skipped++;
      else if (outcome.kind === "error") {
        failed++;
        if (!firstError) firstError = outcomeMessage(outcome);
        if (outcome.error.helperDown) break; // the rest would fail the same way
      }
    }
  }
  try {
    // As Windows' Refresh All does for cube formulas: they ask the cube again.
    await deps.host.recalculate();
  } catch {
    // Excel is busy; the formulas recalculate with the next change.
  }
  if (unbound + lost + skipped + failed === 0) return null;
  const parts = [`${done} of ${defs.length} reports refreshed.`];
  if (unbound > 0) parts.push(`${unbound} ${unbound === 1 ? "needs" : "need"} a connection (open the Cube Explorer panel).`);
  if (lost > 0) parts.push(`${lost} could not be found on a sheet.`);
  if (skipped > 0) parts.push(`${skipped} would have covered other cells and ${skipped === 1 ? "was" : "were"} left as ${skipped === 1 ? "it was" : "they were"}.`);
  if (failed > 0 && firstError) parts.push(firstError);
  return parts.join(" ");
}
