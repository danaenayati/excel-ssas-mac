// The task pane. Two screens: Connections, and the Explorer for the report being
// edited. All sheet access goes through `host`, all cube access through `api`; this
// file only wires user actions to the reducers and the runner, and decides what to
// show.

import { useCallback, useEffect, useMemo, useReducer, useRef, useState } from "react";
import { isCancelled, type Api } from "./api";
import { Areas } from "./components/Areas";
import { Connections } from "./components/Connections";
import { CubePicker } from "./components/CubePicker";
import { FieldList } from "./components/FieldList";
import { MemberPicker } from "./components/MemberPicker";
import { Menu, type MenuEntry } from "./components/Menu";
import { PreviewSheet } from "./components/PreviewSheet";
import { StatusBar, type Drill } from "./components/StatusBar";
import { CopyButton, Dialog, ErrorNote, Icon, Spinner } from "./components/common";
import { CalculatedItemDialog, CalculatedMeasureDialog, GroupDialog } from "./components/CalculationDialogs";
import { NumberFormatDialog } from "./components/NumberFormatDialog";
import { ShowAsDialog } from "./components/ShowAsDialog";
import { refreshEveryMenu, refreshEveryNote } from "./components/refreshEvery";
import { chartMenuItems, type ChartActions } from "./components/chartMenu";
import { StyleOptions } from "./components/StyleOptions";
import { InsertSlicer } from "./components/InsertSlicer";
import { Slicers } from "./components/Slicers";
import { conversionBlocker, convertReport, formulaConnection } from "./sheet/convert";
import { rectA1 } from "./sheet/geometry";
import {
  describeCell,
  type CellAddress,
  type Extent,
  type ReportDef,
  type SheetClick,
  type SheetHost,
  type SheetSelection,
} from "./sheet/host";
import { REFRESH_MINUTES, REFRESH_ON_OPEN } from "./sheet/host";
import { guessFieldCell, labelsCaption, type MemberHit } from "./sheet/layout";
import { PreviewHost } from "./sheet/preview";
import {
  findField,
  grandTotalsShown,
  isEmptyLayout,
  calculatedMeasureName,
  layoutReducer,
  locate,
  needsBaseField,
  reportForm,
  selectionSummary,
  subtotalsChoice,
  valueIndex,
  type BooleanOption,
  type LayoutAction,
  type ReportForm,
  type SubtotalsChoice,
} from "./state/layout";
import { currentReducer, decidePlace, newReportDef, placeLabel, reportForSheet } from "./state/report";
import {
  ReportRunner,
  toUiError,
  type RunOptions,
  type RunOutcome,
  type UiError,
} from "./state/runner";
import { useCached, type Cache } from "./state/useCached";
import { clockTime, excelWasBusy, normalizeMinutes, type AutoRefreshResult } from "./state/autoRefresh";
import { useAutoRefresh } from "./state/useAutoRefresh";
import { syncReports } from "./state/sync";
import { useWorkbookSync } from "./state/useWorkbookSync";
import type {
  CalculatedMeasure,
  Catalog,
  Connection,
  Cube,
  Fields,
  LayoutOptions,
  PlacedField,
  ShowValuesAs,
  Status,
} from "./types";
import { plural } from "./util";

interface Props {
  api: Api;
  host: SheetHost;
  /** The page was opened with ?mock=1: canned data, no helper. */
  mock: boolean;
}

type Screen = "connections" | "explorer";

type Question =
  | { kind: "overwrite"; target: Extent; answer: (ok: boolean) => void }
  | { kind: "deleteReport"; def: ReportDef; onSheet: boolean }
  | { kind: "convert"; def: ReportDef }
  | { kind: "reset"; what: "database" | "cube"; apply: () => void }
  | { kind: "query"; mdx: string }
  | { kind: "numberFormat"; measure: string; caption: string; current: string }
  | { kind: "showAs"; measure: string; showAs: ShowValuesAs }
  | { kind: "groups"; hierarchy: string }
  | { kind: "calculatedItems"; hierarchy: string }
  | { kind: "calculatedMeasure"; measure: CalculatedMeasure | null };

/** Pause between a layout change and its query, so that ticking three fields is one query. */
const DEBOUNCE_MS = 400;
/** Warn about the helper's certificate when it has this many days left, or fewer. */
const CERT_WARNING_DAYS = 30;
/** While the user clicks around one sheet, re-read where the reports are this often at most. */
const LOCATION_REFRESH_MS = 3000;
/** Rows asked for by Show Details, as Excel's own default for cube drill-through. */
const DETAILS_MAX_ROWS = 1000;
/** Said when a report with charts goes: as Windows does with a PivotChart, they stay as pictures. */
const CHARTS_STAY = " Its charts stay as they are but no longer update.";

function upsert(list: ReportDef[], def: ReportDef): ReportDef[] {
  const index = list.findIndex((d) => d.id === def.id);
  if (index < 0) return [...list, def];
  const next = list.slice();
  next[index] = def;
  return next;
}

function useLatest<T>(value: T) {
  const ref = useRef(value);
  ref.current = value;
  return ref;
}

const OPTION_LABELS: [BooleanOption, string][] = [
  ["nonEmpty", "Hide empty"],
  ["repeatLabels", "Repeat All Item Labels"],
  // Windows: PivotTable Options > Totals & Filters, for cube PivotTables.
  ["includeFilteredItems", "Include filtered items in totals"],
];

/** Excel's Subtotals choices (Design tab), in Excel's words. */
const SUBTOTALS: { value: SubtotalsChoice; label: string }[] = [
  { value: "none", label: "Do Not Show Subtotals" },
  { value: "bottom", label: "Show all Subtotals at Bottom of Group" },
  { value: "top", label: "Show all Subtotals at Top of Group" },
];

/** Excel's Report Layout choices (Design tab). */
const REPORT_LAYOUTS: { value: ReportForm; label: string }[] = [
  { value: "compact", label: "Compact Form" },
  { value: "outline", label: "Outline Form" },
  { value: "tabular", label: "Tabular Form" },
];

/** Excel's Grand Totals choices (Design tab), with Excel's own reading of rows/columns. */
const GRAND_TOTALS: { value: string; label: string; row: boolean; column: boolean }[] = [
  { value: "none", label: "Off for Rows and Columns", row: false, column: false },
  { value: "both", label: "On for Rows and Columns", row: true, column: true },
  // "Rows" means the total at the end of each row: the Grand Total column.
  { value: "rows", label: "On for Rows Only", row: false, column: true },
  { value: "columns", label: "On for Columns Only", row: true, column: false },
];

function grandTotalsChoice(options: LayoutOptions): string {
  const shown = grandTotalsShown(options);
  const found = GRAND_TOTALS.find((g) => g.row === shown.row && g.column === shown.column);
  return found ? found.value : "both";
}

/** "For empty cells show": applied when the box is left or Return is pressed. */
function EmptyCellText({ value, onCommit }: { value: string; onCommit: (text: string) => void }) {
  const [text, setText] = useState(value);
  useEffect(() => setText(value), [value]);
  return (
    <input
      className="input empty-text"
      value={text}
      maxLength={20}
      aria-label="For empty cells show"
      placeholder="(nothing)"
      onChange={(event) => setText(event.target.value)}
      onBlur={() => {
        if (text !== value) onCommit(text);
      }}
      onKeyDown={(event) => {
        if (event.key === "Enter") onCommit(text);
      }}
    />
  );
}

export function App({ api, host, mock }: Props) {
  const [phase, setPhase] = useState<"loading" | "ready" | UiError>("loading");
  const [status, setStatus] = useState<Status | null>(null);
  const [connections, setConnections] = useState<Connection[]>([]);
  const [reports, setReports] = useState<ReportDef[]>([]);
  const [locations, setLocations] = useState<Record<string, Extent | null>>({});
  const [current, dispatch] = useReducer(currentReducer, null);
  const [screen, setScreen] = useState<Screen>("connections");
  const [selection, setSelection] = useState<SheetSelection | null>(null);
  const [busy, setBusy] = useState<string | null>(null);
  const [error, setError] = useState<UiError | null>(null);
  const [note, setNote] = useState<string | null>(null);
  const [defer, setDefer] = useState(false);
  const [ran, setRan] = useState<{ id: string; revision: number } | null>(null);
  const [question, setQuestion] = useState<Question | null>(null);
  const [picker, setPicker] = useState<string | null>(null);
  /** Set while the checklist was opened from "Row Labels" or "Column Labels": it offers that area's fields. */
  const [pickerFrom, setPickerFrom] = useState<"rows" | "columns" | null>(null);
  useEffect(() => {
    if (picker === null) setPickerFrom(null);
  }, [picker]);
  const [search, setSearch] = useState("");
  const [reportMenu, setReportMenu] = useState<HTMLElement | null>(null);
  /** The Insert Slicers panel is open. */
  const [insertSlicer, setInsertSlicer] = useState(false);
  const [generation, setGeneration] = useState(0);
  const [refreshOnOpen, setRefreshOnOpen] = useState(false);
  /** "Refresh Every": the workbook's interval in minutes, 0 when off. */
  const [refreshMinutes, setRefreshMinutes] = useState(0);
  /** Set at start-up when the workbook asks for a refresh on open; cleared once done. */
  const [openingRefresh, setOpeningRefresh] = useState(false);
  const [, setTick] = useState(0);

  const currentRef = useLatest(current);
  const reportsRef = useLatest(reports);
  const connectionsRef = useLatest(connections);
  const ranRef = useLatest(ran);
  const busyPrefix = useRef("");
  const cache = useRef<Cache>(new Map()).current;
  const lastSheet = useRef<string | null>(null);
  const lastLocationRead = useRef(0);
  const locationSeq = useRef(0);
  /** Set to keep the note through the next run, which would otherwise clear it. */
  const keepNote = useRef(false);
  /** The Show Details request in flight, so that Cancel can stop it. */
  const detailsFlight = useRef<{ abort: AbortController; connectionId: string } | null>(null);
  /** When the pane last wrote a report definition (see useWorkbookSync). */
  const lastOwnWrite = useRef(0);

  // -- runner ---------------------------------------------------------------

  const runner = useMemo(
    () =>
      new ReportRunner({
        api,
        host,
        confirmOverwrite: (_def, target, signal) =>
          new Promise<boolean>((resolve) => {
            const answer = (ok: boolean) => {
              setQuestion((q) => (q && q.kind === "overwrite" && q.answer === answer ? null : q));
              resolve(ok);
            };
            signal.addEventListener("abort", () => answer(false));
            setQuestion({ kind: "overwrite", target, answer });
          }),
        onStage: (stage) =>
          setBusy(busyPrefix.current + (stage === "query" ? "Asking the cube…" : "Writing to the sheet…")),
      }),
    [api, host],
  );

  const refreshSelection = useCallback(() => {
    host.getSelection().then(setSelection, () => undefined);
  }, [host]);

  /**
   * Read from the workbook where every report is now. What the pane remembers goes
   * stale behind its back: a deleted sheet takes a report's cells with it. Resolves
   * with what was read, or null when Excel was busy or a newer read took over.
   */
  const refreshLocations = useCallback(async (): Promise<Record<string, Extent | null> | null> => {
    const seq = ++locationSeq.current;
    lastLocationRead.current = Date.now();
    const found: Record<string, Extent | null> = {};
    try {
      for (const def of reportsRef.current) found[def.id] = await host.getReportExtent(def.id);
    } catch {
      return null; // Excel is busy; the next selection change tries again.
    }
    if (seq !== locationSeq.current) return null;
    setLocations((map) => ({ ...map, ...found }));
    return found;
  }, [host, reportsRef]);

  /** One report's location, read now, and remembered. Null when it has no cells. */
  const liveExtent = useCallback(
    async (reportId: string): Promise<Extent | null> => {
      const extent = await host.getReportExtent(reportId);
      setLocations((map) => ({ ...map, [reportId]: extent }));
      return extent;
    },
    [host],
  );

  /** Fold the result of a run into the pane's state. `single` runs report their own notes. */
  const applyOutcome = useCallback(
    (def: ReportDef, outcome: RunOutcome, single: boolean) => {
      lastOwnWrite.current = Date.now();
      const isCurrent = () => {
        const cur = currentRef.current;
        return cur !== null && cur.def.id === def.id;
      };
      const onSheet = (extent: Extent, saved: ReportDef = def) => {
        setReports((list) => upsert(list, saved));
        // A chart deleted in Excel has just dropped out of the saved definition.
        dispatch({ type: "charts", id: def.id, charts: saved.charts ?? [] });
        setLocations((map) => ({ ...map, [def.id]: extent }));
        // Addressed by id rather than checked here: the reducer sees the report being
        // edited as it is now, even if this runs before the pane has re-rendered.
        dispatch({ type: "saved", id: def.id });
        refreshSelection();
      };
      switch (outcome.kind) {
        case "written":
          onSheet(outcome.extent, outcome.def);
          if (single && isCurrent() && outcome.chartError) setError(outcome.chartError);
          break;
        case "empty":
          // The placeholder is on the sheet and the definition saved, or (a new report
          // that has not been given a place) nothing happened.
          if (outcome.extent) onSheet(outcome.extent);
          break;
        case "cancelled":
          if (single && isCurrent()) setNote("Cancelled. The sheet was not changed.");
          break;
        case "declined":
          if (single && isCurrent()) setNote("Not updated, so that the other cells stay as they are.");
          break;
        case "needs-anchor":
          setLocations((map) => ({ ...map, [def.id]: null }));
          // The selected cell is inside another report, the one place it could have gone.
          if (single && isCurrent()) setNote("Select an empty cell outside other reports, then choose Place here.");
          break;
        case "error":
          if (single && isCurrent()) setError(outcome.error);
          break;
        case "superseded":
          break;
      }
      setTick((n) => n + 1);
    },
    [currentRef, refreshSelection],
  );

  /** Run one report and show the result. Resolves with null when there is nothing to ask the cube. */
  const execute = useCallback(
    async (def: ReportDef, revision: number, options: RunOptions): Promise<RunOutcome | null> => {
      const done = () => setRan({ id: def.id, revision });
      const empty = isEmptyLayout(def.layout);
      // A report with fields needs its connection (the banner asks for one) and a cube.
      // One without fields only puts its placeholder on the sheet.
      const askable =
        connectionsRef.current.some((c) => c.id === def.connectionId) && def.catalog !== "" && def.cube !== "";
      if (!empty && !askable) {
        done();
        return null;
      }
      setError(null);
      if (keepNote.current) keepNote.current = false;
      else setNote(null);
      busyPrefix.current = "";
      setBusy(empty ? "Writing to the sheet…" : "Asking the cube…");
      const outcome = await runner.run(def, options);
      if (outcome.kind !== "superseded") {
        if (!runner.busy) setBusy(null);
        done();
      }
      applyOutcome(def, outcome, true);
      return outcome;
    },
    [applyOutcome, connectionsRef, runner],
  );

  const runCurrent = useCallback(async () => {
    const cur = currentRef.current;
    if (!cur) return;
    // The selected cell is used only when the report has no cells: it is new, or its
    // sheet was deleted. Either way it appears where the user is, as in Excel.
    await execute(cur.def, cur.revision, { anchor: cur.anchor, useSelection: true });
  }, [currentRef, execute]);
  const runCurrentRef = useLatest(runCurrent);

  // Every change to the definition re-queries: at once for Refresh and drill, after a
  // short pause for layout edits, and not at all while "Defer updates" is on.
  const runKey = current ? `${current.def.id}:${current.revision}` : "";
  useEffect(() => {
    const cur = currentRef.current;
    if (!cur || cur.revision === 0) return;
    const last = ranRef.current;
    if (last && last.id === cur.def.id && last.revision === cur.revision) return;
    if (defer && !cur.urgent) return;
    const timer = window.setTimeout(() => void runCurrentRef.current(), cur.urgent ? 0 : DEBOUNCE_MS);
    return () => window.clearTimeout(timer);
  }, [runKey, defer, currentRef, ranRef, runCurrentRef]);

  // -- start-up -------------------------------------------------------------

  const openReport = useCallback((def: ReportDef) => {
    dispatch({ type: "open", def, persisted: true });
    setScreen("explorer");
    setError(null);
    setNote(null);
    setPicker(null);
    setSearch("");
  }, []);

  const startDraft = useCallback(
    (connectionId: string, catalog = "", cube = "", anchor: CellAddress | null = null): ReportDef => {
      const draft = newReportDef(connectionId, catalog, cube);
      dispatch({ type: "open", def: draft, persisted: false, anchor });
      setScreen("explorer");
      setError(null);
      setNote(null);
      setPicker(null);
      setSearch("");
      return draft;
    },
    [],
  );

  const boot = useCallback(async () => {
    setPhase("loading");
    try {
      const [helper, list] = await Promise.all([api.status(), api.connections()]);
      setStatus(helper);
      setConnections(list);
    } catch (failure) {
      setPhase(toUiError(failure));
      return;
    }
    let defs: ReportDef[] = [];
    const where: Record<string, Extent | null> = {};
    let active: SheetSelection | null = null;
    let refresh = false;
    let minutes = 0;
    try {
      defs = await host.loadReports();
      for (const def of defs) where[def.id] = await host.getReportExtent(def.id);
      active = await host.getSelection();
      refresh = await host.getFlag(REFRESH_ON_OPEN);
      minutes = normalizeMinutes(await host.getNumber(REFRESH_MINUTES));
    } catch {
      // The sheet cannot be read right now (a cell is being edited, say). The pane
      // still opens; reports show up at the next start.
    }
    setRefreshOnOpen(refresh);
    setRefreshMinutes(minutes);
    setOpeningRefresh(refresh && defs.length > 0);
    setReports(defs);
    setLocations(where);
    setSelection(active);
    lastSheet.current = active ? active.cell.sheetId : null;
    lastLocationRead.current = Date.now();
    // The report at the selected cell, else one on the sheet in view, else a new one
    // there: never a report on another sheet, where ticked fields would not show.
    const hit = active && active.report ? defs.find((d) => d.id === active.report?.id) : undefined;
    const choice = hit
      ? { kind: "open" as const, id: hit.id }
      : active && defs.length > 0
        ? reportForSheet(null, active.cell.sheetId, defs, where)
        : null;
    const first = choice && choice.kind === "open" ? defs.find((d) => d.id === choice.id) : defs[0];
    if (choice && choice.kind === "new") startDraft(defs[0].connectionId, defs[0].catalog, defs[0].cube);
    else if (first) openReport(first);
    else setScreen("connections");
    setPhase("ready");
  }, [api, host, openReport, startDraft]);

  useEffect(() => {
    void boot();
  }, [boot]);

  // "Refresh when the workbook opens": once, after start-up, like Refresh All.
  const refreshAllRef = useRef<() => Promise<unknown>>(async () => undefined);
  useEffect(() => {
    if (phase !== "ready" || !openingRefresh) return;
    setOpeningRefresh(false);
    void refreshAllRef.current();
  }, [phase, openingRefresh]);

  /**
   * The user arrived on another sheet, with the selected cell in no report: edit the
   * report on that sheet, or a new one there (see `reportForSheet`). Notes stay, so that
   * "rows on sheet Details1" survives Show Details moving to its new sheet.
   */
  const followSheet = useCallback(
    (sheetId: string, found: Record<string, Extent | null>) => {
      const cur = currentRef.current;
      if (!cur || lastSheet.current !== sheetId) return;
      const choice = reportForSheet(cur, sheetId, reportsRef.current, found);
      if (choice.kind === "keep") return;
      const def =
        choice.kind === "open"
          ? reportsRef.current.find((d) => d.id === choice.id)
          : newReportDef(cur.def.connectionId, cur.def.catalog, cur.def.cube);
      if (!def) return;
      dispatch({ type: "open", def, persisted: choice.kind === "open" });
      setError(null);
      setPicker(null);
    },
    [currentRef, reportsRef],
  );

  // Selecting a cell inside a report makes that report the one being edited. Arriving
  // on another sheet (or a sheet being deleted) is when report locations go stale.
  useEffect(() => {
    if (phase !== "ready") return;
    return host.onSelectionChanged((next) => {
      setSelection(next);
      const sheetChanged = lastSheet.current !== next.cell.sheetId;
      lastSheet.current = next.cell.sheetId;
      if (sheetChanged || Date.now() - lastLocationRead.current > LOCATION_REFRESH_MS) {
        void refreshLocations().then((found) => {
          if (found && sheetChanged && !next.report) followSheet(next.cell.sheetId, found);
        });
      }
      if (!next.report) return;
      const cur = currentRef.current;
      if (cur && cur.def.id === next.report.id) return;
      const id = next.report.id;
      const def = reportsRef.current.find((d) => d.id === id);
      if (def) openReport(def);
    });
  }, [phase, host, currentRef, reportsRef, openReport, refreshLocations, followSheet]);

  const reloadConnections = useCallback(async () => {
    setConnections(await api.connections());
  }, [api]);

  // -- filtering from the sheet -------------------------------------------------

  // The filter cells of a report (a page filter's cells, the captions of row and column
  // fields) work as buttons: one click opens that field's checklist here in the pane.
  // Moving onto one with the keyboard does not; the status bar offers the same action.
  const handleClick = useCallback(
    (click: SheetClick) => {
      const id = click.report.id;
      const cur = currentRef.current;
      const saved = reportsRef.current.find((d) => d.id === id) ?? null;
      const target = cur && cur.def.id === id ? cur.def : saved;
      if (!target) return;

      let onSheet: PlacedField | null = null;
      // Compact form's "Row Labels" / "Column Labels": the checklist offers every field of that area.
      let labels: "rows" | "columns" | null = null;
      const known = runner.buttonAt(id, click.report);
      if (known !== undefined) {
        onSheet = known;
        labels = runner.labelsAt(id, click.report);
      } else if (saved) {
        // No result in memory (the workbook was just reopened): go by where the saved
        // layout puts its filter cells, and for a caption cell believe it only if the
        // cell really holds that field's caption.
        const guess = guessFieldCell(saved.layout, click.report.row, click.report.col);
        const field = guess ? saved.layout[guess.area][guess.index] ?? null : null;
        const caption = guess && guess.labels ? labelsCaption(guess.area) : field ? field.caption : null;
        if (guess && field && (guess.area === "filters" || click.value === caption)) {
          onSheet = field;
          if (guess.labels && guess.area !== "filters") labels = guess.area;
        }
      }
      if (!onSheet) return; // an ordinary cell

      // The field may have been taken out of the layout in the pane since.
      const field = labels ? target.layout[labels][0] ?? null : findField(target.layout, onSheet.hierarchy);
      if (!field) return;
      if (!cur || cur.def.id !== id) openReport(target);
      else setScreen("explorer");
      // Without its connection the report shows the "choose a connection" banner instead.
      if (connectionsRef.current.some((c) => c.id === target.connectionId)) {
        setPickerFrom(labels);
        setPicker(field.hierarchy);
      }
    },
    [connectionsRef, currentRef, openReport, reportsRef, runner],
  );

  useEffect(() => {
    if (phase !== "ready") return;
    return host.onCellClicked(handleClick);
  }, [phase, host, handleClick]);

  // -- data for the report being edited ---------------------------------------

  const def = current ? current.def : null;
  const connection = def ? connections.find((c) => c.id === def.connectionId) ?? null : null;
  const connectionId = connection ? connection.id : null;
  const catalog = def ? def.catalog : "";
  const cube = def ? def.cube : "";

  const catalogs = useCached<Catalog[]>(
    cache,
    connectionId ? `${generation}|catalogs|${connectionId}` : null,
    () => api.catalogs(connectionId as string),
  );
  const cubes = useCached<Cube[]>(
    cache,
    connectionId && catalog ? `${generation}|cubes|${connectionId}|${catalog}` : null,
    () => api.cubes(connectionId as string, catalog),
  );
  const fields = useCached<Fields>(
    cache,
    connectionId && catalog && cube ? `${generation}|fields|${connectionId}|${catalog}|${cube}` : null,
    () => api.fields(connectionId as string, catalog, cube),
  );

  // With nothing chosen yet, start on the first database and the first cube.
  useEffect(() => {
    if (!def || def.catalog !== "" || !catalogs.data || catalogs.data.length === 0) return;
    dispatch({ type: "defaults", catalog: catalogs.data[0].name });
  }, [def, catalogs.data]);
  useEffect(() => {
    if (!def || def.catalog === "" || def.cube !== "" || !cubes.data || cubes.data.length === 0) return;
    dispatch({ type: "defaults", cube: cubes.data[0].name });
  }, [def, cubes.data]);

  // -- actions ----------------------------------------------------------------

  const act = (action: LayoutAction, urgent = false) => {
    // A value shown along a base field that has just left Rows and Columns goes back
    // to No Calculation (the reducer does that); say so, as Excel would.
    const cur = currentRef.current;
    if (cur && action.type !== "setShowAs") {
      const before = cur.def.layout;
      const after = layoutReducer(before, action);
      const lost = before.values.filter((v) => {
        if (!needsBaseField(v.showAs)) return false;
        const now = after.values.find((a) => a.measure === v.measure && a.caption === v.caption);
        return now !== undefined && !needsBaseField(now.showAs);
      });
      if (lost.length > 0) {
        const fieldCaption = (name: string | null | undefined) =>
          [...before.rows, ...before.columns, ...before.filters].find((f) => f.hierarchy === name)?.caption ?? "its base field";
        keepNote.current = true;
        setNote(
          lost
            .map((v) => `${v.caption} is back to No Calculation: ${fieldCaption(v.baseField)} is no longer in Rows or Columns.`)
            .join(" "),
        );
      }
    }
    dispatch({ type: "layout", action, urgent });
  };

  const chooseConnection = (chosen: Connection) => {
    const cur = currentRef.current;
    if (cur && cur.persisted && !connectionsRef.current.some((c) => c.id === cur.def.connectionId)) {
      // The report being edited has lost its connection: this is the new one for it.
      rebind(chosen.id);
      setScreen("explorer");
      return;
    }
    if (cur && cur.def.connectionId === chosen.id) {
      setScreen("explorer");
      return;
    }
    startDraft(chosen.id, "", "", cur && !cur.persisted ? cur.anchor : null);
  };

  const rebind = (id: string) => {
    const cur = currentRef.current;
    if (!cur) return;
    dispatch({ type: "rebind", connectionId: id });
    if (!cur.persisted) return;
    const next = { ...cur.def, connectionId: id };
    host.saveReport(next).then(
      () => {
        lastOwnWrite.current = Date.now();
        setReports((list) => upsert(list, next));
      },
      (failure) => setError(toUiError(failure)),
    );
  };

  const changeCube = (what: "database" | "cube", value: string) => {
    const cur = currentRef.current;
    if (!cur) return;
    const apply = () =>
      dispatch(what === "database" ? { type: "setCatalog", catalog: value } : { type: "setCube", cube: value });
    const hasFields =
      !isEmptyLayout(cur.def.layout) || cur.def.layout.filters.length > 0 || (cur.def.layout.slicers ?? []).length > 0;
    if (hasFields) setQuestion({ kind: "reset", what, apply });
    else apply();
  };

  /** The selected cell and the current report's location, both read from the workbook now. */
  const readPlace = async (reportId: string) => {
    try {
      const active = await host.getSelection();
      const extent = await liveExtent(reportId);
      setSelection(active);
      return { active, extent };
    } catch (failure) {
      setError(toUiError(failure));
      return null;
    }
  };

  /**
   * The button next to the report's location. If the report being edited has no cells
   * (it is new, or its sheet was deleted) it is placed at the selected cell, fields and
   * all. If it is on a sheet, a new report is started at the selected cell.
   */
  const placeHere = async () => {
    const cur = currentRef.current;
    if (!cur) return;
    const place = await readPlace(cur.def.id);
    if (!place) return;
    setError(null);
    const decision = decidePlace(cur.def.id, place.extent, place.active);
    const where = describeCell(place.active.cell);
    if (decision.kind === "blocked") {
      setNote(
        decision.inside === "this"
          ? "This cell is part of this report. Select a cell outside it to start another report."
          : "This cell is inside another report. Select a cell outside it first.",
      );
      return;
    }
    const options: RunOptions = { anchor: place.active.cell, useSelection: false };
    if (decision.kind === "new") {
      // Starts with its placeholder on the sheet, so the click has something to show.
      const draft = startDraft(cur.def.connectionId, cur.def.catalog, cur.def.cube, place.active.cell);
      const outcome = await execute(draft, 0, options);
      if (outcome && outcome.kind === "empty" && outcome.extent) {
        setNote(`New report at ${where}. Tick a field to fill it.`);
      }
      return;
    }
    // Remembered too, so that a later try goes to the same cell if this one fails.
    dispatch({ type: "pin", anchor: place.active.cell });
    const outcome = await execute(cur.def, cur.revision, options);
    if (outcome && outcome.kind === "empty" && outcome.extent) {
      setNote(`Report placed at ${where}. Tick a field to fill it.`);
    } else if (outcome === null) {
      setNote(`The report will go to ${where} once it has a connection and a cube.`);
    }
  };

  /** Report menu: take the report from where it is and write it at the selected cell. */
  const moveHere = async () => {
    const cur = currentRef.current;
    if (!cur) return;
    const place = await readPlace(cur.def.id);
    if (!place) return;
    if (place.active.report) {
      setError(null);
      setNote("Select a cell outside every report first, then choose Move this report here.");
      return;
    }
    const outcome = await execute(cur.def, cur.revision, {
      anchor: place.active.cell,
      useSelection: false,
      relocate: true,
    });
    if (outcome === null) {
      setNote("Choose a connection for this report first.");
    } else if (outcome.kind === "written" || (outcome.kind === "empty" && outcome.extent)) {
      setNote(`Report moved to ${describeCell(place.active.cell)}.`);
    }
  };

  /**
   * Show Details for the selected value cell: ask the helper for the rows behind it and
   * put them on a new sheet, as double-clicking a number in a PivotTable does.
   */
  const showDetails = async () => {
    const cur = currentRef.current;
    const at = selection && selection.report;
    if (!cur || !at || at.id !== cur.def.id) return;
    const target = runner.detailsFor(cur.def.id, at);
    const conn = connectionsRef.current.find((c) => c.id === cur.def.connectionId);
    if (!target || !conn) return;
    setError(null);
    setNote(null);
    setBusy("Getting the details…");
    const flight = { abort: new AbortController(), connectionId: conn.id };
    detailsFlight.current = flight;
    try {
      const response = await api.details(
        conn.id,
        {
          catalog: cur.def.catalog,
          cube: cur.def.cube,
          layout: target.layout,
          members: target.members,
          measure: target.measure,
          maxRows: DETAILS_MAX_ROWS,
        },
        flight.abort.signal,
      );
      setBusy("Writing the details…");
      const sheetName = await host.addDetailsSheet(response.columns, response.rows);
      const count = response.rows.length.toLocaleString("en-US");
      setNote(
        response.truncated
          ? `Showing the first ${count} rows, on sheet ${sheetName}.`
          : `${plural(response.rows.length, "row", "rows")} on sheet ${sheetName}.`,
      );
    } catch (failure) {
      if (isCancelled(failure)) setNote("Cancelled.");
      else setError(toUiError(failure));
    } finally {
      if (detailsFlight.current === flight) detailsFlight.current = null;
      setBusy(runner.busy ? "Asking the cube…" : null);
    }
  };

  /**
   * Refresh All. `auto` when "Refresh Every" runs it: the note then says when, and when
   * Excel was busy it stops and answers "busy", so that the next tick tries again.
   */
  const refreshAll = async (auto = false): Promise<AutoRefreshResult> => {
    const cur = currentRef.current;
    const list = reportsRef.current.map((d) => (cur && cur.def.id === d.id ? cur.def : d));
    const autoNote = () => `Refreshed automatically at ${clockTime(Date.now())}.`;
    if (list.length === 0) {
      // No reports, but there may be cube formulas (Convert to Formulas): ask them again.
      if (await recalculateFormulas()) setNote(auto ? autoNote() : "Cube formulas recalculated.");
      autoRefresh.noteRefreshed();
      return "done";
    }
    setError(null);
    setNote(null);
    let done = 0;
    let unbound = 0;
    let lost = 0;
    let skipped = 0;
    let failed = 0;
    let firstError: UiError | null = null;
    let stopped = false;
    for (let i = 0; i < list.length; i++) {
      const item = list[i];
      // A report with no fields has nothing to refresh; its placeholder stays as it is.
      if (isEmptyLayout(item.layout)) continue;
      if (!connectionsRef.current.some((c) => c.id === item.connectionId)) {
        unbound++;
        continue;
      }
      busyPrefix.current = list.length > 1 ? `Report ${i + 1} of ${list.length}: ` : "";
      setBusy(busyPrefix.current + "Asking the cube…");
      const outcome = await runner.run(item, { anchor: null, useSelection: false });
      applyOutcome(item, outcome, false);
      if (outcome.kind === "superseded") {
        busyPrefix.current = "";
        return "busy"; // another run took over and owns the status line now
      }
      if (cur && cur.def.id === item.id) setRan({ id: item.id, revision: cur.revision });
      if (outcome.kind === "cancelled") {
        stopped = true;
        break;
      }
      if (outcome.kind === "written") done++;
      else if (outcome.kind === "needs-anchor") lost++;
      else if (outcome.kind === "declined") skipped++;
      else if (outcome.kind === "error") {
        if (auto && excelWasBusy(outcome.error)) {
          // A cell is being edited: leave it, and try again at the next tick.
          busyPrefix.current = "";
          if (!runner.busy) setBusy(null);
          return "busy";
        }
        failed++;
        if (!firstError) firstError = outcome.error;
        if (outcome.error.helperDown) break; // the rest would fail the same way
      }
    }
    busyPrefix.current = "";
    if (!runner.busy) setBusy(null);
    // As Windows' Refresh All does for cube formulas: they ask the cube again.
    if (!stopped) await recalculateFormulas();
    const parts = [
      stopped
        ? `Cancelled after ${plural(done, "report", "reports")}.`
        : auto
          ? autoNote()
          : `${plural(done, "report", "reports")} refreshed.`,
    ];
    if (unbound > 0) parts.push(`${unbound} ${unbound === 1 ? "needs" : "need"} a connection.`);
    if (lost > 0) parts.push(`${lost} could not be found on a sheet.`);
    if (skipped > 0) parts.push(`${skipped} left as ${skipped === 1 ? "it was" : "they were"} to keep other cells.`);
    if (failed > 0) parts.push(`${failed} failed.`);
    setNote(parts.join(" "));
    if (firstError) setError(firstError);
    autoRefresh.noteRefreshed();
    return "done";
  };

  refreshAllRef.current = refreshAll;

  // "Refresh Every": Refresh All when due, never while the pane is busy with something else.
  const autoRefresh = useAutoRefresh({
    active: phase === "ready",
    minutes: refreshMinutes,
    canRun: () =>
      busy === null &&
      question === null &&
      picker === null &&
      !dirty &&
      detailsFlight.current === null &&
      !runner.busy,
    refresh: () => refreshAll(true),
  });

  // A right-click command (Expand, Collapse, ...) saved a definition from outside the
  // pane: take it, so that the pane never writes its older copy back over it. The report
  // being edited is reopened with it, unless the pane has a change of its own waiting.
  useWorkbookSync({
    host,
    active: phase === "ready",
    paneBusy: () => runner.busy,
    lastOwnWrite: () => lastOwnWrite.current,
    apply: (loaded) => {
      const cur = currentRef.current;
      const last = ranRef.current;
      const pending =
        runner.busy ||
        (cur !== null && cur.revision > 0 && !(last !== null && last.id === cur.def.id && last.revision === cur.revision));
      const result = syncReports({
        loaded,
        known: reportsRef.current,
        current: cur && cur.persisted ? cur.def : null,
        pending,
      });
      if (result.reports === reportsRef.current) return;
      // Their cells were rewritten elsewhere: what the runner remembers no longer fits,
      // and their size may have changed.
      for (const id of result.changed) runner.forget(id);
      if (result.changed.length > 0) {
        void refreshLocations();
        refreshSelection();
      }
      setReports(result.reports);
      if (result.reopen) {
        dispatch({ type: "open", def: result.reopen, persisted: true, anchor: cur ? cur.anchor : null });
        setRan(null);
      }
      setTick((n) => n + 1);
    },
  });

  const askDeleteReport = async (target: ReportDef) => {
    let extent: Extent | null = null;
    try {
      extent = await liveExtent(target.id);
    } catch {
      // Excel is busy: ask as if it had no cells; deleting still removes the definition.
    }
    setQuestion({ kind: "deleteReport", def: target, onSheet: extent !== null });
  };

  const deleteReport = async (target: ReportDef, clearCells: boolean) => {
    setQuestion(null);
    try {
      if (clearCells) await host.clearReport(target.id);
      await host.deleteReport(target.id);
    } catch (failure) {
      setError(toUiError(failure));
      return;
    }
    const charted = runner.chartsOf(target).length > 0; // they stay, as a PivotChart's last picture does
    runner.forget(target.id);
    const rest = reportsRef.current.filter((d) => d.id !== target.id);
    setReports(rest);
    setLocations((map) => {
      const next = { ...map };
      delete next[target.id];
      return next;
    });
    refreshSelection();
    const cur = currentRef.current;
    if (cur && cur.def.id === target.id) {
      if (rest.length > 0) openReport(rest[0]);
      else startDraft(target.connectionId, target.catalog, target.cube);
    }
    setNote(
      (clearCells || isEmptyLayout(target.layout)
        ? "Report deleted and its cells cleared."
        : "Report deleted. Its numbers stay as plain cells.") + (charted ? CHARTS_STAY : ""),
    );
  };

  /** Convert to Formulas, step 1: say why not, or ask (Windows warns it cannot be undone). */
  const askConvert = (target: ReportDef) => {
    const blocker = conversionBlocker(target.layout);
    if (blocker) {
      setNote(null);
      setError({ message: blocker });
      return;
    }
    setQuestion({ kind: "convert", def: target });
  };

  /**
   * Convert to Formulas, step 2: the report's cells become cube formulas, and the report
   * itself is removed (its definition and name; the cells stay). Works from the pane's
   * memory of the last result, so a report not shown in this session is refreshed first.
   */
  const convertToFormulas = async (target: ReportDef) => {
    setQuestion(null);
    setError(null);
    const conn = connectionsRef.current.find((c) => c.id === target.connectionId);
    const connection = conn ? formulaConnection(conn.name, target.catalog, target.cube) : null;
    if (!conn || !connection) {
      setError({
        message: conn
          ? "The database or cube name contains \"/\", which a cube formula cannot name. This report cannot be converted."
          : "Choose a connection for this report first.",
      });
      return;
    }
    let extent: Extent | null;
    try {
      extent = await liveExtent(target.id);
    } catch (failure) {
      setError(toUiError(failure));
      return;
    }
    if (!extent) {
      setNote("This report is not on a sheet, so there is nothing to convert.");
      return;
    }
    let runtime = runner.runtimes.get(target.id) ?? null;
    if (!runtime || !runtime.sheet || runtime.rows !== extent.rows || runtime.cols !== extent.cols) {
      const cur = currentRef.current;
      if (!cur || cur.def.id !== target.id) return;
      const outcome = await execute(cur.def, cur.revision, { anchor: null, useSelection: false });
      if (!outcome || outcome.kind !== "written") return; // the run has said why
      runtime = outcome.runtime;
    }
    const result = convertReport(runtime, connection);
    if ("refused" in result) {
      setError({ message: result.refused });
      return;
    }
    setBusy("Writing formulas…");
    try {
      await host.writeFormulas(target.id, result.formulas, result.numberFormats);
      await host.deleteReport(target.id);
    } catch (failure) {
      setError(toUiError(failure));
      return;
    } finally {
      setBusy(runner.busy ? "Asking the cube…" : null);
    }
    const charted = runner.chartsOf(target).length > 0;
    runner.forget(target.id);
    const rest = reportsRef.current.filter((d) => d.id !== target.id);
    setReports(rest);
    setLocations((map) => {
      const next = { ...map };
      delete next[target.id];
      return next;
    });
    refreshSelection();
    const cur = currentRef.current;
    if (cur && cur.def.id === target.id) {
      if (rest.length > 0) openReport(rest[0]);
      else startDraft(target.connectionId, target.catalog, target.cube);
    }
    setNote(
      "Converted to cube formulas. They read the cube again whenever Excel recalculates (Cmd+Option+F9, or Refresh All here)." +
        (result.filtered ? " Totals now count every item, including those the report's filters hid." : "") +
        (charted ? CHARTS_STAY : ""),
    );
  };

  /** Recalculate the workbook so that cube formulas ask the cube again. False when it failed. */
  const recalculateFormulas = async (): Promise<boolean> => {
    if (host.kind !== "excel") return false;
    try {
      await host.recalculate();
      return true;
    } catch {
      return false;
    }
  };

  // -- what to show -------------------------------------------------------------

  const runtime = def ? runner.runtimes.get(def.id) ?? null : null;
  const location = def ? locations[def.id] ?? null : null;

  // What the selected cell offers: drill for a member, "Filter <field>…" for a field.
  let drill: Drill = { state: "none" };
  let filterField: PlacedField | null = null;
  let filterLabels: "rows" | "columns" | null = null;
  let detailsTarget: ReturnType<ReportRunner["detailsFor"]> = null;
  if (def && selection && selection.report && selection.report.id === def.id) {
    const at = selection.report;
    detailsTarget = runner.detailsFor(def.id, at);
    let onSheet = runner.fieldAt(def.id, at);
    if (onSheet === undefined) {
      drill = isEmptyLayout(def.layout) ? { state: "none" } : { state: "stale" };
      // Even with no result in memory, the filter rows are known to head the block.
      const saved = reports.find((d) => d.id === def.id) ?? def;
      const guess = guessFieldCell(saved.layout, at.row, at.col);
      onSheet = guess && guess.area === "filters" ? saved.layout.filters[guess.index] ?? null : null;
    } else {
      const hit = runner.memberAt(def.id, at.row, at.col);
      if (hit) drill = { state: "member", hit };
      filterLabels = runner.labelsAt(def.id, at);
    }
    // Offer it only if the field is still in the layout being edited.
    if (onSheet) filterField = filterLabels ? def.layout[filterLabels][0] ?? null : findField(def.layout, onSheet.hierarchy);
  }

  // The selected cell is in a report written elsewhere (a right-click command, or before
  // the pane opened): ask the cube for it again, quietly, rather than for a Refresh.
  const recallId = drill.state === "stale" && def && !runner.runtimes.has(def.id) ? def.id : null;
  useEffect(() => {
    if (recallId === null || busy !== null) return;
    const saved = reportsRef.current.find((d) => d.id === recallId);
    if (!saved) return;
    void runner.recall(saved).then((ok) => {
      if (!ok) return;
      setTick((n) => n + 1);
      refreshSelection();
    });
  }, [recallId, busy, runner, reportsRef, refreshSelection]);
  // "Row Labels" with several fields: Windows names the button, and each field's selection.
  const labelsFields = def && filterLabels ? def.layout[filterLabels] : [];
  const filterSummary =
    filterField && filterLabels && labelsFields.length > 1
      ? {
          caption: labelsCaption(filterLabels),
          summary: labelsFields.map((f) => `${f.caption}: ${selectionSummary(f)}`).join("; "),
        }
      : filterField
        ? { caption: filterField.caption, summary: selectionSummary(filterField) }
        : null;

  const expand = (hit: MemberHit) =>
    act({ type: "expand", hierarchy: hit.member.hierarchy, member: hit.member.uniqueName }, true);
  const collapse = (hit: MemberHit) => {
    if (!def) return;
    act(
      {
        type: "collapse",
        hierarchy: hit.member.hierarchy,
        member: hit.member.uniqueName,
        descendants: runner.descendants(def.id, hit.axis, hit.member),
      },
      true,
    );
  };

  const dirty =
    current !== null &&
    current.revision > 0 &&
    !(ran !== null && ran.id === current.def.id && ran.revision === current.revision);
  const pickerField = def && picker ? findField(def.layout, picker) : null;
  /** The cube's measures and the report's calculated measures: what sorts and filters can use. */
  const allMeasures = [
    ...(fields.data ? fields.data.measures : []),
    ...(def?.layout.calculatedMeasures ?? []).map((m) => ({
      uniqueName: calculatedMeasureName(m.name),
      caption: m.name,
      group: "",
      folder: "",
    })),
  ];
  const isDateField = (hierarchy: string): boolean => {
    const placed = def ? findField(def.layout, hierarchy) : null;
    const found = fields.data
      ? fields.data.dimensions.flatMap((d) => d.hierarchies).find((h) => h.uniqueName === hierarchy)
      : undefined;
    if (!found) return false;
    // A field placed as one level is a date field when that level holds dates.
    if (placed && placed.level) {
      return Boolean((found.levels ?? []).find((l) => l.uniqueName === placed.level?.uniqueName)?.isDate);
    }
    return Boolean(found.isDate);
  };
  const pickerPlace = def && picker ? locate(def.layout, picker) : null;
  const pickerArea = pickerPlace ? pickerPlace.area : null;

  const reportLabel = (d: ReportDef): string => {
    const where = locations[d.id];
    return (where ? describeCell(where) : "Not on a sheet") + (d.cube ? ` · ${d.cube}` : "");
  };
  const draftLabel = current && !current.persisted
    ? current.anchor
      ? `New report · ${describeCell(current.anchor)}`
      : "New report · at the selected cell"
    : "";

  // Insert Chart and Delete Chart (components/chartMenu.ts).
  const chartActions: ChartActions = {
    runner,
    refresh: (target) => {
      const cur = currentRef.current;
      if (!cur || cur.def.id !== target.id) return Promise.resolve(null);
      return execute(cur.def, cur.revision, { anchor: null, useSelection: false });
    },
    setBusy: (text) => setBusy(text ?? (runner.busy ? "Asking the cube…" : null)),
    setError,
    setNote,
    saved: (saved) => {
      setReports((list) => upsert(list, saved));
      dispatch({ type: "charts", id: saved.id, charts: saved.charts ?? [] });
    },
  };

  const reportMenuItems = (): MenuEntry[] => {
    const items: MenuEntry[] = [];
    if (current && current.persisted) {
      items.push({
        label: "Move this report here",
        disabled: !location || busy !== null,
        run: () => void moveHere(),
      });
      items.push("separator");
    }
    items.push(
      { label: "Insert Slicer…", disabled: !fields.data || !connection, run: () => setInsertSlicer(true) },
      "separator",
    );
    items.push({
      label: "Reload field list",
      run: () => {
        cache.clear();
        setGeneration((n) => n + 1);
      },
    });
    items.push({
      label: "Refresh when the workbook opens",
      checked: refreshOnOpen,
      run: () => {
        const on = !refreshOnOpen;
        host.setFlag(REFRESH_ON_OPEN, on).then(
          () => {
            setRefreshOnOpen(on);
            setNote(
              on
                ? "Every report in this workbook will refresh when the workbook opens."
                : "Reports will no longer refresh when the workbook opens.",
            );
          },
          (failure) => setError(toUiError(failure)),
        );
      },
    });
    items.push(
      refreshEveryMenu(refreshMinutes, (minutes) => {
        host.setNumber(REFRESH_MINUTES, minutes > 0 ? minutes : null).then(
          () => {
            setRefreshMinutes(minutes);
            setNote(refreshEveryNote(minutes));
          },
          (failure) => setError(toUiError(failure)),
        );
      }),
    );
    if (current && current.persisted) {
      const target = current.def;
      const chartState = { busy: busy !== null || dirty, canInsert: location !== null && !isEmptyLayout(target.layout) };
      items.push("separator", ...chartMenuItems(chartActions, target, chartState));
      items.push("separator", {
        label: "Convert to Formulas…",
        disabled: !location || busy !== null || isEmptyLayout(target.layout),
        run: () => askConvert(target),
      });
      items.push({
        label: "Delete report…",
        danger: true,
        run: () => void askDeleteReport(target),
      });
    }
    return items;
  };

  const certDays = status ? status.certDaysLeft : null;

  let body: JSX.Element;
  if (phase === "loading") {
    body = (
      <div className="center">
        <Spinner label="Starting" /> Starting…
      </div>
    );
  } else if (phase !== "ready") {
    body = (
      <div className="screen">
        <header className="bar">
          <h1 className="bar-title">Cube Explorer</h1>
        </header>
        <div className="scroll">
          <ErrorNote error={phase} action={{ label: "Try again", run: () => void boot() }} />
        </div>
      </div>
    );
  } else if (screen === "connections" || !current || !def) {
    const orphan = current !== null && current.persisted && connection === null;
    body = (
      <>
        {orphan && <p className="banner">Choose a connection for this report.</p>}
        <Connections
          api={api}
          connections={connections}
          activeId={connectionId}
          onChoose={chooseConnection}
          onChanged={reloadConnections}
          onBack={current ? () => setScreen("explorer") : undefined}
        />
      </>
    );
  } else {
    const others = reports.filter((d) => d.id !== def.id);
    const showSwitch = others.length > 0;
    const lost = current.persisted && !location;
    const form = reportForm(def.layout.options);
    body = (
      <div className="screen explorer">
        <header className="bar">
          <button
            type="button"
            className="icon-button"
            onClick={() => setScreen("connections")}
            aria-label="Connections"
            title="Connections"
          >
            <Icon name="chevron-left" />
          </button>
          <h1 className="bar-title" title={connection ? connection.name : undefined}>
            {connection ? connection.name : "No connection"}
          </h1>
          <button
            type="button"
            className="icon-button"
            aria-label="Report menu"
            aria-haspopup="menu"
            title="More"
            onClick={(event) => setReportMenu(reportMenu ? null : event.currentTarget)}
          >
            <Icon name="more" />
          </button>
        </header>

        {certDays !== null && certDays <= CERT_WARNING_DAYS && (
          <p className="banner">
            {certDays <= 0
              ? "The helper's security certificate has expired."
              : `The helper's security certificate expires in ${plural(certDays, "day", "days")}.`}{" "}
            Run install.sh again to renew it.
          </p>
        )}

        <div className="report-bar">
          <span className="row-label">Report</span>
          {showSwitch ? (
            <select
              className="select"
              aria-label="Report"
              value={def.id}
              onChange={(event) => {
                const next = reports.find((d) => d.id === event.target.value);
                if (next) openReport(next);
              }}
            >
              {!current.persisted && <option value={def.id}>{draftLabel}</option>}
              {reports.map((d) => (
                <option key={d.id} value={d.id}>
                  {reportLabel(d)}
                </option>
              ))}
            </select>
          ) : (
            <span className="report-where" title={current.persisted ? reportLabel(def) : draftLabel}>
              {current.persisted ? reportLabel(def) : draftLabel}
            </span>
          )}
          <button
            type="button"
            className="button small"
            onClick={() => void placeHere()}
            disabled={busy !== null}
            title={
              location
                ? "Start another report at the selected cell"
                : "Put this report, with its fields, at the selected cell"
            }
          >
            {placeLabel(location)}
          </button>
        </div>

        {lost && (
          <p className="hint-line" role="status">
            This report has no cells on a sheet (its sheet may have been deleted). Change a field or
            click Refresh to put it, with its fields, at the selected cell.
          </p>
        )}

        {!connection && (
          <div className="rebind" role="alert">
            <p>Choose a connection for this report. It was made with one that is not on this Mac.</p>
            <div className="rebind-row">
              <select
                className="select"
                aria-label="Connection for this report"
                value=""
                onChange={(event) => {
                  if (event.target.value !== "") rebind(event.target.value);
                }}
              >
                <option value="">Choose a connection</option>
                {connections.map((c) => (
                  <option key={c.id} value={c.id}>
                    {c.name}
                  </option>
                ))}
              </select>
              <button type="button" className="button small" onClick={() => setScreen("connections")}>
                Add…
              </button>
            </div>
          </div>
        )}

        <CubePicker
          catalogs={catalogs}
          cubes={cubes}
          catalog={catalog}
          cube={cube}
          disabled={!connection || busy !== null}
          onCatalog={(value) => changeCube("database", value)}
          onCube={(value) => changeCube("cube", value)}
        />

        <FieldList
          fields={fields.data}
          loading={fields.loading}
          error={fields.error}
          onRetry={fields.retry}
          placeholder={
            !connection
              ? "Fields appear here once the report has a connection."
              : "Choose a database and a cube to see its fields."
          }
          layout={def.layout}
          search={search}
          onSearch={setSearch}
          onToggleMeasure={(measure, on) =>
            act(on ? { type: "addMeasure", measure } : { type: "removeMeasure", measure: measure.uniqueName })
          }
          onToggleHierarchy={(hierarchy, on) =>
            act(on ? { type: "addHierarchy", hierarchy } : { type: "remove", name: hierarchy.uniqueName })
          }
          onToggleKpi={(kpi, part, on) =>
            act(on ? { type: "addKpi", kpi, part } : { type: "remove", name: `kpi:${kpi.name}:${part}` })
          }
          onToggleSet={(set, on) => {
            if (!set.hierarchy) return;
            act(
              on
                ? {
                    type: "addSet",
                    set: {
                      uniqueName: set.uniqueName,
                      caption: set.caption,
                      hierarchy: set.hierarchy,
                      allMember: set.allMember,
                      hierarchies: set.hierarchies,
                      allMembers: set.allMembers,
                    },
                  }
                : { type: "remove", name: set.hierarchy },
            );
          }}
          onToggleLevel={(hierarchy, level, on) =>
            act(
              on
                ? {
                    type: "addHierarchy",
                    hierarchy: {
                      uniqueName: hierarchy.uniqueName,
                      caption: hierarchy.caption,
                      allMember: hierarchy.allMember,
                      level: { uniqueName: level.uniqueName, number: level.number, caption: level.caption },
                    },
                  }
                : { type: "remove", name: hierarchy.uniqueName },
            )
          }
          onToggleCalculated={(measure, on) => {
            const name = calculatedMeasureName(measure.name);
            act(on ? { type: "addMeasure", measure: { uniqueName: name, caption: measure.name } } : { type: "removeMeasure", measure: name });
          }}
          onEditCalculated={(measure) => setQuestion({ kind: "calculatedMeasure", measure })}
        />

        <Areas
          layout={def.layout}
          dispatch={act}
          onFilter={setPicker}
          onShowAsForm={(measure, showAs) => setQuestion({ kind: "showAs", measure, showAs })}
          onCustomFormat={(measure) => {
            const value = def.layout.values[valueIndex(def.layout.values, measure)];
            if (value) {
              setQuestion({
                kind: "numberFormat",
                measure,
                caption: value.caption,
                current: value.numberFormat ?? "",
              });
            }
          }}
        />

        {def.layout.slicers && def.layout.slicers.length > 0 && (
          <Slicers api={api} connectionId={connectionId} catalog={catalog} cube={cube} slicers={def.layout.slicers} dispatch={act} />
        )}

        <div className="options">
          <label className="row-field options-wide">
            <span className="row-label">Subtotals</span>
            <select
              className="select"
              value={subtotalsChoice(def.layout.options)}
              onChange={(event) => act({ type: "setSubtotals", at: event.target.value as SubtotalsChoice })}
            >
              {SUBTOTALS.map((s) => (
                // Tabular form has no subtotals at the top of a group; Excel greys the choice out.
                <option key={s.value} value={s.value} disabled={s.value === "top" && form === "tabular"}>
                  {s.label}
                </option>
              ))}
            </select>
          </label>
          <label className="row-field options-wide">
            <span className="row-label">Grand totals</span>
            <select
              className="select"
              value={grandTotalsChoice(def.layout.options)}
              onChange={(event) => {
                const chosen = GRAND_TOTALS.find((g) => g.value === event.target.value);
                if (chosen) act({ type: "setGrandTotals", row: chosen.row, column: chosen.column });
              }}
            >
              {GRAND_TOTALS.map((g) => (
                <option key={g.value} value={g.value}>
                  {g.label}
                </option>
              ))}
            </select>
          </label>
          <label className="row-field options-wide">
            <span className="row-label">Report layout</span>
            <select
              className="select"
              value={form}
              onChange={(event) => act({ type: "setForm", form: event.target.value as ReportForm })}
            >
              {REPORT_LAYOUTS.map((l) => (
                <option key={l.value} value={l.value}>
                  {l.label}
                </option>
              ))}
            </select>
          </label>
          <StyleOptions
            options={def.layout.options}
            onStyle={(style) => act({ type: "setStyle", style })}
            onOption={(option, value) => act({ type: "setStyleOption", option, value })}
          />
          {OPTION_LABELS.map(([option, label]) => {
            // Compact form writes every row field in one column: nothing to repeat (Excel greys it out).
            const off = option === "repeatLabels" && form === "compact";
            return (
              <label key={option} className={off ? "check is-disabled" : "check"}>
                <input
                  type="checkbox"
                  checked={Boolean(def.layout.options[option])}
                  disabled={off}
                  onChange={(event) => act({ type: "setOption", option, value: event.target.checked })}
                />
                {label}
              </label>
            );
          })}
          <label className="check">
            <input type="checkbox" checked={defer} onChange={(event) => setDefer(event.target.checked)} />
            Defer updates
          </label>
          {defer && (
            <button
              type="button"
              className="button small options-update"
              disabled={!dirty || busy !== null}
              onClick={() => dispatch({ type: "touch" })}
            >
              Update
            </button>
          )}
          <label className="row-field options-wide">
            <span className="row-label options-label">For empty cells show</span>
            <EmptyCellText
              value={def.layout.options.emptyCellText ?? ""}
              onCommit={(text) => act({ type: "setEmptyCellText", text })}
            />
          </label>
        </div>

        <StatusBar
          busy={busy}
          onCancel={() => {
            const flight = detailsFlight.current;
            if (flight) {
              flight.abort.abort();
              api.cancel(flight.connectionId).catch(() => undefined);
            } else {
              runner.cancel();
            }
          }}
          error={error}
          onDismissError={() => setError(null)}
          onRetry={error && error.helperDown ? () => dispatch({ type: "touch" }) : undefined}
          note={note}
          runtime={runtime}
          drill={drill}
          filter={connection ? filterSummary : null}
          onFilter={() => {
            if (!filterField) return;
            setPickerFrom(filterLabels);
            setPicker(filterField.hierarchy);
          }}
          details={Boolean(detailsTarget) && connection !== null}
          onDetails={() => void showDetails()}
          onExpand={expand}
          onCollapse={collapse}
          onShowQuery={() => runtime && setQuestion({ kind: "query", mdx: runtime.stats.mdx })}
          onRefresh={() => dispatch({ type: "touch" })}
          onRefreshAll={() => void refreshAll()}
          canRefresh={connection !== null && catalog !== "" && cube !== ""}
          canRefreshAll={reports.length > 0 || host.kind === "excel" /* cube formulas, even with no report */}
        />

        {reportMenu && (
          <Menu anchor={reportMenu} label="Report" items={reportMenuItems()} onClose={() => setReportMenu(null)} />
        )}
      </div>
    );
  }

  return (
    <div className={host.kind === "preview" ? "app with-preview" : "app"}>
      <div className="pane">
        {mock && <p className="strip">Sample data only. Not connected to the helper.</p>}
        {body}

        {insertSlicer && def && fields.data && (
          <InsertSlicer
            key={def.id}
            fields={fields.data}
            layout={def.layout}
            onCancel={() => setInsertSlicer(false)}
            onApply={(hierarchies) => {
              setInsertSlicer(false);
              act({ type: "setSlicers", hierarchies });
            }}
          />
        )}

        {pickerField && pickerField.set && (
          <Dialog title={pickerField.caption} onCancel={() => setPicker(null)} buttons={[{ label: "OK", primary: true, run: () => setPicker(null) }]}>
            <p>
              &ldquo;{pickerField.caption}&rdquo; is a set defined in the cube. Its items are the set&rsquo;s own, so they
              cannot be filtered or sorted here, as in Excel.
            </p>
          </Dialog>
        )}

        {pickerField && !pickerField.set && def && connection && (
          <MemberPicker
            key={pickerField.hierarchy}
            api={api}
            connectionId={connection.id}
            catalog={def.catalog}
            cube={def.cube}
            field={pickerField}
            choose={
              pickerFrom && pickerArea === pickerFrom
                ? { fields: def.layout[pickerFrom], onChoose: (hierarchy) => setPicker(hierarchy) }
                : undefined
            }
            onCancel={() => setPicker(null)}
            onApply={(members, captions) => {
              setPicker(null);
              act({ type: "setMembers", hierarchy: pickerField.hierarchy, members, captions });
            }}
            menu={
              pickerArea === "rows" || pickerArea === "columns"
                ? {
                    measures: allMeasures,
                    defaultMeasure:
                      def.layout.values.length > 0
                        ? def.layout.values[0].measure
                        : fields.data && fields.data.measures.length > 0
                          ? fields.data.measures[0].uniqueName
                          : "",
                    expandable: runner.expandable(def.id, pickerField.hierarchy),
                    onAction: (action) => {
                      setPicker(null);
                      act(action);
                    },
                    isDate: isDateField(pickerField.hierarchy),
                    onGroups: () => {
                      setPicker(null);
                      setQuestion({ kind: "groups", hierarchy: pickerField.hierarchy });
                    },
                    onCalculatedItems: () => {
                      setPicker(null);
                      setQuestion({ kind: "calculatedItems", hierarchy: pickerField.hierarchy });
                    },
                  }
                : undefined
            }
          />
        )}

        {question && question.kind === "groups" && def && connection && (() => {
          const field = findField(def.layout, question.hierarchy);
          if (!field) return null;
          return (
            <GroupDialog
              api={api}
              connectionId={connection.id}
              catalog={def.catalog}
              cube={def.cube}
              field={field}
              onCancel={() => setQuestion(null)}
              onApply={(items) => {
                setQuestion(null);
                act({ type: "setCustomItems", hierarchy: field.hierarchy, items });
              }}
            />
          );
        })()}

        {question && question.kind === "calculatedItems" && def && connection && (() => {
          const field = findField(def.layout, question.hierarchy);
          if (!field) return null;
          return (
            <CalculatedItemDialog
              api={api}
              connectionId={connection.id}
              catalog={def.catalog}
              cube={def.cube}
              field={field}
              value={def.layout.values.find((v) => !v.kpi && (v.showAs ?? "normal") === "normal") ?? null}
              onCancel={() => setQuestion(null)}
              onApply={(items) => {
                setQuestion(null);
                act({ type: "setCustomItems", hierarchy: field.hierarchy, items });
              }}
            />
          );
        })()}

        {question && question.kind === "calculatedMeasure" && def && connection && (
          <CalculatedMeasureDialog
            api={api}
            connectionId={connection.id}
            catalog={def.catalog}
            cube={def.cube}
            measures={fields.data ? fields.data.measures : []}
            existing={def.layout.calculatedMeasures ?? []}
            editing={question.measure}
            onCancel={() => setQuestion(null)}
            onSave={(measure, previousName) => {
              setQuestion(null);
              act({ type: "setCalculatedMeasure", measure, previousName });
            }}
            onDelete={(name) => {
              setQuestion(null);
              act({ type: "removeCalculatedMeasure", name });
            }}
          />
        )}

        {question && question.kind === "showAs" && def && connection && (() => {
          const value = def.layout.values[valueIndex(def.layout.values, question.measure)];
          if (!value) return null;
          return (
            <ShowAsDialog
              api={api}
              connectionId={connection.id}
              catalog={def.catalog}
              cube={def.cube}
              value={value}
              showAs={question.showAs}
              axisFields={[...def.layout.rows, ...def.layout.columns]}
              onCancel={() => setQuestion(null)}
              onApply={(choice) => {
                setQuestion(null);
                act({ type: "setShowAs", measure: question.measure, ...choice });
              }}
            />
          );
        })()}

        {question && question.kind === "numberFormat" && (
          <NumberFormatDialog
            caption={question.caption}
            initial={question.current}
            onCancel={() => setQuestion(null)}
            onApply={(format) => {
              setQuestion(null);
              act({ type: "setNumberFormat", measure: question.measure, format });
            }}
          />
        )}

        {question && question.kind === "overwrite" && (
          <Dialog
            title="Replace other cells?"
            onCancel={() => question.answer(false)}
            buttons={[
              { label: "Cancel", run: () => question.answer(false), focus: true },
              { label: "Replace", run: () => question.answer(true), danger: true },
            ]}
          >
            <p>
              The report now needs {question.target.sheetName}!{rectA1(question.target)}. Some of those
              cells already hold something that is not part of this report.
            </p>
            <p>Replacing them cannot be undone from this pane.</p>
          </Dialog>
        )}

        {question && question.kind === "deleteReport" && (
          <Dialog
            title="Delete this report?"
            onCancel={() => setQuestion(null)}
            buttons={
              question.onSheet && !isEmptyLayout(question.def.layout)
                ? [
                    { label: "Cancel", run: () => setQuestion(null), focus: true },
                    { label: "Keep values", run: () => void deleteReport(question.def, false) },
                    { label: "Delete and clear", run: () => void deleteReport(question.def, true), danger: true },
                  ]
                : [
                    { label: "Cancel", run: () => setQuestion(null), focus: true },
                    {
                      label: "Delete",
                      danger: true,
                      // With no fields, what is on the sheet is only its placeholder.
                      run: () => void deleteReport(question.def, question.onSheet),
                    },
                  ]
            }
          >
            <p>
              {question.onSheet && !isEmptyLayout(question.def.layout)
                ? "It will no longer refresh. You can keep its numbers as plain cells, or clear them from the sheet."
                : "It will be removed from this workbook."}
            </p>
          </Dialog>
        )}

        {question && question.kind === "convert" && (
          <Dialog
            title="Convert to Formulas?"
            onCancel={() => setQuestion(null)}
            buttons={[
              { label: "Cancel", run: () => setQuestion(null), focus: true },
              { label: "Convert", run: () => void convertToFormulas(question.def), danger: true },
            ]}
          >
            <p>
              Each cell of this report becomes a cube formula (CUBE.MEMBER, CUBE.VALUE) that you can move,
              copy and edit. The formulas keep their numbers up to date when Excel recalculates.
            </p>
            <p>The report itself is removed and cannot be changed from this pane again. This cannot be undone.</p>
          </Dialog>
        )}

        {question && question.kind === "reset" && (
          <Dialog
            title={`Change the ${question.what}?`}
            onCancel={() => setQuestion(null)}
            buttons={[
              { label: "Cancel", run: () => setQuestion(null), focus: true },
              {
                label: "Change",
                danger: true,
                run: () => {
                  const { apply } = question;
                  setQuestion(null);
                  apply();
                },
              },
            ]}
          >
            <p>
              The fields in this report belong to the current cube. Changing the {question.what} removes
              them and empties the report on the sheet.
            </p>
          </Dialog>
        )}

        {question && question.kind === "query" && (
          <Dialog
            title="Query sent to the cube"
            wide
            onCancel={() => setQuestion(null)}
            buttons={[{ label: "Close", run: () => setQuestion(null) }]}
          >
            <textarea className="query-box" readOnly value={question.mdx} rows={12} aria-label="MDX query" />
            <CopyButton text={question.mdx} />
          </Dialog>
        )}
      </div>
      {host instanceof PreviewHost && <PreviewSheet host={host} />}
    </div>
  );
}
