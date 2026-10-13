// The task pane. Two screens: Connections, and the Explorer for the report being
// edited. All sheet access goes through `host`, all cube access through `api`; this
// file only wires user actions to the reducers and the runner, and decides what to
// show.

import { useCallback, useEffect, useMemo, useReducer, useRef, useState } from "react";
import type { Api } from "./api";
import { Areas } from "./components/Areas";
import { Connections } from "./components/Connections";
import { CubePicker } from "./components/CubePicker";
import { FieldList } from "./components/FieldList";
import { MemberPicker } from "./components/MemberPicker";
import { Menu, type MenuEntry } from "./components/Menu";
import { PreviewSheet } from "./components/PreviewSheet";
import { StatusBar, type Drill } from "./components/StatusBar";
import { CopyButton, Dialog, ErrorNote, Icon, Spinner } from "./components/common";
import { rectA1 } from "./sheet/geometry";
import {
  describeCell,
  type CellAddress,
  type Extent,
  type ReportDef,
  type SheetHost,
  type SheetSelection,
} from "./sheet/host";
import type { MemberHit } from "./sheet/layout";
import { PreviewHost } from "./sheet/preview";
import { findField, isEmptyLayout, type LayoutAction } from "./state/layout";
import { currentReducer, newReportDef } from "./state/report";
import { ReportRunner, toUiError, type RunOutcome, type UiError } from "./state/runner";
import { useCached, type Cache } from "./state/useCached";
import type { Catalog, Connection, Cube, Fields, LayoutOptions, Status } from "./types";
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
  | { kind: "deleteReport"; def: ReportDef; hasCells: boolean }
  | { kind: "reset"; what: "database" | "cube"; apply: () => void }
  | { kind: "query"; mdx: string };

/** Pause between a layout change and its query, so that ticking three fields is one query. */
const DEBOUNCE_MS = 400;
/** Warn about the helper's certificate when it has this many days left, or fewer. */
const CERT_WARNING_DAYS = 30;

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

const OPTION_LABELS: [keyof LayoutOptions, string][] = [
  ["grandTotals", "Grand totals"],
  ["subtotals", "Subtotals"],
  ["nonEmpty", "Hide empty"],
];

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
  const [search, setSearch] = useState("");
  const [reportMenu, setReportMenu] = useState<HTMLElement | null>(null);
  const [generation, setGeneration] = useState(0);
  const [, setTick] = useState(0);

  const currentRef = useLatest(current);
  const reportsRef = useLatest(reports);
  const connectionsRef = useLatest(connections);
  const locationsRef = useLatest(locations);
  const ranRef = useLatest(ran);
  const busyPrefix = useRef("");
  const cache = useRef<Cache>(new Map()).current;

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

  /** Fold the result of a run into the pane's state. `single` runs report their own notes. */
  const applyOutcome = useCallback(
    (def: ReportDef, outcome: RunOutcome, single: boolean) => {
      const isCurrent = () => {
        const cur = currentRef.current;
        return cur !== null && cur.def.id === def.id;
      };
      switch (outcome.kind) {
        case "written":
          setReports((list) => upsert(list, def));
          setLocations((map) => ({ ...map, [def.id]: outcome.extent }));
          if (isCurrent()) dispatch({ type: "saved" });
          refreshSelection();
          break;
        case "empty":
          if (outcome.cleared) {
            setReports((list) => upsert(list, def));
            host.getReportExtent(def.id).then(
              (extent) => setLocations((map) => ({ ...map, [def.id]: extent })),
              () => undefined,
            );
            refreshSelection();
          }
          break;
        case "cancelled":
          if (single && isCurrent()) setNote("Cancelled. The sheet was not changed.");
          break;
        case "declined":
          if (single && isCurrent()) setNote("Not updated, so that the other cells stay as they are.");
          break;
        case "needs-anchor":
          if (single && isCurrent()) {
            setNote(
              currentRef.current && currentRef.current.persisted
                ? "The cells of this report can no longer be found. Select a cell and choose Place here."
                : "Select an empty cell outside other reports, then choose Insert here.",
            );
          }
          break;
        case "error":
          if (single && isCurrent()) setError(outcome.error);
          break;
        case "superseded":
          break;
      }
      setTick((n) => n + 1);
    },
    [currentRef, host, refreshSelection],
  );

  const runCurrent = useCallback(async () => {
    const cur = currentRef.current;
    if (!cur) return;
    const { def, revision } = cur;
    const done = () => setRan({ id: def.id, revision });
    // No connection on this Mac: the banner asks for one. No cube yet: nothing to ask.
    if (!connectionsRef.current.some((c) => c.id === def.connectionId) || def.catalog === "" || def.cube === "") {
      done();
      return;
    }
    setError(null);
    setNote(null);
    busyPrefix.current = "";
    setBusy("Asking the cube…");
    const outcome = await runner.run(def, { anchor: cur.anchor, useSelection: !cur.persisted });
    if (outcome.kind !== "superseded") {
      if (!runner.busy) setBusy(null);
      done();
    }
    applyOutcome(def, outcome, true);
  }, [applyOutcome, connectionsRef, currentRef, runner]);
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
    (connectionId: string, catalog = "", cube = "", anchor: CellAddress | null = null) => {
      dispatch({ type: "open", def: newReportDef(connectionId, catalog, cube), persisted: false, anchor });
      setScreen("explorer");
      setError(null);
      setNote(null);
      setPicker(null);
      setSearch("");
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
    try {
      defs = await host.loadReports();
      for (const def of defs) where[def.id] = await host.getReportExtent(def.id);
      active = await host.getSelection();
    } catch {
      // The sheet cannot be read right now (a cell is being edited, say). The pane
      // still opens; reports show up at the next start.
    }
    setReports(defs);
    setLocations(where);
    setSelection(active);
    const hit = active && active.report ? active.report.id : null;
    const first = defs.find((d) => d.id === hit) ?? defs[0];
    if (first) openReport(first);
    else setScreen("connections");
    setPhase("ready");
  }, [api, host, openReport]);

  useEffect(() => {
    void boot();
  }, [boot]);

  // Selecting a cell inside a report makes that report the one being edited.
  useEffect(() => {
    if (phase !== "ready") return;
    return host.onSelectionChanged((next) => {
      setSelection(next);
      if (!next.report) return;
      const cur = currentRef.current;
      if (cur && cur.def.id === next.report.id) return;
      const id = next.report.id;
      const def = reportsRef.current.find((d) => d.id === id);
      if (def) openReport(def);
    });
  }, [phase, host, currentRef, reportsRef, openReport]);

  const reloadConnections = useCallback(async () => {
    setConnections(await api.connections());
  }, [api]);

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

  const act = (action: LayoutAction, urgent = false) => dispatch({ type: "layout", action, urgent });

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
      () => setReports((list) => upsert(list, next)),
      (failure) => setError(toUiError(failure)),
    );
  };

  const changeCube = (what: "database" | "cube", value: string) => {
    const cur = currentRef.current;
    if (!cur) return;
    const apply = () =>
      dispatch(what === "database" ? { type: "setCatalog", catalog: value } : { type: "setCube", cube: value });
    const hasFields = !isEmptyLayout(cur.def.layout) || cur.def.layout.filters.length > 0;
    if (hasFields) setQuestion({ kind: "reset", what, apply });
    else apply();
  };

  const insertHere = async () => {
    const cur = currentRef.current;
    if (!cur) return;
    let active: SheetSelection;
    try {
      active = await host.getSelection();
    } catch (failure) {
      setError(toUiError(failure));
      return;
    }
    setSelection(active);
    setError(null);
    if (active.report) {
      setNote(
        active.report.id === cur.def.id
          ? "This cell is part of this report. Select a cell outside it to start another report."
          : "This cell is inside another report. Select a cell outside it first.",
      );
      return;
    }
    const where = describeCell(active.cell);
    if (locationsRef.current[cur.def.id]) {
      // The report being edited is on the sheet: start another one here, on the same cube.
      startDraft(cur.def.connectionId, cur.def.catalog, cur.def.cube, active.cell);
      setNote(`New report at ${where}. Tick a field to fill it.`);
      return;
    }
    // It has no cells yet (or lost them): this is where it goes.
    dispatch({ type: "pin", anchor: active.cell });
    if (isEmptyLayout(cur.def.layout)) setNote(`The report will start at ${where}. Tick a field to fill it.`);
    else dispatch({ type: "touch" });
  };

  const refreshAll = async () => {
    const cur = currentRef.current;
    const list = reportsRef.current.map((d) => (cur && cur.def.id === d.id ? cur.def : d));
    if (list.length === 0) return;
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
        return; // another run took over and owns the status line now
      }
      if (cur && cur.def.id === item.id) setRan({ id: item.id, revision: cur.revision });
      if (outcome.kind === "cancelled") {
        stopped = true;
        break;
      }
      if (outcome.kind === "written" || outcome.kind === "empty") done++;
      else if (outcome.kind === "needs-anchor") lost++;
      else if (outcome.kind === "declined") skipped++;
      else if (outcome.kind === "error") {
        failed++;
        if (!firstError) firstError = outcome.error;
        if (outcome.error.helperDown) break; // the rest would fail the same way
      }
    }
    busyPrefix.current = "";
    if (!runner.busy) setBusy(null);
    const parts = [stopped ? `Cancelled after ${plural(done, "report", "reports")}.` : `${plural(done, "report", "reports")} refreshed.`];
    if (unbound > 0) parts.push(`${unbound} ${unbound === 1 ? "needs" : "need"} a connection.`);
    if (lost > 0) parts.push(`${lost} could not be found on a sheet.`);
    if (skipped > 0) parts.push(`${skipped} left as ${skipped === 1 ? "it was" : "they were"} to keep other cells.`);
    if (failed > 0) parts.push(`${failed} failed.`);
    setNote(parts.join(" "));
    if (firstError) setError(firstError);
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
    setNote(clearCells ? "Report deleted and its cells cleared." : "Report deleted. Its numbers stay as plain cells.");
  };

  // -- what to show -------------------------------------------------------------

  const runtime = def ? runner.runtimes.get(def.id) ?? null : null;
  const location = def ? locations[def.id] ?? null : null;

  let drill: Drill = { state: "none" };
  if (def && selection && selection.report && selection.report.id === def.id) {
    const at = selection.report;
    if (!runtime || runtime.rows !== at.rows || runtime.cols !== at.cols) {
      drill = isEmptyLayout(def.layout) ? { state: "none" } : { state: "stale" };
    } else {
      const hit = runner.memberAt(def.id, at.row, at.col);
      if (hit) drill = { state: "member", hit };
    }
  }

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

  const reportLabel = (d: ReportDef): string => {
    const where = locations[d.id];
    return (where ? describeCell(where) : "Not on a sheet") + (d.cube ? ` · ${d.cube}` : "");
  };
  const draftLabel = current && !current.persisted
    ? current.anchor
      ? `New report · ${describeCell(current.anchor)}`
      : "New report · at the selected cell"
    : "";

  const reportMenuItems = (): MenuEntry[] => {
    const items: MenuEntry[] = [
      {
        label: "Reload field list",
        run: () => {
          cache.clear();
          setGeneration((n) => n + 1);
        },
      },
    ];
    if (current && current.persisted) {
      const target = current.def;
      items.push("separator", {
        label: "Delete report…",
        danger: true,
        run: () =>
          setQuestion({
            kind: "deleteReport",
            def: target,
            hasCells: Boolean(location) && !isEmptyLayout(target.layout),
          }),
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
          <button type="button" className="button small" onClick={() => void insertHere()} disabled={busy !== null}>
            {current.persisted && !location ? "Place here" : "Insert here"}
          </button>
        </div>

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
            act(on ? { type: "addMeasure", measure } : { type: "remove", name: measure.uniqueName })
          }
          onToggleHierarchy={(hierarchy, on) =>
            act(on ? { type: "addHierarchy", hierarchy } : { type: "remove", name: hierarchy.uniqueName })
          }
        />

        <Areas layout={def.layout} dispatch={act} onFilter={setPicker} />

        <div className="options">
          {OPTION_LABELS.map(([option, label]) => (
            <label key={option} className="check">
              <input
                type="checkbox"
                checked={def.layout.options[option]}
                onChange={(event) => act({ type: "setOption", option, value: event.target.checked })}
              />
              {label}
            </label>
          ))}
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
        </div>

        <StatusBar
          busy={busy}
          onCancel={() => runner.cancel()}
          error={error}
          onDismissError={() => setError(null)}
          onRetry={error && error.helperDown ? () => dispatch({ type: "touch" }) : undefined}
          note={note}
          runtime={runtime}
          drill={drill}
          onExpand={expand}
          onCollapse={collapse}
          onShowQuery={() => runtime && setQuestion({ kind: "query", mdx: runtime.stats.mdx })}
          onRefresh={() => dispatch({ type: "touch" })}
          onRefreshAll={() => void refreshAll()}
          canRefresh={connection !== null && catalog !== "" && cube !== ""}
          canRefreshAll={reports.length > 0}
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

        {pickerField && def && connection && (
          <MemberPicker
            key={pickerField.hierarchy}
            api={api}
            connectionId={connection.id}
            catalog={def.catalog}
            cube={def.cube}
            field={pickerField}
            onCancel={() => setPicker(null)}
            onApply={(members, captions) => {
              setPicker(null);
              act({ type: "setMembers", hierarchy: pickerField.hierarchy, members, captions });
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
              question.hasCells
                ? [
                    { label: "Cancel", run: () => setQuestion(null), focus: true },
                    { label: "Keep values", run: () => void deleteReport(question.def, false) },
                    { label: "Delete and clear", run: () => void deleteReport(question.def, true), danger: true },
                  ]
                : [
                    { label: "Cancel", run: () => setQuestion(null), focus: true },
                    { label: "Delete", run: () => void deleteReport(question.def, false), danger: true },
                  ]
            }
          >
            <p>
              {question.hasCells
                ? "It will no longer refresh. You can keep its numbers as plain cells, or clear them from the sheet."
                : "It will be removed from this workbook."}
            </p>
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
