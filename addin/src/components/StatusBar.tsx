// The bottom of the Explorer: what the selected cell can do (Expand / Collapse), the
// last error or note, and the status line with counts, elapsed time, Show query,
// Refresh and Refresh All. While a query runs it shows a spinner and Cancel instead.

import type { MemberHit } from "../sheet/layout";
import type { ReportRuntime, UiError } from "../state/runner";
import { formatElapsed, plural } from "../util";
import { ErrorNote, Spinner } from "./common";

export type Drill =
  /** The selected cell is not a drillable header of this report. */
  | { state: "none" }
  /** The cell is in the report, but there is no result in memory to map it with. */
  | { state: "stale" }
  | { state: "member"; hit: MemberHit };

interface Props {
  busy: string | null;
  onCancel: () => void;
  error: UiError | null;
  onDismissError: () => void;
  onRetry?: () => void;
  note: string | null;
  runtime: ReportRuntime | null;
  drill: Drill;
  onExpand: (hit: MemberHit) => void;
  onCollapse: (hit: MemberHit) => void;
  onShowQuery: () => void;
  onRefresh: () => void;
  onRefreshAll: () => void;
  canRefresh: boolean;
  canRefreshAll: boolean;
}

export function StatusBar(props: Props) {
  const { busy, runtime, drill } = props;
  const working = busy !== null;

  return (
    <footer className="status">
      {drill.state === "stale" && <p className="drill muted">Refresh this report to enable drill.</p>}
      {drill.state === "member" && (
        <div className="drill">
          <span className="drill-caption" title={drill.hit.member.caption}>
            {drill.hit.member.caption}
          </span>
          {drill.hit.member.hasChildren ? (
            drill.hit.member.expanded ? (
              <button
                type="button"
                className="button small"
                onClick={() => props.onCollapse(drill.hit)}
                disabled={working}
              >
                Collapse
              </button>
            ) : (
              <button
                type="button"
                className="button small"
                onClick={() => props.onExpand(drill.hit)}
                disabled={working}
              >
                Expand
              </button>
            )
          ) : (
            <span className="muted">Nothing below this item</span>
          )}
        </div>
      )}

      {props.error && (
        <ErrorNote
          error={props.error}
          onDismiss={props.onDismissError}
          action={props.onRetry ? { label: "Try again", run: props.onRetry } : undefined}
        />
      )}
      {props.note && <p className="note info">{props.note}</p>}

      {working ? (
        <div className="status-line">
          <span className="status-text" aria-live="polite">
            <Spinner label={busy} /> {busy}
          </span>
          <button type="button" className="button small" onClick={props.onCancel}>
            Cancel
          </button>
        </div>
      ) : runtime ? (
        <div className="status-line">
          <span className="status-text" aria-live="polite">
            {runtime.dataRows === 0 ? (
              "No data for this selection."
            ) : (
              <>
                {plural(runtime.stats.rows, "row", "rows")} × {plural(runtime.stats.columns, "column", "columns")}
                {" · "}
                {formatElapsed(runtime.stats.elapsedMs)}
              </>
            )}
            {runtime.stats.mdx !== "" && (
              <>
                {" · "}
                <button type="button" className="link" onClick={props.onShowQuery}>
                  Show query
                </button>
              </>
            )}
          </span>
        </div>
      ) : null}
      <div className="status-buttons">
        <button type="button" className="button" onClick={props.onRefreshAll} disabled={working || !props.canRefreshAll}>
          Refresh All
        </button>
        <button
          type="button"
          className="button primary"
          onClick={props.onRefresh}
          disabled={working || !props.canRefresh}
        >
          Refresh
        </button>
      </div>
    </footer>
  );
}
