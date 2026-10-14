// The preview's charts: one small box per chart on the sheet in view, saying what
// Excel would draw (type, title, series, categories) and where. Not a picture: the
// preview is a stand-in for trying the pane in a browser. "Delete in Excel" does what
// selecting the chart in Excel and pressing Delete would.

import { chartKindLabel } from "../sheet/charts";
import { cellA1 } from "../sheet/geometry";
import type { PreviewChart } from "../sheet/preview";

/** Names listed before "…". */
const SHOWN = 4;

function names(list: string[]): string {
  if (list.length === 0) return "none";
  const head = list.slice(0, SHOWN).join(", ");
  return list.length > SHOWN ? `${head}, …` : head;
}

const strip = {
  display: "flex",
  flex: "none",
  gap: 8,
  padding: "6px 10px",
  overflowX: "auto",
  borderTop: "1px solid #d4d4d4",
  background: "#f8f8f8",
  fontSize: 12,
} as const;

const box = {
  flex: "none",
  width: 260,
  margin: 0,
  padding: "6px 8px",
  border: "1px solid #c6c6c6",
  borderRadius: 4,
  background: "#ffffff",
  display: "grid",
  gap: 2,
} as const;

const quiet = { color: "#6a6a6f" } as const;

export function PreviewCharts({
  charts,
  onDelete,
}: {
  charts: readonly PreviewChart[];
  onDelete: (chartId: string) => void;
}) {
  if (charts.length === 0) return null;
  return (
    <div style={strip} role="list" aria-label="Charts on this sheet">
      {charts.map((chart) => {
        const { series, categories } = chart.table;
        return (
          <figure key={chart.id} style={box} role="listitem" data-testid="pv-chart">
            <figcaption>
              <strong>{chart.title}</strong>
              <span style={quiet}>
                {" "}
                · {chartKindLabel(chart.type)} chart at {cellA1(chart.at.row, chart.at.col)}
              </span>
            </figcaption>
            <span title={series.join("\n")}>
              <span style={quiet}>{series.length === 1 ? "1 series" : `${series.length} series`}:</span> {names(series)}
              {chart.type === "Pie" && series.length > 1 && <span style={quiet}> (a pie shows the first)</span>}
            </span>
            <span title={categories.join("\n")}>
              <span style={quiet}>
                {categories.length === 1 ? "1 category" : `${categories.length} categories`}:
              </span>{" "}
              {names(categories)}
            </span>
            <span>
              <button
                type="button"
                className="button small"
                title="As if the chart were deleted in Excel: the report lets go of it at its next update"
                onClick={() => onDelete(chart.id)}
              >
                Delete in Excel
              </button>
            </span>
          </figure>
        );
      })}
    </div>
  );
}
