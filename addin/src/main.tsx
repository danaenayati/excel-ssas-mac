// Start-up: wait for Office.js, choose the sheet host once, choose the API client,
// render. Inside Excel the host is ExcelHost; in an ordinary browser Office.onReady
// resolves with host null and the pane runs against the in-page PreviewHost instead.

import { createRoot } from "react-dom/client";
import { App } from "./App";
import { createApi, isMockRequested } from "./api";
import type { SheetHost } from "./sheet/host";
import { PreviewHost } from "./sheet/preview";
import { ExcelHost } from "./sheet/reports";
import { supportsExcelApi } from "./sheet/writer";
import "./styles.css";

/** True when the page is running inside Excel. Never rejects. */
async function insideExcel(): Promise<boolean> {
  try {
    // office.js is loaded from Microsoft's CDN by a plain script tag; offline, or in
    // a browser that blocks it, the global is simply missing.
    if (typeof Office === "undefined" || typeof Office.onReady !== "function") return false;
    const info = await Office.onReady();
    return info !== null && info !== undefined && info.host === Office.HostType.Excel;
  } catch {
    return false;
  }
}

/** True when this page is inside another page's frame. */
function isFramed(): boolean {
  try {
    return window.top !== window.self;
  } catch {
    return true;
  }
}

async function start(): Promise<void> {
  const container = document.getElementById("root");
  if (!container) return;
  const root = createRoot(container);

  const excel = await insideExcel();

  // Outside Excel the pane still works (as a browser preview), and it carries the
  // helper's token. Framed by some other page, that would let the page overlay it and
  // steer the user's clicks, so a framed copy that is not in Excel does nothing.
  if (!excel && isFramed()) {
    root.render(
      <div className="app">
        <div className="pane">
          <div className="center">Open Cube Explorer from the Home tab in Excel.</div>
        </div>
      </div>,
    );
    return;
  }
  document.documentElement.setAttribute("data-host", excel ? "excel" : "preview");

  // Workbook names and settings need ExcelApi 1.4. Every Excel for Mac that can load
  // this add-in has it, but say so plainly rather than fail later.
  if (excel && !supportsExcelApi("1.4")) {
    root.render(
      <div className="app">
        <div className="pane">
          <div className="center">This version of Excel is too old for Cube Explorer. Update Excel and open the pane again.</div>
        </div>
      </div>,
    );
    return;
  }

  const host: SheetHost = excel ? new ExcelHost() : new PreviewHost();
  root.render(<App api={createApi()} host={host} mock={isMockRequested()} />);
}

void start();
