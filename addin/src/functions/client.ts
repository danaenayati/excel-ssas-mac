// The cube functions' way to the helper. They run in Excel's JavaScript-only runtime,
// not in the pane: there is no page, so URLs are absolute, and Microsoft documents that
// runtime as allowing only "simple" CORS requests. So every request is a POST of
// text/plain with no custom header, which needs no preflight, and the install token
// travels in the JSON body instead of the X-Cube-Token header. The helper accepts that
// on the formula routes only (helper api.py).

import { ApiFailure, HelperDown } from "../apiErrors";
import type {
  ApiError,
  FormulaCellsRequest,
  FormulaConnectionsResponse,
  FormulaLookupRequest,
  FormulaResponse,
} from "../types";
import type { FormulaTransport } from "./batch";

/** Where install.sh runs the helper; the same address as in the manifest. */
export const HELPER = "https://localhost:47613";

type Fetch = (url: string, init: { method: string; headers: Record<string, string>; body: string }) => Promise<{
  ok: boolean;
  status: number;
  text(): Promise<string>;
}>;

function isApiError(value: unknown): value is ApiError {
  if (typeof value !== "object" || value === null) return false;
  const v = value as Record<string, unknown>;
  return typeof v.category === "string" && typeof v.message === "string";
}

export function createFormulaTransport(token: string, fetchImpl?: Fetch, base: string = HELPER): FormulaTransport {
  async function post<T>(path: string, body: object): Promise<T> {
    const send: Fetch = fetchImpl ?? ((url, init) => fetch(url, init));
    let status: number;
    let ok: boolean;
    let text: string;
    try {
      const response = await send(base + "/api" + path, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=UTF-8" },
        body: JSON.stringify({ ...body, token }),
      });
      status = response.status;
      ok = response.ok;
      text = await response.text();
    } catch (error) {
      throw new HelperDown(error instanceof Error ? error.message : String(error));
    }
    let parsed: unknown = undefined;
    try {
      parsed = text === "" ? null : JSON.parse(text);
    } catch {
      // handled below
    }
    if (ok && parsed !== undefined && parsed !== null) return parsed as T;
    const error = typeof parsed === "object" && parsed !== null ? (parsed as { error?: unknown }).error : undefined;
    if (isApiError(error)) throw new ApiFailure(error, status);
    throw new ApiFailure(
      {
        category: status === 403 ? "forbidden" : "internal",
        message: `The helper answered with an error (HTTP ${status}).`,
        detail: text.slice(0, 300) || undefined,
      },
      status,
    );
  }

  const id = encodeURIComponent;
  return {
    connections: async () =>
      (await post<FormulaConnectionsResponse>("/formulas/connections", {})).connections,
    cells: async (connectionId, request: FormulaCellsRequest) =>
      (await post<FormulaResponse>(`/connections/${id(connectionId)}/cells`, request)).results,
    lookup: async (connectionId, request: FormulaLookupRequest) =>
      (await post<FormulaResponse>(`/connections/${id(connectionId)}/lookup`, request)).results,
  };
}
