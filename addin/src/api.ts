// Typed client for the helper's JSON API. In production the helper serves this page
// itself, so every URL is relative and same-origin. Every request carries the install
// token in `X-Cube-Token`.

import { ApiFailure, HelperDown, isCancelled } from "./apiErrors";
import { createMockApi } from "./dev/mockApi";
import type {
  ApiError,
  Catalog,
  Connection,
  Cube,
  Fields,
  Grid,
  MembersRequest,
  MembersResponse,
  NewConnection,
  QueryRequest,
  Status,
} from "./types";

export interface Api {
  status(): Promise<Status>;
  connections(): Promise<Connection[]>;
  /** The helper tests the connection before saving it, so this can take a while. */
  addConnection(connection: NewConnection): Promise<Connection>;
  deleteConnection(id: string): Promise<void>;
  catalogs(connectionId: string): Promise<Catalog[]>;
  cubes(connectionId: string, catalog: string): Promise<Cube[]>;
  fields(connectionId: string, catalog: string, cube: string): Promise<Fields>;
  members(connectionId: string, request: MembersRequest): Promise<MembersResponse>;
  /** Aborting `signal` discards the response; it does not stop the query on the server. */
  query(connectionId: string, request: QueryRequest, signal?: AbortSignal): Promise<Grid>;
  /** Stops the query running on this connection, if any. */
  cancel(connectionId: string): Promise<void>;
}

export { ApiFailure, HelperDown, isCancelled };

const TOKEN_HEADER = "X-Cube-Token";

/**
 * The install token. The helper writes it into the page's `cube-token` meta tag when it
 * serves index.html. Under the Vite dev server the tag still holds the placeholder, and
 * the token comes from the `VITE_CUBE_TOKEN` environment variable instead.
 */
export function readToken(): string {
  const meta = document.querySelector('meta[name="cube-token"]');
  const value = (meta && meta.getAttribute("content")) || "";
  if (value !== "" && !value.startsWith("__CUBE_")) return value;
  return import.meta.env.VITE_CUBE_TOKEN || "";
}

function isAbort(error: unknown): boolean {
  return error instanceof Error && error.name === "AbortError";
}

function isApiError(value: unknown): value is ApiError {
  if (typeof value !== "object" || value === null) return false;
  const v = value as Record<string, unknown>;
  return typeof v.category === "string" && typeof v.message === "string";
}

export function createHttpApi(token: string, base = ""): Api {
  async function call<T>(
    method: "GET" | "POST" | "DELETE",
    path: string,
    body?: unknown,
    signal?: AbortSignal,
  ): Promise<T> {
    const headers: Record<string, string> = { [TOKEN_HEADER]: token, Accept: "application/json" };
    if (body !== undefined) headers["Content-Type"] = "application/json";

    let response: Response;
    try {
      response = await fetch(base + "/api" + path, {
        method,
        headers,
        body: body === undefined ? undefined : JSON.stringify(body),
        signal,
        cache: "no-store",
        credentials: "same-origin",
      });
    } catch (error) {
      if (isAbort(error)) {
        throw new ApiFailure({ category: "cancelled", message: "Cancelled." }, 499);
      }
      throw new HelperDown(error instanceof Error ? error.message : String(error));
    }

    let text: string;
    try {
      text = await response.text();
    } catch (error) {
      if (isAbort(error)) {
        throw new ApiFailure({ category: "cancelled", message: "Cancelled." }, 499);
      }
      throw new HelperDown(error instanceof Error ? error.message : String(error));
    }

    let parsed: unknown;
    let readable = true;
    try {
      parsed = text === "" ? null : JSON.parse(text);
    } catch {
      readable = false;
    }

    if (response.ok) {
      if (!readable) {
        throw new ApiFailure(
          {
            category: "internal",
            message: "The helper sent an answer the pane could not read.",
            detail: text.slice(0, 500),
          },
          response.status,
        );
      }
      return parsed as T;
    }

    const error =
      readable && typeof parsed === "object" && parsed !== null
        ? (parsed as { error?: unknown }).error
        : undefined;
    if (isApiError(error)) throw new ApiFailure(error, response.status);

    // Under `npm run dev` the Vite proxy answers 5xx itself when the helper is down.
    if (import.meta.env.DEV && response.status >= 500) {
      throw new HelperDown(`Proxy answered HTTP ${response.status}.`);
    }
    throw new ApiFailure(
      {
        category: response.status === 403 ? "forbidden" : "internal",
        message:
          response.status === 403
            ? "The helper refused the request. Close this pane and open it again."
            : `The helper answered with an error (HTTP ${response.status}).`,
        detail: text.slice(0, 500) || undefined,
      },
      response.status,
    );
  }

  const id = encodeURIComponent;
  const query = (params: Record<string, string>) =>
    "?" +
    Object.keys(params)
      .map((key) => `${key}=${encodeURIComponent(params[key])}`)
      .join("&");

  return {
    status: () => call<Status>("GET", "/status"),
    connections: () => call<Connection[]>("GET", "/connections"),
    addConnection: (connection) => call<Connection>("POST", "/connections", connection),
    deleteConnection: async (connectionId) => {
      await call<{ ok: true }>("DELETE", `/connections/${id(connectionId)}`);
    },
    catalogs: (connectionId) => call<Catalog[]>("GET", `/connections/${id(connectionId)}/catalogs`),
    cubes: (connectionId, catalog) =>
      call<Cube[]>("GET", `/connections/${id(connectionId)}/cubes${query({ catalog })}`),
    fields: (connectionId, catalog, cube) =>
      call<Fields>("GET", `/connections/${id(connectionId)}/fields${query({ catalog, cube })}`),
    members: (connectionId, request) =>
      call<MembersResponse>("POST", `/connections/${id(connectionId)}/members`, request),
    query: (connectionId, request, signal) =>
      call<Grid>("POST", `/connections/${id(connectionId)}/query`, request, signal),
    cancel: async (connectionId) => {
      await call<{ ok: true }>("POST", `/connections/${id(connectionId)}/cancel`);
    },
  };
}

/** True when the page was opened with `?mock=1`: no helper, canned data. */
export function isMockRequested(search: string = window.location.search): boolean {
  return new URLSearchParams(search).get("mock") === "1";
}

export function createApi(): Api {
  if (isMockRequested()) {
    const params = new URLSearchParams(window.location.search);
    return createMockApi({ slow: params.get("slow") === "1" });
  }
  return createHttpApi(readToken());
}
