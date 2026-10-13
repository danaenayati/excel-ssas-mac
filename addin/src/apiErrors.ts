// Errors thrown by the API client. Kept apart from api.ts so the mock client can
// throw the same classes without a circular import.

import type { ApiError } from "./types";

/** The helper answered with an error (`{ "error": ApiError }`, non-2xx). */
export class ApiFailure extends Error {
  readonly error: ApiError;
  readonly status: number;
  constructor(error: ApiError, status: number) {
    super(error.message);
    this.name = "ApiFailure";
    this.error = error;
    this.status = status;
  }
}

/** The request never reached the helper: it is not running, or its certificate is not trusted. */
export class HelperDown extends Error {
  constructor(detail?: string) {
    super(detail || "The helper is not running.");
    this.name = "HelperDown";
  }
}

export function isCancelled(error: unknown): boolean {
  return error instanceof ApiFailure && error.error.category === "cancelled";
}
