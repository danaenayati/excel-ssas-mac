"""One error type for everything the pane has to explain to the user.

`category` is part of the API contract (addin/src/types.ts, ErrorCategory): the pane
acts on it without parsing text. `message` is a plain sentence that can be shown as
is; `detail` carries technical text (usually the server's own words).
"""

from __future__ import annotations

CATEGORIES = frozenset(
    {
        "unreachable",
        "auth",
        "auth-mechanism",
        "instance",
        "too-large",
        "server",
        "cancelled",
        "bad-request",
        "not-found",
        "forbidden",
        "internal",
    }
)

HTTP_STATUS = {
    "bad-request": 400,
    "auth": 401,
    "auth-mechanism": 401,
    "forbidden": 403,
    "not-found": 404,
    "too-large": 413,
    "cancelled": 499,
    "internal": 500,
    "unreachable": 502,
    "instance": 502,
    "server": 502,
}


class CubeError(Exception):
    def __init__(self, category: str, message: str, detail: str | None = None) -> None:
        if category not in CATEGORIES:
            raise ValueError(f"unknown error category {category!r}")
        super().__init__(message)
        self.category = category
        self.message = message
        self.detail = detail

    @property
    def http_status(self) -> int:
        return HTTP_STATUS[self.category]

    def to_json(self) -> dict:
        body = {"category": self.category, "message": self.message}
        if self.detail:
            body["detail"] = self.detail
        return body


class ProtocolError(CubeError):
    """The server sent bytes this client cannot interpret."""

    def __init__(self, detail: str) -> None:
        super().__init__("server", "The server sent a response this add-in could not read.", detail)


class IncompleteMessage(Exception):
    """More bytes are needed. Internal to the framing code; never reaches the API."""
