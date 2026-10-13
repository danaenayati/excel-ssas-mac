"""Saved connections and the install token.

`config.json` holds the token and each connection's server and username. Passwords
are never written there: they go to the macOS Keychain through `keyring`.
"""

from __future__ import annotations

import json
import os
import secrets as token_source
import threading
import uuid
from pathlib import Path
from typing import Protocol

from .errors import CubeError

APP_NAME = "CubeBridge"
DEMO_ID = "demo"
DEMO_CONNECTION = {"id": DEMO_ID, "name": "Sample cube", "kind": "demo", "server": "(built in)", "username": ""}


def default_directory() -> Path:
    return Path.home() / "Library" / "Application Support" / APP_NAME


class Secrets(Protocol):
    def get(self, connection_id: str) -> str | None: ...

    def set(self, connection_id: str, password: str) -> None: ...

    def delete(self, connection_id: str) -> None: ...


class KeychainSecrets:
    def get(self, connection_id: str) -> str | None:
        import keyring

        return keyring.get_password(APP_NAME, connection_id)

    def set(self, connection_id: str, password: str) -> None:
        import keyring

        keyring.set_password(APP_NAME, connection_id, password)

    def delete(self, connection_id: str) -> None:
        import keyring
        import keyring.errors

        try:
            keyring.delete_password(APP_NAME, connection_id)
        except keyring.errors.PasswordDeleteError:
            pass


class MemorySecrets:
    """For tests."""

    def __init__(self) -> None:
        self._items: dict[str, str] = {}

    def get(self, connection_id: str) -> str | None:
        return self._items.get(connection_id)

    def set(self, connection_id: str, password: str) -> None:
        self._items[connection_id] = password

    def delete(self, connection_id: str) -> None:
        self._items.pop(connection_id, None)


class Store:
    def __init__(self, directory: Path, secrets: Secrets | None = None) -> None:
        self.directory = Path(directory)
        self._secrets = secrets or KeychainSecrets()
        self._lock = threading.Lock()
        self._data: dict | None = None

    @property
    def _path(self) -> Path:
        return self.directory / "config.json"

    def _load(self) -> dict:
        if self._data is not None:
            return self._data
        if self._path.exists():
            try:
                data = json.loads(self._path.read_text(encoding="utf-8"))
                if not isinstance(data, dict) or not isinstance(data.get("connections", []), list):
                    raise ValueError("unexpected structure")
            except (ValueError, OSError) as exc:
                # Do not overwrite: the file may hold connections the user wants back.
                raise CubeError(
                    "internal",
                    "The helper's settings file is damaged. Run install.sh again to repair it.",
                    f"{self._path}: {exc}",
                ) from exc
            data.setdefault("connections", [])
            self._data = data
            if not data.get("token"):
                data["token"] = token_source.token_urlsafe(32)
                self._save()
        else:
            self._data = {"token": token_source.token_urlsafe(32), "connections": []}
            self._save()
        return self._data

    def _save(self) -> None:
        self.directory.mkdir(parents=True, exist_ok=True)
        os.chmod(self.directory, 0o700)
        temporary = self._path.with_suffix(".tmp")
        descriptor = os.open(temporary, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
            json.dump(self._data, handle, ensure_ascii=False, indent=2)
        os.chmod(temporary, 0o600)
        os.replace(temporary, self._path)

    @property
    def token(self) -> str:
        with self._lock:
            return self._load()["token"]

    def connections(self) -> list[dict]:
        with self._lock:
            return [dict(DEMO_CONNECTION)] + [dict(c) for c in self._load()["connections"]]

    def get(self, connection_id: str) -> dict:
        for connection in self.connections():
            if connection["id"] == connection_id:
                return connection
        raise CubeError("not-found", "That connection no longer exists on this Mac.")

    def add(self, *, name: str, server: str, username: str, password: str) -> dict:
        connection = {
            "id": uuid.uuid4().hex[:12],
            "name": name.strip() or server,
            "kind": "ssas",
            "server": server,
            "username": username,
        }
        with self._lock:
            self._secrets.set(connection["id"], password)
            self._load()["connections"].append(connection)
            self._save()
        return dict(connection)

    def delete(self, connection_id: str) -> None:
        if connection_id == DEMO_ID:
            raise CubeError("bad-request", "The sample cube is built in and cannot be removed.")
        with self._lock:
            data = self._load()
            remaining = [c for c in data["connections"] if c["id"] != connection_id]
            if len(remaining) == len(data["connections"]):
                raise CubeError("not-found", "That connection no longer exists on this Mac.")
            data["connections"] = remaining
            self._save()
            self._secrets.delete(connection_id)

    def password(self, connection_id: str) -> str:
        self.get(connection_id)
        password = self._secrets.get(connection_id)
        if password is None:
            raise CubeError(
                "auth",
                "The saved password for this connection is missing from the Keychain. "
                "Remove the connection and add it again.",
            )
        return password
