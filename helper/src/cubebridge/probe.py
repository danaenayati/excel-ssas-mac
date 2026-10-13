"""Connection test for a real server: `python -m cubebridge probe`.

Run by the user in Terminal, because only they can type the password. It walks the
same path the add-in uses (name lookup, port, sign-in, metadata, one tiny query) and
says in plain words where it stops. The password is read with echo off and is never
printed or logged.
"""

from __future__ import annotations

import getpass
import logging
import socket
import sys
from collections.abc import Callable

from .backends.ssas_backend import SsasBackend
from .errors import CubeError
from .pivot.grid import build_grid
from .pivot.model import Layout
from .ssas import resolve
from .ssas.session import Session

Printer = Callable[[str], None]


def _step(out: Printer, label: str, result: str) -> None:
    out(f"  {label:.<26} {result}")


def _port_open(host: str, port: int, timeout: float = 3.0) -> bool:
    try:
        socket.create_connection((host, port), timeout=timeout).close()
        return True
    except OSError:
        return False


def run(server: str, username: str, password: str, out: Printer = print) -> bool:
    """Return True when every step passed."""
    try:
        target = resolve.parse_server(server)
    except CubeError as exc:
        _step(out, "Server name", f"NOT VALID: {exc.message}")
        return False

    try:
        address = socket.gethostbyname(target.host)
        _step(out, "Name lookup", f"OK ({address})")
    except OSError:
        _step(out, "Name lookup", "FAILED")
        out(f"\n  '{target.host}' could not be found. Connect the company VPN and try again,")
        out("  or check the spelling against what you type in Windows Excel.")
        return False

    try:
        host, port = resolve.endpoint(target)
    except CubeError as exc:
        _step(out, "Instance port", "FAILED")
        out(f"\n  {exc.message}")
        return False
    if target.instance:
        _step(out, "Instance port", f"OK ({port})")

    if not _port_open(host, port):
        _step(out, f"Port {port}", "CLOSED or blocked")
        other = [str(p) for p in (2382, 443, 80) if _port_open(host, p, 2.0)]
        out(f"\n  The server answers by name but not on port {port}.")
        if other:
            out(f"  Other open ports on it: {', '.join(other)}.")
        out("  If the cube runs as a named instance, enter the server as  name\\INSTANCE  or  name:port.")
        return False
    _step(out, f"Port {port}", "OK")

    backend = SsasBackend(server, username, password)
    try:
        try:
            catalogs = backend.catalogs()
        except CubeError as exc:
            _step(out, "Sign-in", "FAILED")
            out(f"\n  {exc.message}")
            if exc.detail:
                out(f"  Detail: {exc.detail}")
            return False
        _step(out, "Sign-in", "OK")

        try:
            session = Session(host, port, username, password)
            session.open()
            try:
                rows = session.discover("DISCOVER_PROPERTIES", {"PropertyName": "DBMSVersion"})
            finally:
                session.close()
            version = next((r.get("Value", "") for r in rows if r.get("PropertyName") == "DBMSVersion"), "")
            _step(out, "Server version", version or "(not reported)")
        except CubeError:
            _step(out, "Server version", "(not reported)")

        names = [c["name"] for c in catalogs]
        _step(out, "Databases", ", ".join(names) if names else "none visible to this account")
        if not names:
            out("\n  Sign-in works, but this account cannot see any database. Ask for read access.")
            return False

        for catalog in names:
            try:
                cubes = backend.cubes(catalog)
            except CubeError as exc:
                _step(out, f"Cubes in {catalog}"[:26], f"FAILED: {exc.message}")
                continue
            _step(out, f"Cubes in {catalog}"[:26], ", ".join(c["caption"] for c in cubes) or "none")
            if not cubes:
                continue
            cube = cubes[0]["name"]
            try:
                layout = Layout.from_json({})
                cellset, text = backend.query(catalog, cube, layout)
                grid = build_grid(cellset, layout, mdx=text)
                fields = backend.fields(catalog, cube)
                _step(
                    out,
                    "Test query",
                    f"OK ({grid['stats']['columns']} value from '{cube}', "
                    f"{len(fields['measures'])} measures, "
                    f"{sum(len(d['hierarchies']) for d in fields['dimensions'])} fields)",
                )
                return True
            except CubeError as exc:
                _step(out, "Test query", "FAILED")
                out(f"\n  {exc.message}")
                if exc.detail:
                    out(f"  Detail: {exc.detail}")
                return False
        out("\n  Sign-in works, but no cube is visible in any database.")
        return False
    finally:
        backend.close()


def main(verbose: bool = False) -> int:
    if verbose:
        # Sizes and step names only: no passwords, no sign-in tokens, no cube data.
        logging.basicConfig(level=logging.DEBUG, format="    [%(name)s] %(message)s", stream=sys.stdout)
    print("Cube connection test")
    print("Nothing here is sent anywhere except to your own server.\n")
    try:
        server = input("Server (as typed in Windows Excel): ").strip()
        username = input("Username (DOMAIN\\user): ").strip()
        password = getpass.getpass("Password (typing is hidden): ")
    except (EOFError, KeyboardInterrupt):
        print("\nCancelled.")
        return 130
    print()
    ok = run(server, username, password)
    print()
    if not ok:
        print("Not working yet. Copy the lines above (they contain no password) when asking for help.")
        return 1
    print("Everything works.")
    try:
        answer = input("Save this connection so it appears in Excel? [Y/n] ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        answer = "n"
    if answer in ("", "y", "yes"):
        from .store import Store, default_directory

        Store(default_directory()).add(name="", server=server, username=username, password=password)
        print("Saved. Open the Cube panel in Excel and choose it from the list.")
    return 0


if __name__ == "__main__":
    sys.exit(main("--verbose" in sys.argv[1:]))
