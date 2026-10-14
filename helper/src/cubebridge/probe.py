"""Connection test for a real server: `python -m cubebridge probe`.

Run by the user in Terminal, because only they can type the password. It walks the
same path the add-in uses (name lookup, port, sign-in, metadata, one tiny query) and
says in plain words where it stops. The password is read with echo off and is never
printed or logged.

The server may be a server name (TCP) or the https:// address of msmdpump.dll, and
the sign-in kind is asked for as the pane asks for it (or given with --sign-in).
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
from .ssas import auth, resolve

Printer = Callable[[str], None]


def _step(out: Printer, label: str, result: str) -> None:
    out(f"  {label:.<26} {result}")


def _port_open(host: str, port: int, timeout: float = 3.0) -> bool:
    try:
        socket.create_connection((host, port), timeout=timeout).close()
        return True
    except OSError:
        return False


def _kerberos_ready(out: Printer, username: str) -> bool:
    from .ssas import kerberos

    try:
        kerberos.check_installed()
    except CubeError as exc:
        _step(out, "Kerberos on this Mac", "NOT INSTALLED")
        out(f"\n  {exc.message}")
        return False
    _step(out, "Kerberos on this Mac", "OK")
    ticket = kerberos.existing_ticket()
    if ticket is not None and kerberos.ticket_matches(ticket[1], username):
        _step(out, "Kerberos ticket", f"already on this Mac ({ticket[1]})")
    elif ticket is not None:
        _step(out, "Kerberos ticket", f"one for {ticket[1]}; not used, the password gets one")
    else:
        _step(out, "Kerberos ticket", "none yet; the password gets one")
    return True


def run(server: str, username: str, password: str, out: Printer = print, sign_in: str = auth.DEFAULT_SIGN_IN) -> bool:
    """Return True when every step passed."""
    try:
        target = resolve.parse_target(server)
        sign_in = auth.read_sign_in(sign_in)
        auth.check_allowed(sign_in, target)
    except CubeError as exc:
        _step(out, "Server name", f"NOT VALID: {exc.message}")
        return False
    web = isinstance(target, resolve.HttpTarget)
    if web:
        _step(out, "Web address", f"OK ({target.scheme}, port {target.port})")
    _step(out, "Sign in with", auth.LABELS[sign_in])
    if sign_in == "kerberos" and not _kerberos_ready(out, username):
        return False

    try:
        address = socket.gethostbyname(target.host)
        _step(out, "Name lookup", f"OK ({address})")
    except OSError:
        _step(out, "Name lookup", "FAILED")
        out(f"\n  '{target.host}' could not be found. Connect the company VPN and try again,")
        out("  or check the spelling against what you type in Windows Excel.")
        return False

    if web:
        host, port = target.host, target.port
    else:
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
        out(f"\n  The server answers by name but not on port {port}.")
        if not web:
            other = [str(p) for p in (2382, 443, 80) if _port_open(host, p, 2.0)]
            if other:
                out(f"  Other open ports on it: {', '.join(other)}.")
            out("  If the cube runs as a named instance, enter the server as  name\\INSTANCE  or  name:port.")
        return False
    _step(out, f"Port {port}", "OK")

    backend = SsasBackend(server, username, password, sign_in=sign_in)
    try:
        try:
            catalogs = backend.catalogs()
        except CubeError as exc:
            _step(out, "Sign-in", "FAILED")
            out(f"\n  {exc.message}")
            if exc.detail:
                out(f"  Detail: {exc.detail}")
            return False
        _step(out, "Sign-in", f"OK ({backend.sign_in_note})" if backend.sign_in_note else "OK")

        try:
            _step(out, "Server version", backend.server_version() or "(not reported)")
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


def enable_verbose(stream=None) -> Callable[[], None]:
    """Print this package's own step logging. Returns a function that undoes it.

    Only the "cubebridge" logger is switched on, never the root logger: the sign-in
    library logs every NTLM message at DEBUG, and those are enough to attack the
    password offline. Our own lines carry step names and byte counts only.
    """
    logger = logging.getLogger("cubebridge")
    handler = logging.StreamHandler(stream or sys.stdout)
    handler.setFormatter(logging.Formatter("    [%(name)s] %(message)s"))
    previous = (logger.level, logger.propagate)
    logger.addHandler(handler)
    logger.setLevel(logging.DEBUG)
    logger.propagate = False
    logging.getLogger("spnego").setLevel(logging.WARNING)

    def undo() -> None:
        logger.removeHandler(handler)
        logger.setLevel(previous[0])
        logger.propagate = previous[1]

    return undo


def sign_in_choices(server: str) -> list[str]:
    """The kinds offered for this server, the default first (as in the pane)."""
    try:
        target = resolve.parse_target(server)
    except CubeError:
        target = None
    secure_web = isinstance(target, resolve.HttpTarget) and target.secure
    return ["ntlm", "kerberos"] + (["basic"] if secure_web else [])


def ask_sign_in(server: str, ask: Callable[[str], str] = input, show: Printer = print) -> str:
    choices = sign_in_choices(server)
    show("Sign in with:")
    for number, kind in enumerate(choices, 1):
        show(f"  {number}  {auth.LABELS[kind]}")
    while True:
        answer = ask(f"Choose 1-{len(choices)} [1]: ").strip()
        if not answer:
            return choices[0]
        if answer.isdigit() and 1 <= int(answer) <= len(choices):
            return choices[int(answer) - 1]
        show("  Type the number of a choice, or press Return for 1.")


def main(verbose: bool = False, sign_in: str | None = None) -> int:
    if verbose:
        enable_verbose()
    print("Cube connection test")
    print("Nothing here is sent anywhere except to your own server.\n")
    try:
        server = input("Server (as typed in Windows Excel, or https://…/msmdpump.dll): ").strip()
        if sign_in is None:
            sign_in = ask_sign_in(server)
        username = input("Username (DOMAIN\\user): ").strip()
        hint = "used only if this Mac has no Kerberos ticket; " if sign_in == "kerberos" else ""
        password = getpass.getpass(f"Password ({hint}typing is hidden): ")
    except (EOFError, KeyboardInterrupt):
        print("\nCancelled.")
        return 130
    print()
    ok = run(server, username, password, sign_in=sign_in)
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

        Store(default_directory()).add(name="", server=server, username=username, password=password, sign_in=sign_in)
        print("Saved. Open the Cube panel in Excel and choose it from the list.")
    return 0


if __name__ == "__main__":
    sys.exit(main("--verbose" in sys.argv[1:]))
