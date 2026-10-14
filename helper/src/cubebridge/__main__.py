"""Command line: `python -m cubebridge serve | setup | probe`."""

from __future__ import annotations

import argparse
import logging
import logging.handlers
import signal
import sys
import threading
from pathlib import Path

from . import __version__, certs
from .api import App, serve
from .store import Store, default_directory

DEFAULT_PORT = 47613
RENEW_BELOW_DAYS = 30


def _setup(directory: Path, renew: bool) -> int:
    """Create the token and certificate. Prints CERT_CREATED or CERT_OK for install.sh."""
    Store(directory).token  # creates config.json on first run
    paths = certs.paths_in(directory / "certs")
    left = certs.days_left(paths.cert)
    if renew or left is None or left < RENEW_BELOW_DAYS or not paths.key.exists() or not paths.ca.exists():
        certs.generate(directory / "certs")
        print("CERT_CREATED")
    else:
        print("CERT_OK")
    return 0


def _serve(directory: Path, port: int) -> int:
    logs = directory / "logs"
    logs.mkdir(parents=True, exist_ok=True)
    handler = logging.handlers.RotatingFileHandler(logs / "helper.log", maxBytes=1_000_000, backupCount=3)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[handler, logging.StreamHandler(sys.stderr)],
    )
    log = logging.getLogger("cubebridge")
    # Our own step-by-step lines (names and byte counts, never content) are kept in
    # the log: they are what explains a connection that does not work.
    log.setLevel(logging.DEBUG)
    logging.getLogger("spnego").setLevel(logging.WARNING)  # it logs sign-in tokens at DEBUG

    paths = certs.paths_in(directory / "certs")
    if not paths.cert.exists() or not paths.key.exists():
        log.error("no certificate in %s; run install.sh", directory / "certs")
        return 2
    app = App(Store(directory), port=port, pane_dir=directory / "pane", cert_path=paths.cert)
    try:
        server = serve(app, certfile=paths.cert, keyfile=paths.key, port=port)
    except OSError as exc:
        log.error("cannot listen on 127.0.0.1:%d (%s)", port, exc)
        return 3

    def stop(*_args) -> None:
        threading.Thread(target=server.shutdown, daemon=True).start()

    signal.signal(signal.SIGTERM, stop)
    signal.signal(signal.SIGINT, stop)
    log.info("cubebridge %s listening on https://localhost:%d", __version__, port)
    try:
        server.serve_forever()
    finally:
        app.close()
        server.server_close()
        log.info("stopped")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="cubebridge", description="Helper for the Cube add-in in Excel for Mac.")
    parser.add_argument("--version", action="version", version=__version__)
    sub = parser.add_subparsers(dest="command", required=True)

    serve_parser = sub.add_parser("serve", help="run the local service (launchd does this)")
    serve_parser.add_argument("--dir", type=Path, default=default_directory())
    serve_parser.add_argument("--port", type=int, default=DEFAULT_PORT)

    setup_parser = sub.add_parser("setup", help="create settings and the localhost certificate")
    setup_parser.add_argument("--dir", type=Path, default=default_directory())
    setup_parser.add_argument("--renew-cert", action="store_true")

    probe_parser = sub.add_parser("probe", help="test the connection to your server")
    probe_parser.add_argument("--verbose", action="store_true", help="also print protocol step sizes")
    probe_parser.add_argument(
        "--sign-in",
        choices=["ntlm", "kerberos", "basic"],
        help="how to sign in (asked for when not given): ntlm, kerberos, or basic (https only)",
    )

    args = parser.parse_args(argv)
    if args.command == "serve":
        return _serve(args.dir, args.port)
    if args.command == "setup":
        return _setup(args.dir, args.renew_cert)
    from . import probe

    return probe.main(verbose=args.verbose, sign_in=args.sign_in)


if __name__ == "__main__":
    sys.exit(main())
