"""Kerberos for real: a throwaway MIT KDC on 127.0.0.1, real tickets, real GSS on both
sides, and nothing in the helper patched.

The stand-in cube servers (TCP and HTTP) sign in with a GSS acceptor that reads the
service keytab; the client is SsasBackend with sign_in="kerberos", as a saved connection
makes it. Each check runs in a process of its own, so that the GSS library reads this
realm's configuration and nothing else. Skipped when MIT Kerberos (krb5kdc) is not
installed: `brew install krb5` to run it.
"""

from __future__ import annotations

import json
import os
import shutil
import socket
import subprocess
import sys
import textwrap
import time
from pathlib import Path

import pytest

pytest.importorskip("gssapi")

REALM = "CUBE.TEST"
USER = f"dana@{REALM}"
PASSWORD = "Test-Pass-123"
HERE = Path(__file__).resolve().parent


def _krb5_sbin() -> Path | None:
    found = shutil.which("krb5kdc")
    if found:
        return Path(found).parent
    for prefix in ("/opt/homebrew/opt/krb5/sbin", "/usr/local/opt/krb5/sbin"):
        if (Path(prefix) / "krb5kdc").exists():
            return Path(prefix)
    return None


SBIN = _krb5_sbin()
pytestmark = pytest.mark.skipif(SBIN is None, reason="MIT Kerberos (krb5kdc) is not installed")


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


@pytest.fixture(scope="module")
def realm(tmp_path_factory):
    home = tmp_path_factory.mktemp("krb")
    port = _free_port()
    (home / "krb5.conf").write_text(textwrap.dedent(f"""\
        [libdefaults]
          default_realm = {REALM}
          dns_lookup_kdc = false
          dns_lookup_realm = false
          rdns = false
        [realms]
          {REALM} = {{
            kdc = 127.0.0.1:{port}
          }}
        [domain_realm]
          localhost = {REALM}
        """))
    (home / "kdc.conf").write_text(textwrap.dedent(f"""\
        [kdcdefaults]
          kdc_listen = 127.0.0.1:{port}
          kdc_tcp_listen = 127.0.0.1:{port}
        [realms]
          {REALM} = {{
            database_name = {home}/principal
            key_stash_file = {home}/stash
            supported_enctypes = aes256-cts-hmac-sha1-96:normal
          }}
        [logging]
          kdc = FILE:{home}/kdc.log
        """))
    env = dict(os.environ, KRB5_CONFIG=str(home / "krb5.conf"), KRB5_KDC_PROFILE=str(home / "kdc.conf"))

    def admin(query: str) -> None:
        subprocess.run([SBIN / "kadmin.local", "-q", query], env=env, check=True, capture_output=True)

    subprocess.run([SBIN / "kdb5_util", "create", "-r", REALM, "-s", "-P", "master-" + str(port)],
                   env=env, check=True, capture_output=True)
    admin(f"addprinc -pw {PASSWORD} {USER}")
    for service in ("MSOLAPSvc.3/localhost", "HTTP/localhost"):
        admin(f"addprinc -randkey {service}@{REALM}")
        admin(f"ktadd -k {home}/service.keytab {service}@{REALM}")
    kdc = subprocess.Popen([SBIN / "krb5kdc", "-n"], env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        deadline = time.monotonic() + 10
        while True:
            try:
                socket.create_connection(("127.0.0.1", port), timeout=0.5).close()
                break
            except OSError:
                if time.monotonic() > deadline:
                    raise RuntimeError("the test KDC did not start")
                time.sleep(0.1)
        yield {
            "KRB5_CONFIG": str(home / "krb5.conf"),
            "KRB5_KTNAME": f"FILE:{home}/service.keytab",
            "KRB5CCNAME": f"FILE:{home}/client.cc",
            "log": home / "kdc.log",
        }
    finally:
        kdc.terminate()
        kdc.wait(timeout=10)


CHECK = textwrap.dedent("""\
    import json, sys
    sys.path[:0] = [sys.argv[1], sys.argv[2]]
    import spnego
    import xml_samples as x
    from fake_http_xmla import FakeXmlaHttp
    from fake_ssas import FakeSsas
    from cubebridge.backends.ssas_backend import SsasBackend
    from cubebridge.errors import CubeError

    kind, user, password = sys.argv[3], sys.argv[4], sys.argv[5]
    data = {"DBSCHEMA_CATALOGS": x.rowset([{"CATALOG_NAME": "Sales", "DESCRIPTION": ""}])}
    accepted = []

    def acceptor(service):
        context = spnego.server(hostname="localhost", service=service, protocol="kerberos")
        accepted.append(context)
        return context

    if kind == "tcp":
        server = FakeSsas(data, acceptor=lambda: acceptor("MSOLAPSvc.3"))
        address = lambda s: f"localhost:{s.port}"
    else:
        server = FakeXmlaHttp(data, offers=("Negotiate",), acceptor=lambda scheme: acceptor("HTTP"))
        address = lambda s: f"http://localhost:{s.port}/olap/msmdpump.dll"
    out = {}
    with server:
        backend = SsasBackend(address(server), user, password, sign_in="kerberos")
        try:
            out["catalogs"] = [c["name"] for c in backend.catalogs()]
            out["again"] = [c["name"] for c in backend.catalogs()]
        except CubeError as exc:
            out["error"] = [exc.category, exc.message]
        finally:
            backend.close()
    done = [c for c in accepted if c.complete]
    out["client"] = done[0].client_principal if done else None
    print(json.dumps(out))
    """)


def _run(realm, kind: str, password: str = PASSWORD) -> dict:
    env = dict(os.environ, **{k: v for k, v in realm.items() if k != "log"})
    src = str(HERE.parent / "src")
    result = subprocess.run([sys.executable, "-I", "-c", CHECK, src, str(HERE), kind, USER, password],
                            env=env, capture_output=True, text=True, timeout=120)
    assert result.returncode == 0, result.stderr[-2000:]
    return json.loads(result.stdout.strip().splitlines()[-1])


@pytest.mark.parametrize("kind,service", [("tcp", "MSOLAPSvc.3"), ("http", "HTTP")])
def test_signs_in_with_a_real_ticket_and_keeps_the_sign_in(realm, kind, service):
    out = _run(realm, kind)
    assert out.get("error") is None
    assert out["catalogs"] == out["again"] == ["Sales"]
    assert out["client"] == USER
    # The ticket asked for is the service name Windows asks for.
    assert f"for {service}/localhost@{REALM}" in realm["log"].read_text()


def test_a_wrong_password_is_refused_in_plain_words(realm):
    out = _run(realm, "tcp", password="not-the-password")
    assert out["error"][0] == "auth"
    assert "user name and password" in out["error"][1]
