import ipaddress
import json
import stat
import subprocess
import sys

import pytest
from cryptography import x509
from cryptography.x509.oid import ExtendedKeyUsageOID

from cubebridge import certs
from cubebridge.errors import CubeError
from cubebridge.store import MemorySecrets, Store


@pytest.fixture
def store(tmp_path):
    return Store(tmp_path / "app", secrets=MemorySecrets())


def test_token_is_created_once_and_survives_a_restart(tmp_path):
    first = Store(tmp_path / "app", secrets=MemorySecrets())
    token = first.token
    assert len(token) >= 32
    assert Store(tmp_path / "app", secrets=MemorySecrets()).token == token


def test_config_is_private_to_the_user(store):
    store.token
    mode = stat.S_IMODE((store.directory / "config.json").stat().st_mode)
    assert mode == 0o600
    assert stat.S_IMODE(store.directory.stat().st_mode) == 0o700


def test_demo_connection_always_exists_and_cannot_be_deleted(store):
    assert store.connections() == [
        {"id": "demo", "name": "Sample cube", "kind": "demo", "server": "(built in)", "username": ""}
    ]
    with pytest.raises(CubeError):
        store.delete("demo")


def test_add_get_delete_and_password_stays_out_of_the_file(store):
    added = store.add(name="", server="bi-server", username="CORP\\dana", password="hunter2-secret")
    assert added["kind"] == "ssas" and added["name"] == "bi-server" and added["server"] == "bi-server"
    assert "password" not in added
    assert store.get(added["id"]) == added
    assert store.password(added["id"]) == "hunter2-secret"
    text = (store.directory / "config.json").read_text()
    assert "hunter2-secret" not in text and "CORP\\\\dana" in text
    assert [c["id"] for c in store.connections()] == ["demo", added["id"]]

    store.delete(added["id"])
    assert [c["id"] for c in store.connections()] == ["demo"]
    with pytest.raises(CubeError) as info:
        store.get(added["id"])
    assert info.value.category == "not-found"
    with pytest.raises(CubeError):
        store.password(added["id"])


def test_connections_survive_a_restart_and_ids_are_unique(tmp_path):
    secrets = MemorySecrets()
    first = Store(tmp_path / "app", secrets=secrets)
    a = first.add(name="A", server="srv", username="u", password="p1")
    b = first.add(name="B", server="srv", username="u", password="p2")
    assert a["id"] != b["id"]
    again = Store(tmp_path / "app", secrets=secrets)
    assert [c["name"] for c in again.connections()] == ["Sample cube", "A", "B"]
    assert again.password(b["id"]) == "p2"


def test_a_corrupt_config_is_reported_not_silently_replaced(tmp_path):
    directory = tmp_path / "app"
    directory.mkdir()
    (directory / "config.json").write_text("{not json")
    with pytest.raises(CubeError) as info:
        Store(directory, secrets=MemorySecrets()).token
    assert info.value.category == "internal"
    assert (directory / "config.json").read_text() == "{not json"


# -- certificates ------------------------------------------------------------------
def test_certificate_is_for_localhost_only_and_the_ca_key_is_not_kept(tmp_path):
    paths = certs.generate(tmp_path)
    assert sorted(p.name for p in tmp_path.iterdir()) == ["ca.pem", "localhost.key", "localhost.pem"]
    assert stat.S_IMODE(paths.key.stat().st_mode) == 0o600

    leaf = x509.load_pem_x509_certificate(paths.cert.read_bytes())
    san = leaf.extensions.get_extension_for_class(x509.SubjectAlternativeName).value
    assert san.get_values_for_type(x509.DNSName) == ["localhost"]
    assert ipaddress.ip_address("127.0.0.1") in san.get_values_for_type(x509.IPAddress)
    eku = leaf.extensions.get_extension_for_class(x509.ExtendedKeyUsage).value
    assert list(eku) == [ExtendedKeyUsageOID.SERVER_AUTH]
    assert leaf.extensions.get_extension_for_class(x509.BasicConstraints).value.ca is False
    lifetime = leaf.not_valid_after_utc - leaf.not_valid_before_utc
    assert lifetime.days <= 825  # Apple rejects longer-lived TLS certificates

    ca = x509.load_pem_x509_certificate(paths.ca.read_bytes())
    constraints = ca.extensions.get_extension_for_class(x509.NameConstraints).value
    assert [n.value for n in constraints.permitted_subtrees if isinstance(n, x509.DNSName)] == ["localhost"]
    assert leaf.issuer == ca.subject
    assert b"PRIVATE KEY" not in paths.ca.read_bytes()
    assert 815 <= certs.days_left(paths.cert) <= 820


@pytest.mark.skipif(sys.platform != "darwin", reason="uses the macOS trust evaluator")
def test_macos_accepts_the_chain_for_localhost_and_rejects_other_names(tmp_path):
    paths = certs.generate(tmp_path)

    def verify(name):
        return subprocess.run(
            ["/usr/bin/security", "verify-cert", "-c", str(paths.cert), "-r", str(paths.ca), "-p", "ssl", "-s", name],
            capture_output=True,
            text=True,
        )

    ok = verify("localhost")
    assert ok.returncode == 0, ok.stderr
    assert verify("example.com").returncode != 0


def test_days_left_is_none_for_a_missing_file(tmp_path):
    assert certs.days_left(tmp_path / "nope.pem") is None


def test_generated_config_is_valid_json(store):
    store.add(name="A", server="srv", username="u", password="p")
    data = json.loads((store.directory / "config.json").read_text())
    assert set(data) == {"token", "connections"}


# -- review fixes --------------------------------------------------------------------
def test_two_processes_sharing_the_settings_file_do_not_lose_each_others_changes(tmp_path):
    secrets = MemorySecrets()
    helper = Store(tmp_path / "app", secrets=secrets)
    helper.token  # the long-running helper has read the file
    probe = Store(tmp_path / "app", secrets=secrets)
    saved = probe.add(name="From probe", server="srv", username="u", password="p")

    assert [c["name"] for c in helper.connections()] == ["Sample cube", "From probe"]
    helper.add(name="From Excel", server="srv2", username="u", password="p")
    names = [c["name"] for c in Store(tmp_path / "app", secrets=secrets).connections()]
    assert names == ["Sample cube", "From probe", "From Excel"]
    assert helper.password(saved["id"]) == "p"
    assert helper.token == probe.token
