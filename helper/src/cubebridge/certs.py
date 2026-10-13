"""The certificate that lets Excel's web view talk to https://localhost.

Office add-ins must be served over HTTPS, so the helper needs a certificate macOS
trusts. A private CA is created, used once to sign a localhost certificate, and its
private key is then discarded. The CA is also name-constrained to localhost. Trusting
it therefore cannot be turned into trust for any other site, by anyone.
"""

from __future__ import annotations

import datetime as dt
import ipaddress
import os
from dataclasses import dataclass
from pathlib import Path

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID

CA_COMMON_NAME = "CubeBridge Local CA (localhost only)"
LEAF_DAYS = 820  # macOS rejects TLS certificates valid for more than 825 days
CA_DAYS = 3650


@dataclass(frozen=True)
class CertPaths:
    ca: Path
    cert: Path
    key: Path


def paths_in(directory: Path) -> CertPaths:
    directory = Path(directory)
    return CertPaths(ca=directory / "ca.pem", cert=directory / "localhost.pem", key=directory / "localhost.key")


def generate(directory: Path) -> CertPaths:
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    now = dt.datetime.now(dt.timezone.utc) - dt.timedelta(minutes=5)

    ca_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    ca_name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, CA_COMMON_NAME)])
    ca_cert = (
        x509.CertificateBuilder()
        .subject_name(ca_name)
        .issuer_name(ca_name)
        .public_key(ca_key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now)
        .not_valid_after(now + dt.timedelta(days=CA_DAYS))
        .add_extension(x509.BasicConstraints(ca=True, path_length=0), critical=True)
        .add_extension(
            x509.KeyUsage(
                digital_signature=False, content_commitment=False, key_encipherment=False,
                data_encipherment=False, key_agreement=False, key_cert_sign=True, crl_sign=True,
                encipher_only=False, decipher_only=False,
            ),
            critical=True,
        )
        .add_extension(
            x509.NameConstraints(
                permitted_subtrees=[
                    x509.DNSName("localhost"),
                    x509.IPAddress(ipaddress.ip_network("127.0.0.1/32")),
                    x509.IPAddress(ipaddress.ip_network("::1/128")),
                ],
                excluded_subtrees=None,
            ),
            critical=True,
        )
        .add_extension(x509.SubjectKeyIdentifier.from_public_key(ca_key.public_key()), critical=False)
        .sign(ca_key, hashes.SHA256())
    )

    leaf_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    leaf_cert = (
        x509.CertificateBuilder()
        .subject_name(x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "localhost")]))
        .issuer_name(ca_name)
        .public_key(leaf_key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now)
        .not_valid_after(now + dt.timedelta(days=LEAF_DAYS))
        .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
        .add_extension(
            x509.KeyUsage(
                digital_signature=True, content_commitment=False, key_encipherment=True,
                data_encipherment=False, key_agreement=False, key_cert_sign=False, crl_sign=False,
                encipher_only=False, decipher_only=False,
            ),
            critical=True,
        )
        .add_extension(x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]), critical=False)
        .add_extension(
            x509.SubjectAlternativeName(
                [
                    x509.DNSName("localhost"),
                    x509.IPAddress(ipaddress.ip_address("127.0.0.1")),
                    x509.IPAddress(ipaddress.ip_address("::1")),
                ]
            ),
            critical=False,
        )
        .add_extension(x509.AuthorityKeyIdentifier.from_issuer_public_key(ca_key.public_key()), critical=False)
        .add_extension(x509.SubjectKeyIdentifier.from_public_key(leaf_key.public_key()), critical=False)
        .sign(ca_key, hashes.SHA256())
    )
    del ca_key  # never written anywhere

    paths = paths_in(directory)
    paths.ca.write_bytes(ca_cert.public_bytes(serialization.Encoding.PEM))
    paths.cert.write_bytes(leaf_cert.public_bytes(serialization.Encoding.PEM))
    descriptor = os.open(paths.key, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(descriptor, "wb") as handle:
        handle.write(
            leaf_key.private_bytes(
                serialization.Encoding.PEM,
                serialization.PrivateFormat.PKCS8,
                serialization.NoEncryption(),
            )
        )
    os.chmod(paths.key, 0o600)
    return paths


def days_left(cert_path: Path) -> int | None:
    try:
        cert = x509.load_pem_x509_certificate(Path(cert_path).read_bytes())
    except (OSError, ValueError):
        return None
    return (cert.not_valid_after_utc - dt.datetime.now(dt.timezone.utc)).days
