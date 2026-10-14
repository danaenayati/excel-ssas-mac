"""Kerberos sign-in ("Windows Authentication") through this Mac's own GSS library.

Used by both bindings: the TCP Authenticate exchange (service MSOLAPSvc.3) and HTTP
Negotiate (service HTTP). Before anything is sent to the cube server, `client_context`
settles everything that only needs the company's Kerberos server (the KDC):

1. Credentials. A ticket already on this Mac (kinit, Ticket Viewer, a company sign-in
   extension) is used when it belongs to the saved user. Otherwise one is got with the
   saved user name and password, into a memory cache only this process sees, and kept
   there while it lasts. A wrong password stops at once: trying further realms would
   only count more failed sign-ins against the account.
2. The service's Kerberos name (SPN). Windows asks for MSOLAPSvc.3/<host> (with
   :<instance> for a named instance), <host> being the server's full DNS name. That name
   is looked up here (an IP address in reverse), and the ticket for it is asked for in
   the realm of the server's domain and then in the user's own.

The context returned has made its first token already; Session sends it as usual.

Tokens are wrapped in SPNEGO that offers Kerberos alone, as Windows' Negotiate does
when it uses Kerberos. The server can never pick NTLM out of it, so nothing falls back
to sending the password another way. Every failure becomes a plain sentence saying
what to do, and the ones Kerberos cannot get past say to choose NTLM.

Needs the `gssapi` and `krb5` packages, which have ready-made wheels for both kinds of
Mac and are installed with the helper. The parts of pyspnego used here that are not its
public API are kept to this module (`_KerberosContext`, `_password_credential`,
`_spnego`); pyproject.toml pins pyspnego's minor version for that reason.
"""

from __future__ import annotations

import hashlib
import hmac
import logging
import secrets
import socket
import threading
from dataclasses import dataclass

from spnego._gss import GSSAPIProxy

from ..errors import CubeError
from .auth import LABELS
from .resolve import is_ip_address

log = logging.getLogger("cubebridge.ssas")

KRB5_MECH = "1.2.840.113554.1.2.2"
_CHOOSE_NTLM = f'choose "{LABELS["ntlm"]}"'
MAX_REALMS = 3
# Failures that are about the user's credential, not about the name asked for.
_ABOUT_THE_USER = frozenset({"wrong-password", "expired", "locked", "clock"})

# Kerberos error codes. MIT and Heimdal number these the same (ERROR_TABLE_BASE_krb5).
_BASE = -1765328384
_CODES = {
    _BASE + 6: "no-user",  # KDC_ERR_C_PRINCIPAL_UNKNOWN
    _BASE + 7: "no-service",  # KDC_ERR_S_PRINCIPAL_UNKNOWN
    _BASE + 18: "locked",  # KDC_ERR_CLIENT_REVOKED
    _BASE + 23: "expired",  # KDC_ERR_KEY_EXPIRED
    _BASE + 24: "wrong-password",  # KDC_ERR_PREAUTH_FAILED
    _BASE + 31: "wrong-password",  # KRB_AP_ERR_BAD_INTEGRITY: the reply does not decrypt with it
    _BASE + 37: "clock",  # KRB_AP_ERR_SKEW
    _BASE + 154: "no-kdc",  # KRB5_REALM_UNKNOWN
    _BASE + 156: "no-kdc",  # KRB5_KDC_UNREACH
    _BASE + 220: "no-kdc",  # KRB5_REALM_CANT_RESOLVE
}
# The same, recognised by the library's wording when no code came with the error.
# Each alternative is a group of fragments that must all appear.
_TEXTS: tuple[tuple[str, tuple[tuple[str, ...], ...]], ...] = (
    ("clock", (("clock skew",),)),
    ("expired", (("password has expired",), ("password expired",), ("key has expired",))),
    ("locked", (("revoked",), ("locked out",), ("account disabled",))),
    (
        "wrong-password",
        (("preauthentication failed",), ("pre-authentication failed",), ("password incorrect",), ("incorrect password",)),
    ),
    ("no-user", (("client not found",), ("client (", "unknown"), ("client '", "not found"))),
    ("no-service", (("server not found",), ("server (", "unknown"), ("server '", "not found"))),
    (
        "no-kdc",
        (("any kdc",), ("find kdc",), ("kdc for realm",), ("kdc in realm",), ("network address for kdc",), ("realm unknown",)),
    ),
    (
        "no-ticket",
        (("no credentials",), ("credentials cache",), ("cache not found",), ("missingcredentials",), ("expiredcredentials",)),
    ),
)


@dataclass
class Choice:
    """What `client_context` settled on: shown by the probe, kept in the log."""

    credential: str  # "ticket" (already on this Mac) or "password" (got just now)
    realm: str
    service: str  # the service principal, e.g. MSOLAPSvc.3/bi.corp.local@CORP.LOCAL


@dataclass(eq=False)
class _Failed(Exception):
    reason: str
    error: BaseException | None = None
    where: str = ""  # the realm or service name it happened with


# -- names ---------------------------------------------------------------------------
def split_user(username: str) -> tuple[str, str]:
    """'CORP\\dana' -> ('dana', 'CORP'); 'dana@corp.local' -> ('dana', 'corp.local')."""
    if "\\" in username:
        domain, _, user = username.partition("\\")
        return user, domain
    if "@" in username:
        user, _, domain = username.rpartition("@")
        return user, domain
    return username, ""


def _clean(name: str) -> str:
    return (name or "").strip().rstrip(".").lower()


def _canonical_name(host: str) -> str:
    try:
        infos = socket.getaddrinfo(host, None, type=socket.SOCK_STREAM, flags=socket.AI_CANONNAME)
    except (OSError, UnicodeError):
        return ""
    return next((info[3] for info in infos if info[3]), "")


def _reverse_name(address: str) -> str:
    try:
        return socket.gethostbyaddr(address)[0]
    except (OSError, UnicodeError):
        return ""


def service_hosts(host: str) -> list[str]:
    """The host part of the service's Kerberos name: the full DNS name first, as Windows
    uses it, then the name as typed."""
    if is_ip_address(host):
        name = _clean(_reverse_name(host))
        if not name or is_ip_address(name):
            raise CubeError(
                "auth-mechanism",
                "Kerberos needs the server's name, not its IP address, and this Mac could not "
                "find the name for that address. Enter the server's full name (for example "
                f"bi.corp.example.com), or {_CHOOSE_NTLM}.",
                "no reverse DNS name for the server's address",
            )
        return [name]
    names: list[str] = []
    for name in (_clean(_canonical_name(host)), _clean(host)):
        if name and not is_ip_address(name) and name not in names:
            names.append(name)
    return names


def _domain_realm(host: str) -> str:
    return host.split(".", 1)[1].upper() if "." in host and not is_ip_address(host) else ""


def service_names(service: str, hosts: list[str], instance: str | None, client_realm: str) -> list[str]:
    names: list[str] = []
    for host in hosts:
        base = f"{service}/{host}" + (f":{instance}" if instance else "")
        for realm in (_domain_realm(host), client_realm):
            if realm and f"{base}@{realm}" not in names:
                names.append(f"{base}@{realm}")
    return names


def realm_candidates(username: str, hosts: list[str]) -> list[str]:
    """Realms to get a ticket in: the domain of a user@domain name; else the server's
    domain and the domains above it (bi.corp.example.com: CORP.EXAMPLE.COM, EXAMPLE.COM)."""
    _user, domain = split_user(username)
    found: list[str] = []
    if domain and ("@" in username or "." in domain):
        found.append(domain.upper())
    else:
        for host in hosts:
            realm = _domain_realm(host)
            while "." in realm:
                found.append(realm)
                realm = realm.split(".", 1)[1]
    unique = list(dict.fromkeys(found))
    return unique[:MAX_REALMS]


def ticket_matches(principal: str, username: str) -> bool:
    """Whether a ticket for `principal` is the saved user's own."""
    ticket_user, _, ticket_realm = principal.rpartition("@")
    user, domain = split_user(username)
    if not ticket_user or ticket_user.lower() != user.lower():
        return False
    if "@" in username and domain:
        return ticket_realm.upper() == domain.upper()
    return True


# -- errors --------------------------------------------------------------------------
def _facts(exc: BaseException) -> tuple[list[int], str]:
    """Error codes and text from an exception and everything it wraps."""
    codes: list[int] = []
    texts: list[str] = []
    seen: set[int] = set()
    stack: list[BaseException | None] = [exc]
    while stack:
        error = stack.pop()
        if error is None or id(error) in seen:
            continue
        seen.add(id(error))
        texts.append(type(error).__name__)
        for name in ("code", "min_code"):
            value = getattr(error, name, None)
            if isinstance(value, int) and not isinstance(value, bool):
                codes.append(value - 2**32 if value >= 2**31 else value)
        try:
            texts.append(str(error))
        except Exception:  # noqa: BLE001 - a broken __str__ must not hide the real error
            pass
        stack.extend([getattr(error, "base_error", None), error.__cause__, error.__context__])
    return codes, " ".join(texts).lower()


def classify(exc: BaseException) -> str:
    codes, text = _facts(exc)
    for code in codes:
        if code in _CODES:
            return _CODES[code]
    for reason, alternatives in _TEXTS:
        if any(all(part in text for part in group) for group in alternatives):
            return reason
    return "other"


def _failure(failed: _Failed, *, service: str, instance: str | None) -> CubeError:
    reason, where = failed.reason, failed.where
    kind = type(failed.error).__name__ if failed.error is not None else ""
    detail = "; ".join(part for part in (f"Kerberos: {reason}", kind, where) if part)
    if reason == "wrong-password":
        return CubeError("auth", "Kerberos refused the password for this user name. Check it: it is your Windows password.", detail)
    if reason == "no-user":
        return CubeError(
            "auth",
            "Kerberos does not know this user name. Enter it as name@company-domain "
            "(for example dana@corp.example.com).",
            detail,
        )
    if reason == "expired":
        return CubeError(
            "auth",
            "Your Windows password has expired. Change it (on a Windows PC or the company's "
            "password page), then enter the new one here.",
            detail,
        )
    if reason == "locked":
        return CubeError("auth", "Kerberos says this account is locked or disabled. Ask IT to unlock it.", detail)
    if reason == "clock":
        return CubeError(
            "auth",
            "This Mac's clock is more than 5 minutes off the company's, which Kerberos does not "
            "allow. Turn on Set time and date automatically (System Settings > General > Date & "
            "Time) and try again.",
            detail,
        )
    if reason == "no-kdc":
        return CubeError(
            "auth-mechanism",
            "This Mac cannot reach the company's Kerberos server"
            + (f" for {where}" if where and "/" not in where else "")
            + f". Kerberos needs the company network or VPN. Connect it and try again, or {_CHOOSE_NTLM}.",
            detail,
        )
    if reason == "no-service":
        what = "web server" if service == "HTTP" else "Analysis Services server"
        hint = "" if instance or service == "HTTP" else " If the cube runs as a named instance, enter the server as name\\INSTANCE."
        return CubeError(
            "auth-mechanism",
            f"Kerberos does not know this {what} (it has no Kerberos name {service}/…), so it "
            f"cannot sign in to it. Ask IT to register one, or {_CHOOSE_NTLM}.{hint}",
            detail,
        )
    return CubeError("auth", f"Kerberos sign-in failed. Check the user name and password, or {_CHOOSE_NTLM}.", detail)


def _not_installed(detail: str) -> CubeError:
    return CubeError(
        "auth-mechanism",
        f"Windows Authentication (Kerberos) is not installed on this Mac. Run install.sh again "
        f"to add it, or {_CHOOSE_NTLM}.",
        detail,
    )


# -- the GSS side --------------------------------------------------------------------
def _modules():
    try:
        import gssapi
        from spnego import _gss
    except ImportError as exc:
        raise _not_installed(str(exc)) from exc
    if not (_gss.HAS_GSSAPI and _gss.HAS_IOV):
        raise _not_installed(_gss.GSSAPI_IMP_ERR or _gss.GSSAPI_IOV_IMP_ERR or "gssapi without IOV")
    return gssapi, _gss


def check_installed() -> None:
    """Raise the plain "not installed" error when this Mac cannot do Kerberos."""
    _modules()


def existing_ticket():
    """(credential, principal) of a valid Kerberos ticket already on this Mac, or None."""
    gssapi, _ = _modules()
    try:
        credential = gssapi.Credentials(usage="initiate", mechs=[gssapi.OID.from_int_seq(KRB5_MECH)])
        if not credential.lifetime:  # raises when the ticket has expired
            return None
        return credential, str(credential.name)
    except Exception:  # noqa: BLE001 - no ticket, an expired one, or no cache at all
        return None


def _password_credential(principal: str, password: str):
    """A ticket got with the password, in a memory cache only this process can see."""
    gssapi, gss = _modules()
    raw = gss._kinit(principal.encode("utf-8"), gss._encode_kerb_password(password))
    return gssapi.Credentials(base=raw, usage="initiate")


# Tickets got with a password, kept while they last (ten hours, usually), so a new
# connection does not ask the KDC again with the password. Keyed by user and by a keyed
# hash of the password, so a changed password never gets the old ticket.
_tickets: dict[str, tuple[bytes, object]] = {}
_tickets_lock = threading.Lock()
_TICKET_KEY = secrets.token_bytes(32)
_MIN_LIFETIME = 300  # seconds a kept ticket must still have


def _ticket_for(principal: str, password: str):
    digest = hmac.new(_TICKET_KEY, password.encode("utf-8", "surrogatepass"), hashlib.sha256).digest()
    with _tickets_lock:
        kept = _tickets.get(principal)
    if kept is not None and hmac.compare_digest(kept[0], digest):
        try:
            if (kept[1].lifetime or 0) > _MIN_LIFETIME:
                return kept[1]
        except Exception:  # noqa: BLE001 - expired or unusable: get a new one
            pass
    credential = _password_credential(principal, password)
    with _tickets_lock:
        _tickets[principal] = (digest, credential)
    return credential


class _KerberosContext(GSSAPIProxy):
    """pyspnego's Kerberos context, aimed at an exact service principal.

    pyspnego names the target `service@host`, which macOS turns into a principal in the
    realm named after the host's domain. That cannot express MSOLAPSvc.3/host:INSTANCE
    (the realm would come out as "DOMAIN:INSTANCE") nor try the user's own realm, so the
    GSS context is made here with the principal chosen above. The tokens, wrapping and
    unwrapping are pyspnego's.
    """

    def __init__(self, credential, target: str, channel_bindings=None) -> None:
        service, _, rest = target.partition("/")
        super().__init__(
            hostname=rest.split("@", 1)[0],
            service=service,
            channel_bindings=channel_bindings,
            protocol="kerberos",
            _gssapi_credential=credential,
        )
        self.target = target

    def step(self, in_token: bytes | None = None, *, channel_bindings=None) -> bytes | None:
        if self._context is None:
            import gssapi
            from gssapi.raw import ChannelBindings

            extra = {}
            bindings = channel_bindings or self.channel_bindings
            if bindings:
                extra["channel_bindings"] = ChannelBindings(
                    initiator_address_type=bindings.initiator_addrtype,
                    initiator_address=bindings.initiator_address,
                    acceptor_address_type=bindings.acceptor_addrtype,
                    acceptor_address=bindings.acceptor_address,
                    application_data=bindings.application_data,
                )
            self._context = gssapi.SecurityContext(
                name=gssapi.Name(self.target, name_type=gssapi.NameType.kerberos_principal),
                creds=self._credential,
                usage="initiate",
                mech=self._mech,
                flags=self._context_req,
                **extra,
            )
        return super().step(in_token, channel_bindings=channel_bindings)


def _new_context(credential, target: str, channel_bindings):
    return _KerberosContext(credential, target, channel_bindings)


class _Primed:
    """A context whose first token was made already: `step()` hands that out first."""

    def __init__(self, context, token: bytes) -> None:
        self._inner = context
        self._token: bytes | None = token

    def step(self, in_token: bytes | None = None, *, channel_bindings=None) -> bytes | None:
        if self._token is not None:
            token, self._token = self._token, None
            return token
        return self._inner.step(in_token, channel_bindings=channel_bindings)

    def __getattr__(self, name: str):
        return getattr(self._inner, name)


def _spnego(context, host: str, service: str):
    """SPNEGO around the Kerberos context, offering Kerberos and nothing else."""
    from spnego._context import GSSMech
    from spnego._negotiate import NegotiateProxy

    return NegotiateProxy(
        hostname=host,
        service=service,
        protocol="negotiate",
        _negotiate_contexts={GSSMech.kerberos: context},
    )


# -- the whole sign-in ---------------------------------------------------------------
def _start(credential, realm: str, source: str, hosts, service, instance, channel_bindings, wrap):
    """Ask for the service ticket under each candidate name; return the first that works."""
    misses: list[_Failed] = []
    for target in service_names(service, hosts, instance, realm):
        context = _new_context(credential, target, channel_bindings)
        try:
            token = context.step()
        except Exception as exc:  # noqa: BLE001 - classified below; the text is never shown
            failed = _Failed(classify(exc), exc, target)
            log.debug("Kerberos: no ticket for %s (%s)", target, failed.reason)
            if failed.reason in _ABOUT_THE_USER:
                raise failed from exc  # another name for the server would not change that
            # An unknown name, no way to the realm it was asked in, or a cross-realm
            # failure in words not recognised here: try the next name.
            misses.append(failed)
            continue
        if not token:
            raise _Failed("other", None, target)
        log.debug("Kerberos: ticket for %s (%s)", target, "already on this Mac" if source == "ticket" else "got with the password")
        ready = _Primed(context, token)
        result = _spnego(ready, hosts[0], service) if wrap else ready
        result.cube_choice = Choice(source, realm, target)
        return result
    # A name the KDC does not know says more than a realm (guessed from the host) without
    # a KDC, and that more than anything else: report the most telling one.
    for reason in ("no-service", "no-kdc"):
        for miss in misses:
            if miss.reason == reason:
                raise miss
    raise (misses or [_Failed("no-service")])[0]


def client_context(
    username: str,
    password: str,
    host: str,
    *,
    service: str,
    instance: str | None = None,
    channel_bindings=None,
    wrap: bool = True,
):
    """A ready Kerberos (SPNEGO) context for `service` on `host`; see the module text."""
    _modules()
    hosts = service_hosts(host)
    unknown_service: _Failed | None = None

    ticket = existing_ticket()
    if ticket is not None and ticket_matches(ticket[1], username):
        credential, principal = ticket
        log.debug("Kerberos: using the ticket already on this Mac")
        try:
            return _start(credential, principal.rpartition("@")[2].upper(), "ticket", hosts, service, instance, channel_bindings, wrap)
        except _Failed as failed:
            if failed.reason == "no-service":
                unknown_service = failed
            log.debug("Kerberos: the ticket on this Mac did not work (%s); trying the password", failed.reason)

    realms = realm_candidates(username, hosts)
    if not realms:
        raise CubeError(
            "auth",
            "Kerberos needs your company's domain. Enter the user name as name@company-domain "
            "(for example dana@corp.example.com).",
            "no realm: the user name has no domain and the server name has none either",
        )
    user, _domain = split_user(username)
    best: _Failed | None = None
    for realm in realms:
        log.debug("Kerberos: getting a ticket in %s with the password", realm)
        try:
            credential = _ticket_for(f"{user}@{realm}", password)
        except Exception as exc:  # noqa: BLE001 - classified; the text is never shown
            failed = _Failed(classify(exc), exc, realm)
            log.debug("Kerberos: no ticket in %s (%s)", realm, failed.reason)
            if failed.reason not in ("no-user", "no-kdc"):
                raise _failure(failed, service=service, instance=instance) from exc
            # A realm with a KDC that does not know the user says more than one without a KDC.
            if best is None or (best.reason == "no-kdc" and failed.reason == "no-user"):
                best = failed
            continue
        try:
            return _start(credential, realm, "password", hosts, service, instance, channel_bindings, wrap)
        except _Failed as failed:
            raise _failure(failed, service=service, instance=instance) from failed.error
    if unknown_service is not None and (best is None or best.reason == "no-kdc"):
        best = unknown_service
    assert best is not None
    raise _failure(best, service=service, instance=instance) from best.error
