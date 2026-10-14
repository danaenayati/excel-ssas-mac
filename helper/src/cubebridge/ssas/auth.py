"""How a connection signs in: the three kinds and where each one is allowed.

- "ntlm": user name and password, NTLM tokens (the default; connections saved before
  sign-in kinds existed have it).
- "kerberos": Windows' Negotiate with a Kerberos ticket, the one already on this Mac
  or one got with the saved user name and password (kerberos.py).
- "basic": user name and password in the Authorization header of every request. HTTP
  only, and only over https, since the password itself crosses the network.

The labels are the pane's own wording (addin/src/state/signIn.ts), so an error can
say exactly which choice to pick.
"""

from __future__ import annotations

from ..errors import CubeError
from .resolve import HttpTarget, Target

SIGN_INS = ("ntlm", "kerberos", "basic")
DEFAULT_SIGN_IN = "ntlm"

LABELS = {
    "ntlm": "User name and password (NTLM)",
    "kerberos": "Windows Authentication (Kerberos)",
    "basic": "User name and password over HTTPS (Basic)",
}


def read_sign_in(value: object) -> str:
    """The sign-in kind of a request or a saved connection. Absent means NTLM."""
    if value is None or value == "":
        return DEFAULT_SIGN_IN
    if not isinstance(value, str) or value not in SIGN_INS:
        raise CubeError("bad-request", "'signIn' must be one of: ntlm, kerberos, basic.")
    return value


def check_allowed(sign_in: str, target: Target | HttpTarget) -> None:
    """Refuse a combination that would send the password where it must not go."""
    if sign_in != "basic":
        return
    if not isinstance(target, HttpTarget):
        raise CubeError(
            "bad-request",
            f'Basic sign-in is only for a web address (https://…/msmdpump.dll). For a server name, '
            f'choose "{LABELS["ntlm"]}" or "{LABELS["kerberos"]}".',
        )
    if not target.secure:
        raise CubeError(
            "bad-request",
            "Basic sign-in sends the password with every request, so it needs an https:// address. "
            f'Use https://, or choose "{LABELS["ntlm"]}".',
        )


def ntlm_context(username: str, password: str, host: str, service: str, channel_bindings=None):
    import spnego

    return spnego.client(
        username=username,
        password=password,
        hostname=host,
        service=service,
        protocol="ntlm",
        channel_bindings=channel_bindings,
    )
