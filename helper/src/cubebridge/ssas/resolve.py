"""Turn what the user types as "server" into where to connect.

Accepted forms are the ones Windows Excel accepts for an on-premises server:
`host`, `host:port` and `host\\instance` for the TCP binding (`Target`), and the
http:// or https:// address of an msmdpump.dll endpoint (`HttpTarget`).

A default instance listens on 2383. A named instance uses a port of its own, which
the SQL Server Browser service on TCP 2382 hands out. Microsoft does not publish that
exchange. The lookup here sends the request the ADOMD.NET client is understood to
send (an unauthenticated DISCOVER_INSTANCES) and has NOT been verified against a real
server, so a failure always tells the user the reliable alternative: `server:port`.
"""

from __future__ import annotations

import ipaddress
import re
from collections.abc import Callable
from dataclasses import dataclass
from urllib.parse import urlsplit

from ..errors import CubeError
from . import dime, envelopes, sealing, xmla
from .transport import MessageReader, SocketChannel, send_message

DEFAULT_PORT = 2383
REDIRECTOR_PORT = 2382

_LABEL = r"[A-Za-z0-9_](?:[A-Za-z0-9_-]{0,61}[A-Za-z0-9_])?"
_HOST = re.compile(rf"{_LABEL}(?:\.{_LABEL})*")
_INSTANCE = re.compile(r"[A-Za-z0-9_$#-]{1,64}")
_URL = re.compile(r"^[A-Za-z][A-Za-z0-9+.-]*://")
_WEB_SCHEMES = {"http": 80, "https": 443}


@dataclass(frozen=True)
class Target:
    host: str
    port: int | None = None
    instance: str | None = None


@dataclass(frozen=True)
class HttpTarget:
    """An XMLA endpoint over HTTP, such as https://bi.corp.local/olap/msmdpump.dll."""

    scheme: str  # "http" or "https"
    host: str
    port: int
    path: str  # path and query, always starting with "/"

    @property
    def secure(self) -> bool:
        return self.scheme == "https"

    @property
    def url(self) -> str:
        host = f"[{self.host}]" if ":" in self.host else self.host
        default = _WEB_SCHEMES[self.scheme] == self.port
        return f"{self.scheme}://{host}{'' if default else f':{self.port}'}{self.path}"


def _bad(message: str) -> CubeError:
    return CubeError("bad-request", message)


def is_ip_address(host: str) -> bool:
    try:
        ipaddress.ip_address(host)
        return True
    except ValueError:
        return False


def parse_target(text: str) -> Target | HttpTarget:
    """What the user typed as "server": a server name (TCP) or a web address (HTTP)."""
    text = (text or "").strip()
    if not text:
        raise _bad("Enter the server name.")
    if _URL.match(text) and text.split(":", 1)[0].lower() in _WEB_SCHEMES:
        return parse_url(text)
    if _URL.match(text):
        raise _bad(
            "That kind of web address is not supported. Enter the server name as in Windows "
            "Excel, or the http:// or https:// address of msmdpump.dll."
        )
    return parse_server(text)


def parse_url(text: str) -> HttpTarget:
    text = (text or "").strip()
    if any(ch <= " " or ch == "\x7f" for ch in text):
        raise _bad("The web address must not contain spaces.")
    try:
        parts = urlsplit(text)
        port = parts.port
    except ValueError as exc:
        raise _bad("The port in the web address must be a number between 1 and 65535.") from exc
    scheme = parts.scheme.lower()
    if scheme not in _WEB_SCHEMES:
        raise _bad("The web address must start with http:// or https://.")
    if parts.username is not None or parts.password is not None:
        raise _bad("Put the user name and password in their own boxes, not in the web address.")
    if parts.fragment or text.endswith("#"):
        raise _bad("The web address must not contain '#'.")
    host = parts.hostname or ""
    if not host or (not is_ip_address(host) and (len(host) > 253 or not _HOST.fullmatch(host))):
        raise _bad("The server name in the web address is not valid.")
    if port == 0:
        raise _bad("The port in the web address must be a number between 1 and 65535.")
    path = parts.path or "/"
    if parts.query:
        path += "?" + parts.query
    return HttpTarget(scheme=scheme, host=host, port=port or _WEB_SCHEMES[scheme], path=path)


def parse_server(text: str) -> Target:
    text = (text or "").strip()
    if not text:
        raise _bad("Enter the server name.")
    if _URL.match(text):
        raise _bad(
            "That is a web address. This add-in connects to company server names "
            "(for example bi-server or bi-server\\INSTANCE)."
        )
    host, port, instance = text, None, None
    if "\\" in text:
        host, _, instance = text.partition("\\")
        if not _INSTANCE.fullmatch(instance):
            raise _bad("The instance name after the backslash is not valid.")
    elif ":" in text:
        host, _, port_text = text.partition(":")
        if not port_text.isdigit() or not 1 <= int(port_text) <= 65535:
            raise _bad("The port after the colon must be a number between 1 and 65535.")
        port = int(port_text)
    if len(host) > 253 or not _HOST.fullmatch(host):
        raise _bad("The server name is not valid. Use the name you type in Windows Excel.")
    return Target(host=host, port=port, instance=instance)


def endpoint(target: Target, lookup: Callable[[str, str], int] | None = None) -> tuple[str, int]:
    if target.port:
        return target.host, target.port
    if target.instance:
        return target.host, (lookup or instance_port)(target.host, target.instance)
    return target.host, DEFAULT_PORT


def _instance_error(host: str, instance: str, detail: str) -> CubeError:
    return CubeError(
        "instance",
        f"Couldn't find the port of instance '{instance}' on {host}. Ask a Windows colleague "
        f"for the port and enter the server as {host}:port.",
        detail,
    )


def port_from_reply(reply: bytes, instance: str) -> int:
    try:
        rows = xmla.parse(reply).rows
    except CubeError as exc:
        raise _instance_error("the server", instance, exc.detail or exc.message) from exc
    for row in rows:
        if row.get("INSTANCE_NAME", "").lower() == instance.lower():
            port = row.get("INSTANCE_PORT_NUMBER", "")
            if port.isdigit() and 1 <= int(port) <= 65535:
                return int(port)
    raise _instance_error("the server", instance, "the instance is not in the server's list")


def instance_port(host: str, instance: str, redirector_port: int = REDIRECTOR_PORT, timeout: float = 8.0) -> int:
    try:
        channel = SocketChannel(host, redirector_port, timeout)
    except CubeError as exc:
        raise _instance_error(host, instance, exc.detail or exc.message) from exc
    try:
        channel.set_timeout(timeout)
        request = envelopes.discover("DISCOVER_INSTANCES", {"INSTANCE_NAME": instance})
        send_message(channel, sealing.BOM + request, dime.OPTIONS_FIRST)
        reply = MessageReader(channel, max_message=1024 * 1024).read_message()
    except CubeError as exc:
        raise _instance_error(host, instance, exc.detail or exc.message) from exc
    finally:
        channel.close()
    try:
        return port_from_reply(reply, instance)
    except CubeError as exc:
        raise _instance_error(host, instance, exc.detail or exc.message) from exc
