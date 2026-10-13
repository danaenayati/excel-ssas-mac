"""Turn what the user types as "server" into a host and a TCP port.

Accepted forms are the ones Windows Excel accepts for an on-premises server:
`host`, `host:port` and `host\\instance`.

A default instance listens on 2383. A named instance uses a port of its own, which
the SQL Server Browser service on TCP 2382 hands out. Microsoft does not publish that
exchange. The lookup here sends the request the ADOMD.NET client is understood to
send (an unauthenticated DISCOVER_INSTANCES) and has NOT been verified against a real
server, so a failure always tells the user the reliable alternative: `server:port`.
"""

from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass

from ..errors import CubeError
from . import dime, envelopes, sealing, xmla
from .transport import MessageReader, SocketChannel, send_message

DEFAULT_PORT = 2383
REDIRECTOR_PORT = 2382

_HOST = re.compile(r"^[A-Za-z0-9]([A-Za-z0-9._-]{0,251}[A-Za-z0-9])?$")
_INSTANCE = re.compile(r"^[A-Za-z0-9_$#-]{1,64}$")
_URL = re.compile(r"^[A-Za-z][A-Za-z0-9+.-]*://")


@dataclass(frozen=True)
class Target:
    host: str
    port: int | None = None
    instance: str | None = None


def _bad(message: str) -> CubeError:
    return CubeError("bad-request", message)


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
        if not _INSTANCE.match(instance):
            raise _bad("The instance name after the backslash is not valid.")
    elif ":" in text:
        host, _, port_text = text.partition(":")
        if not port_text.isdigit() or not 1 <= int(port_text) <= 65535:
            raise _bad("The port after the colon must be a number between 1 and 65535.")
        port = int(port_text)
    if not _HOST.match(host):
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
