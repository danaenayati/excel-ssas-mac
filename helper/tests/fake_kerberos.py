"""A stand-in Kerberos mechanism for tests: no KDC, and no cryptography worth the name.

Client and acceptor share a key. The tokens have Kerberos's shape only: an AP-REQ
naming the service, an AP-REP back, then per-message wrap results whose header is 60
bytes (the size AES-256 gives in SSPI's layout) and carries a sequence number in each
direction. Out-of-order or tampered frames are refused, as a real context would.
"""

from __future__ import annotations

import hashlib
import hmac
import struct

from spnego._context import WinRMWrapResult

HEADER_SIZE = 60
AP_REQ = b"\x60FAKE-AP-REQ "
AP_REP = b"\x60FAKE-AP-REP"
_MAGIC = b"FAKEKRB5"


def _xor(data: bytes, key: bytes, direction: bytes, seq: int) -> bytes:
    stream = b""
    block = 0
    while len(stream) < len(data):
        stream += hashlib.sha256(key + direction + struct.pack("<II", seq, block)).digest()
        block += 1
    return bytes(a ^ b for a, b in zip(data, stream))


class _Side:
    def __init__(self, key: bytes, out: bytes, into: bytes) -> None:
        self.key = key
        self._out = out
        self._in = into
        self._send_seq = 0
        self._recv_seq = 0
        self.complete = False
        self.negotiated_protocol = "kerberos"
        self._requires_mech_list_mic = False
        self.wrapped_header_sizes: list[int] = []

    # What pyspnego's NegotiateProxy asks of the context it wraps.
    def _reset_ntlm_crypto_state(self, outgoing: bool = True) -> None:
        pass

    def sign(self, data: bytes, qop=None) -> bytes:
        return hmac.new(self.key, b"mic" + data, hashlib.sha256).digest()

    def verify(self, data: bytes, mic: bytes) -> int:
        if not hmac.compare_digest(mic, self.sign(data)):
            raise ValueError("bad mechListMIC")
        return 0

    def wrap_winrm(self, data: bytes) -> WinRMWrapResult:
        seq = self._send_seq
        self._send_seq += 1
        mac = hmac.new(self.key, self._out + struct.pack("<I", seq) + data, hashlib.sha256).digest()
        header = (_MAGIC + struct.pack("<I", seq) + mac).ljust(HEADER_SIZE, b"\x00")
        self.wrapped_header_sizes.append(len(header))
        return WinRMWrapResult(header=header, data=_xor(data, self.key, self._out, seq), padding_length=0)

    def unwrap_winrm(self, header: bytes, data: bytes) -> bytes:
        if len(header) != HEADER_SIZE or not header.startswith(_MAGIC):
            raise ValueError("not a wrap token")
        seq = struct.unpack_from("<I", header, len(_MAGIC))[0]
        if seq != self._recv_seq:
            raise ValueError(f"frame {seq} out of order, expected {self._recv_seq}")
        self._recv_seq += 1
        plain = _xor(data, self.key, self._in, seq)
        mac = hmac.new(self.key, self._in + struct.pack("<I", seq) + plain, hashlib.sha256).digest()
        if not hmac.compare_digest(header[len(_MAGIC) + 4 : len(_MAGIC) + 36], mac):
            raise ValueError("frame does not verify")
        return plain


class FakeKerberosClient(_Side):
    def __init__(self, target: str, key: bytes = b"shared session key") -> None:
        super().__init__(key, b"c2s", b"s2c")
        self.target = target
        self._sent = False

    def step(self, in_token: bytes | None = None, *, channel_bindings=None) -> bytes | None:
        if not self._sent:
            self._sent = True
            self.channel_bindings = channel_bindings
            return AP_REQ + self.target.encode()
        if in_token != AP_REP:
            raise ValueError("the AP-REP does not verify")
        self.complete = True
        return None


class FakeKerberosAcceptor(_Side):
    def __init__(self, service: str = "MSOLAPSvc.3", key: bytes = b"shared session key") -> None:
        super().__init__(key, b"s2c", b"c2s")
        self.service = service
        self.target = ""

    def step(self, in_token: bytes | None = None, *, channel_bindings=None) -> bytes | None:
        if not in_token or not in_token.startswith(AP_REQ):
            raise ValueError("not a Kerberos ticket")
        target = in_token[len(AP_REQ) :].decode()
        if not target.startswith(self.service + "/"):
            raise ValueError(f"ticket for the wrong service: {target}")
        self.target = target
        self.complete = True
        return AP_REP


def spnego_acceptor(service: str = "MSOLAPSvc.3"):
    """pyspnego's own SPNEGO acceptor around the fake Kerberos one."""
    import spnego
    from spnego._context import GSSMech

    inner = FakeKerberosAcceptor(service)
    outer = spnego.server(protocol="negotiate", _negotiate_contexts={GSSMech.kerberos: inner})
    outer.fake = inner
    return outer
