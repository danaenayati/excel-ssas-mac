import struct

import pytest

from cubebridge.errors import IncompleteMessage, ProtocolError
from cubebridge.ssas import sealing


class FakeContext:
    """XOR 'cipher' with a 16-byte token that carries a sequence number."""

    def __init__(self, padding=0):
        self.seq = 0
        self.padding = padding

    def wrap_winrm(self, data):
        token = b"T" * 12 + struct.pack("<I", self.seq)
        self.seq += 1
        return token, bytes(b ^ 0x5A for b in data), self.padding

    def unwrap_winrm(self, header, data):
        assert len(header) == 16
        return bytes(b ^ 0x5A for b in data)


def frames(blob):
    out, offset = [], 0
    while offset < len(blob):
        data_size, token_size = struct.unpack_from("<HH", blob, offset)
        offset += 4
        out.append((blob[offset : offset + data_size], blob[offset + data_size : offset + data_size + token_size]))
        offset += data_size + token_size
    return out


def test_frame_is_sizes_then_ciphertext_then_token():
    blob = sealing.seal_frame(FakeContext(), b"abc")
    assert blob[:4] == struct.pack("<HH", 3, 16)
    assert blob[4:7] == bytes(b ^ 0x5A for b in b"abc")
    assert blob[7:] == b"T" * 12 + struct.pack("<I", 0)


def test_bom_is_its_own_first_frame_and_body_is_chunked():
    payload = b"x" * (sealing.MAX_CHUNK * 2 + 10)
    parts = frames(sealing.seal_message(FakeContext(), payload))
    assert bytes(b ^ 0x5A for b in parts[0][0]) == sealing.BOM
    assert [len(p[0]) for p in parts[1:]] == [sealing.MAX_CHUNK, sealing.MAX_CHUNK, 10]
    # tokens carry consecutive sequence numbers: frames were sealed in wire order
    assert [struct.unpack("<I", p[1][12:])[0] for p in parts] == [0, 1, 2, 3]


def test_unseal_round_trip_strips_the_bom():
    payload = "<Envelope>سلام</Envelope>".encode() * 500
    blob = sealing.seal_message(FakeContext(), payload)
    assert sealing.unseal_message(FakeContext(), blob) == payload


def test_truncated_header_and_truncated_frame_are_incomplete_not_malformed():
    blob = sealing.seal_message(FakeContext(), b"hello world")
    with pytest.raises(IncompleteMessage):
        sealing.unseal_message(FakeContext(), blob[:2])
    with pytest.raises(IncompleteMessage):
        sealing.unseal_message(FakeContext(), blob[:-1])


def test_a_mechanism_that_pads_is_refused():
    with pytest.raises(ProtocolError):
        sealing.seal_frame(FakeContext(padding=4), b"abc")
