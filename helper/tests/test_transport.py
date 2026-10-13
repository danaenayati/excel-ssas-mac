import pytest

from cubebridge.errors import CubeError, ProtocolError
from cubebridge.ssas import dime
from cubebridge.ssas.transport import MessageReader, send_message


class ScriptedChannel:
    """Hands back the response in pieces of at most `piece` bytes."""

    def __init__(self, response: bytes, piece: int = 65536):
        self.response = response
        self.piece = piece
        self.pos = 0
        self.recv_calls = 0
        self.sent = bytearray()

    def send(self, data):
        self.sent += data

    def recv(self, size):
        self.recv_calls += 1
        n = min(size, self.piece)
        chunk = self.response[self.pos : self.pos + n]
        self.pos += len(chunk)
        return chunk

    def close(self):
        pass


def chunked(payload: bytes, size: int) -> bytes:
    """One DIME message split into several records, as a server may send it."""
    parts = [payload[i : i + size] for i in range(0, len(payload), size)] or [b""]
    out = bytearray()
    for i, part in enumerate(parts):
        first, last = i == 0, i == len(parts) - 1
        out += dime.encode_record(
            part,
            options=dime.OPTIONS_NEGOTIATED if first else b"",
            type_=dime.TYPE_TEXT_XML if first else b"",
            mb=first,
            me=last,
            cf=not last,
            type_t=1 if first else 0,
        )
    return bytes(out)


def test_message_arriving_one_byte_at_a_time():
    reader = MessageReader(ScriptedChannel(dime.encode_record(b"<hello/>"), piece=1))
    assert reader.read_message() == b"<hello/>"


def test_two_messages_in_one_segment_are_both_delivered():
    wire = dime.encode_record(b"<one/>") + dime.encode_record(b"<two/>")
    reader = MessageReader(ScriptedChannel(wire))
    assert reader.read_message() == b"<one/>"
    assert reader.read_message() == b"<two/>"


def test_chunked_message_is_reassembled():
    payload = bytes(range(256)) * 40
    reader = MessageReader(ScriptedChannel(chunked(payload, 1000), piece=777))
    assert reader.read_message() == payload


def test_large_message_reads_in_linear_time():
    payload = b"z" * (5 * 1024 * 1024)
    channel = ScriptedChannel(chunked(payload, 60000), piece=65536)
    assert MessageReader(channel).read_message() == payload
    # one recv per ~64 KiB, not one full re-parse per recv
    assert channel.recv_calls <= len(payload) // 65536 + 200


def test_eof_mid_message_is_reported_as_unreachable():
    wire = dime.encode_record(b"<truncated/>")[:-3]
    with pytest.raises(CubeError) as info:
        MessageReader(ScriptedChannel(wire)).read_message()
    assert info.value.category == "unreachable"


def test_oversized_declared_length_is_refused_before_buffering():
    wire = bytearray(dime.encode_record(b"x"))
    wire[8:12] = (200 * 1024 * 1024).to_bytes(4, "big")
    with pytest.raises(ProtocolError):
        MessageReader(ScriptedChannel(bytes(wire)), max_message=64 * 1024 * 1024).read_message()


def test_server_choosing_compression_is_refused():
    wire = dime.encode_record(b"garbage", options=bytes([dime.OPT_NEGO | dime.OPT_RESP_XPRESS, 0, 0, 0]))
    with pytest.raises(ProtocolError):
        MessageReader(ScriptedChannel(wire)).read_message()


def test_send_message_writes_one_record_with_the_given_options():
    channel = ScriptedChannel(b"")
    send_message(channel, b"<q/>", dime.OPTIONS_NEGOTIATED)
    header = dime.decode_header(bytes(channel.sent[:12]))
    assert (header.mb, header.me, header.data_len) == (True, True, 4)
    assert bytes(channel.sent[12:16]) == dime.OPTIONS_NEGOTIATED
