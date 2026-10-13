import pytest

from cubebridge.errors import ProtocolError
from cubebridge.ssas import dime


def test_header_matches_the_ms_ssas_worked_example():
    # [MS-SSAS] "Authentication" example: VERSION 1, MB and ME set, TYPE_T 1,
    # OPTIONS_LENGTH 4, ID_LENGTH 0, TYPE_LENGTH 8.
    record = dime.encode_record(b"<x/>")
    assert record[:8] == bytes.fromhex("0E10000400000008")
    assert record[8:12] == (4).to_bytes(4, "big")


def test_fields_are_padded_to_four_bytes():
    record = dime.encode_record(b"abcde")  # 5 data bytes -> 3 pad
    # header 12 + options 4 + type "text/xml" 8 + data 5 + pad 3
    assert len(record) == 12 + 4 + 8 + 8
    assert record.endswith(b"abcde\x00\x00\x00")


def test_header_round_trip():
    record = dime.encode_record(b"payload", options=dime.OPTIONS_NEGOTIATED)
    header = dime.decode_header(record[: dime.HEADER_LEN])
    assert (header.mb, header.me, header.cf, header.type_t) == (True, True, False, 1)
    assert (header.options_len, header.id_len, header.type_len, header.data_len) == (4, 0, 8, 7)


def test_continuation_record_flags():
    first = dime.encode_record(b"aa", mb=True, me=False, cf=True)
    last = dime.encode_record(b"bb", mb=False, me=True, cf=False, type_=b"", type_t=0)
    h1 = dime.decode_header(first[:12])
    h2 = dime.decode_header(last[:12])
    assert (h1.mb, h1.me, h1.cf) == (True, False, True)
    assert (h2.mb, h2.me, h2.cf, h2.type_t, h2.type_len) == (False, True, False, 0, 0)


def test_wrong_version_is_a_protocol_error():
    bad = bytearray(dime.encode_record(b"x"))
    bad[0] = (2 << 3) | 0x06
    with pytest.raises(ProtocolError):
        dime.decode_header(bytes(bad[:12]))


def test_padding_helper():
    assert [dime.pad(n) for n in (0, 1, 2, 3, 4, 5)] == [0, 3, 2, 1, 0, 3]


@pytest.mark.parametrize(
    "options, type_",
    [
        (bytes([dime.OPT_RESP_SX, 0, 0, 0]), dime.TYPE_TEXT_XML),
        (bytes([dime.OPT_NEGO | dime.OPT_RESP_XPRESS, 0, 0, 0]), dime.TYPE_TEXT_XML),
        (dime.OPTIONS_NEGOTIATED, b"application/sx"),
        (dime.OPTIONS_NEGOTIATED, b"application/xml+xpress"),
    ],
)
def test_check_negotiated_rejects_encodings_we_do_not_speak(options, type_):
    with pytest.raises(ProtocolError):
        dime.check_negotiated(options, type_)


def test_check_negotiated_accepts_clear_text():
    dime.check_negotiated(dime.OPTIONS_NEGOTIATED, dime.TYPE_TEXT_XML)
    dime.check_negotiated(b"", dime.TYPE_TEXT_XML)
