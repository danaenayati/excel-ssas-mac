"""Kerberos sign-in without a KDC: the KDC and this Mac's ticket cache are played by
`Kdc`, which stands behind the module's few seams to the GSS library."""

import pytest
import xml_samples as x
from fake_kerberos import FakeKerberosAcceptor, FakeKerberosClient, spnego_acceptor
from fake_ssas import FakeSsas

from cubebridge.errors import CubeError
from cubebridge.ssas import kerberos, sealing
from cubebridge.ssas.session import Session

_BASE = -1765328384
C_PRINCIPAL_UNKNOWN = _BASE + 6
S_PRINCIPAL_UNKNOWN = _BASE + 7
PREAUTH_FAILED = _BASE + 24
SKEW = _BASE + 37
KDC_UNREACH = _BASE + 156
KRB5_OID = "1.2.840.113554.1.2.2"
NTLM_OID = "1.3.6.1.4.1.311.2.2.10"

FQDN = "bi.corp.example.com"
REALM = "CORP.EXAMPLE.COM"
SPN = f"MSOLAPSvc.3/{FQDN}@{REALM}"
PASSWORD = "correct horse battery"
CATALOGS = x.rowset([{"CATALOG_NAME": "Sales", "DESCRIPTION": ""}])


class KrbError(Exception):
    """Shaped like krb5.Krb5Error: a numeric code and the library's own words."""

    def __init__(self, code: int, message: str) -> None:
        super().__init__(message)
        self.code = code


class Cred(tuple):
    """A credential: (how it was got, whose), with the seconds it has left."""

    lifetime = 36000


class Kdc:
    def __init__(self) -> None:
        self.ticket: str | None = None  # principal of a ticket already on this Mac
        self.users = {f"dana@{REALM}": PASSWORD}
        self.services = {SPN}
        self.realms = {REALM}  # realms whose KDC answers
        self.no_trust = set()  # realms a service ticket fails in, in words nobody classifies
        self.skew = False
        self.kinits: list[str] = []
        self.asked: list[str] = []
        self.names = {"bi": FQDN, FQDN: FQDN}  # forward lookups (canonical names)
        self.reverse = {"10.1.2.3": FQDN}

    def existing_ticket(self):
        return (Cred(("ticket", self.ticket)), self.ticket) if self.ticket else None

    def password_credential(self, principal: str, password: str):
        self.kinits.append(principal)
        realm = principal.rpartition("@")[2]
        if realm not in self.realms:
            raise KrbError(KDC_UNREACH, f"unable to reach any KDC in realm {realm}")
        if self.skew:
            raise KrbError(SKEW, "Clock skew too great")
        if principal not in self.users:
            raise KrbError(C_PRINCIPAL_UNKNOWN, f"Client ({principal}) unknown")
        if self.users[principal] != password:
            raise KrbError(PREAUTH_FAILED, "Preauthentication failed")
        return Cred(("password", principal))

    def new_context(self, credential, target: str, channel_bindings):
        kdc = self

        class Context(FakeKerberosClient):
            def step(self, in_token=None, *, channel_bindings=None):
                if not self._sent:
                    kdc.asked.append(target)
                    realm = target.rpartition("@")[2]
                    if realm in kdc.no_trust:
                        raise Exception(f"Matching credential (krbtgt/{realm}@{REALM}) not found")
                    if realm not in kdc.realms:
                        raise KrbError(KDC_UNREACH, f"unable to reach any KDC in realm {realm}")
                    if target not in kdc.services:
                        raise KrbError(S_PRINCIPAL_UNKNOWN, f"Server ({target}) unknown")
                return super().step(in_token, channel_bindings=channel_bindings)

        context = Context(target)
        context.credential = credential
        context.bindings = channel_bindings
        return context


@pytest.fixture
def kdc(monkeypatch):
    kdc = Kdc()
    monkeypatch.setattr(kerberos, "_tickets", {})
    monkeypatch.setattr(kerberos, "_modules", lambda: (None, None))
    monkeypatch.setattr(kerberos, "existing_ticket", kdc.existing_ticket)
    monkeypatch.setattr(kerberos, "_password_credential", kdc.password_credential)
    monkeypatch.setattr(kerberos, "_new_context", kdc.new_context)
    monkeypatch.setattr(kerberos, "_canonical_name", lambda host: kdc.names.get(host, ""))
    monkeypatch.setattr(kerberos, "_reverse_name", lambda address: kdc.reverse.get(address, ""))
    return kdc


def start(username="CORP\\dana", password=PASSWORD, host="bi", **kwargs):
    kwargs.setdefault("service", "MSOLAPSvc.3")
    kwargs.setdefault("wrap", False)
    return kerberos.client_context(username, password, host, **kwargs)


def fails(**kwargs) -> CubeError:
    with pytest.raises(CubeError) as info:
        start(**kwargs)
    assert PASSWORD not in info.value.message and PASSWORD not in (info.value.detail or "")
    return info.value


# -- credentials ---------------------------------------------------------------------
def test_a_ticket_already_on_this_mac_is_used_and_the_password_is_not(kdc):
    kdc.ticket = f"dana@{REALM}"
    context = start()
    assert context.credential == ("ticket", kdc.ticket)
    assert kdc.kinits == []
    assert context.cube_choice == kerberos.Choice("ticket", REALM, SPN)


def test_without_a_ticket_one_is_got_in_the_servers_domain_with_the_password(kdc):
    context = start()
    assert kdc.kinits == [f"dana@{REALM}"]
    assert context.credential == ("password", f"dana@{REALM}")
    assert context.cube_choice.credential == "password"
    # The first token is ready before anything goes to the server.
    assert context.step().startswith(b"\x60FAKE-AP-REQ " + SPN.encode())


def test_a_ticket_got_with_the_password_is_kept_for_the_next_connection(kdc):
    start()
    start()
    assert kdc.kinits == [f"dana@{REALM}"]
    kdc.users[f"dana@{REALM}"] = "a new password"
    start(password="a new password")  # a changed password never gets the old ticket
    assert kdc.kinits == [f"dana@{REALM}", f"dana@{REALM}"]


def test_a_kept_ticket_that_is_running_out_is_replaced(kdc, monkeypatch):
    start()
    monkeypatch.setattr(Cred, "lifetime", 60)
    start()
    assert len(kdc.kinits) == 2


def test_a_ticket_for_someone_else_is_not_used(kdc):
    kdc.ticket = f"admin@{REALM}"
    start()
    assert kdc.kinits == [f"dana@{REALM}"]


def test_a_user_at_domain_name_names_the_realm_itself(kdc):
    kdc.users["dana@SALES.EXAMPLE.COM"] = PASSWORD
    kdc.realms.add("SALES.EXAMPLE.COM")
    kdc.services.add(f"MSOLAPSvc.3/{FQDN}@SALES.EXAMPLE.COM")
    start(username="dana@sales.example.com")
    assert kdc.kinits == ["dana@SALES.EXAMPLE.COM"]


def test_the_realm_above_the_servers_domain_is_tried_when_the_user_is_not_in_it(kdc):
    kdc.users = {f"dana@EXAMPLE.COM": PASSWORD}
    kdc.realms.add("EXAMPLE.COM")
    kdc.services.add(f"MSOLAPSvc.3/{FQDN}@EXAMPLE.COM")
    context = start()
    assert kdc.kinits == [f"dana@{REALM}", "dana@EXAMPLE.COM"]
    assert context.cube_choice.realm == "EXAMPLE.COM"


def test_a_wrong_password_is_tried_once_only(kdc):
    error = fails(password="wrong")
    assert kdc.kinits == [f"dana@{REALM}"]  # no second realm: that would be one more failed sign-in
    assert error.category == "auth" and "password" in error.message


def test_no_reachable_kerberos_server_says_so_and_suggests_ntlm(kdc):
    kdc.realms = set()
    error = fails()
    assert error.category == "auth-mechanism"
    assert "Kerberos server" in error.message and "VPN" in error.message and "NTLM" in error.message


def test_an_unknown_user_is_explained(kdc):
    kdc.users = {}
    error = fails()
    assert error.category == "auth" and "user name" in error.message


def test_a_clock_too_far_off_is_explained(kdc):
    kdc.skew = True
    assert "clock" in fails().message


def test_without_any_domain_to_go_on_the_user_is_asked_for_one(kdc):
    kdc.names = {}
    error = fails(username="dana", host="bi")
    assert "name@company-domain" in error.message
    assert kdc.kinits == []


# -- the service's name --------------------------------------------------------------
def test_the_full_dns_name_is_used_for_the_service(kdc):
    start(host="bi")
    assert kdc.asked == [SPN]


def test_a_named_instance_is_part_of_the_service_name(kdc):
    named = f"MSOLAPSvc.3/{FQDN}:TABULAR@{REALM}"
    kdc.services = {named}
    context = start(instance="TABULAR")
    assert context.cube_choice.service == named


def test_an_ip_address_is_looked_up_in_reverse(kdc):
    start(host="10.1.2.3")
    assert kdc.asked == [SPN]


def test_an_ip_address_without_a_name_is_explained(kdc):
    error = fails(host="10.9.9.9")
    assert error.category == "auth-mechanism" and "IP address" in error.message and "NTLM" in error.message


def test_the_name_as_typed_is_tried_after_the_full_name(kdc):
    kdc.services = {f"MSOLAPSvc.3/bi@{REALM}"}
    context = start(host="bi")
    assert kdc.asked == [SPN, f"MSOLAPSvc.3/bi@{REALM}"]
    assert context.cube_choice.service == f"MSOLAPSvc.3/bi@{REALM}"


def test_the_users_realm_is_tried_when_the_servers_domain_has_none(kdc):
    kdc.names = {"bi": "bi.dmz.example.net"}
    kdc.services = {f"MSOLAPSvc.3/bi.dmz.example.net@{REALM}"}
    context = start(username=f"dana@{REALM.lower()}")
    assert kdc.asked[0] == "MSOLAPSvc.3/bi.dmz.example.net@DMZ.EXAMPLE.NET"
    assert context.cube_choice.service == f"MSOLAPSvc.3/bi.dmz.example.net@{REALM}"


def test_a_failure_in_unfamiliar_words_for_one_name_moves_on_to_the_next(kdc):
    kdc.names = {"bi": "bi.dmz.example.net"}
    kdc.no_trust = {"DMZ.EXAMPLE.NET"}
    kdc.services = {f"MSOLAPSvc.3/bi.dmz.example.net@{REALM}"}
    context = start(username=f"dana@{REALM.lower()}")
    assert kdc.asked == ["MSOLAPSvc.3/bi.dmz.example.net@DMZ.EXAMPLE.NET", f"MSOLAPSvc.3/bi.dmz.example.net@{REALM}"]
    assert context.cube_choice.realm == REALM


def test_a_clock_problem_at_the_service_ticket_stops_at_once(kdc, monkeypatch):
    asked = []

    class Skewed:
        def __init__(self, credential, target, bindings):
            self.target = target

        def step(self, in_token=None, **_):
            asked.append(self.target)
            raise KrbError(SKEW, "Clock skew too great")

    monkeypatch.setattr(kerberos, "_new_context", Skewed)
    kdc.names = {"bi": FQDN, FQDN: FQDN}
    error = fails(host="bi")
    assert "clock" in error.message
    assert len(asked) == 1  # not tried again under the server's other names


def test_a_server_kerberos_does_not_know_suggests_ntlm_and_named_instances(kdc):
    kdc.services = set()
    error = fails()
    assert error.category == "auth-mechanism"
    assert "NTLM" in error.message and "name\\INSTANCE" in error.message


def test_http_asks_for_the_web_servers_name(kdc):
    kdc.services = {f"HTTP/{FQDN}@{REALM}"}
    context = start(service="HTTP", channel_bindings="bindings")
    assert context.cube_choice.service == f"HTTP/{FQDN}@{REALM}"
    assert context.bindings == "bindings"


# -- SPNEGO --------------------------------------------------------------------------
def test_spnego_offers_kerberos_and_nothing_else(kdc):
    from spnego._spnego import NegTokenInit, unpack_token

    token = unpack_token(start(wrap=True).step())
    assert isinstance(token, NegTokenInit)
    assert token.mech_types == [KRB5_OID] and NTLM_OID not in token.mech_types
    assert token.mech_token.startswith(b"\x60FAKE-AP-REQ ")


# -- errors --------------------------------------------------------------------------
@pytest.mark.parametrize(
    "error, reason",
    [
        (KrbError(PREAUTH_FAILED, "x"), "wrong-password"),
        (KrbError(KDC_UNREACH, "x"), "no-kdc"),
        (KrbError(S_PRINCIPAL_UNKNOWN, "x"), "no-service"),
        (Exception("Server not found in Kerberos database"), "no-service"),
        (Exception("unable to reach any KDC in realm CORP"), "no-kdc"),
        (Exception("Client (dana@CORP) unknown"), "no-user"),
        (Exception("Clock skew too great"), "clock"),
        (Exception("something else"), "other"),
    ],
)
def test_errors_are_classified_by_code_or_wording(error, reason):
    assert kerberos.classify(error) == reason


def test_an_unsigned_gss_minor_code_and_a_wrapped_error_are_understood():
    class GssError(Exception):
        min_code = PREAUTH_FAILED + 2**32

    class Wrapper(Exception):
        base_error = GssError("Unspecified GSS failure")

    assert kerberos.classify(Wrapper("SpnegoError")) == "wrong-password"


def test_without_the_gss_packages_kerberos_says_it_is_not_installed(monkeypatch):
    from spnego import _gss

    monkeypatch.setattr(_gss, "HAS_GSSAPI", False)
    with pytest.raises(CubeError) as info:
        kerberos.client_context("CORP\\dana", PASSWORD, "bi", service="MSOLAPSvc.3")
    assert info.value.category == "auth-mechanism"
    assert "not installed" in info.value.message and "NTLM" in info.value.message


def test_rc4_padding_is_refused_with_a_plain_sentence():
    class PaddingContext:
        def wrap_winrm(self, data):
            return b"H" * 32, data + b"\x01", 1

    with pytest.raises(sealing.ProtocolError) as info:
        sealing.seal_frame(PaddingContext(), b"abc")
    assert info.value.category == "auth-mechanism" and "NTLM" in info.value.message


# -- over the TCP binding, against the stand-in server --------------------------------
def test_a_kerberos_session_signs_in_and_seals_with_kerberos_sized_tokens(kdc):
    with FakeSsas({"DBSCHEMA_CATALOGS": CATALOGS}, acceptor=spnego_acceptor) as server:
        session = Session(
            "127.0.0.1",
            server.port,
            "CORP\\dana",
            PASSWORD,
            sign_in="kerberos",
            connect_timeout=5,
            io_timeout=5,
            context_factory=lambda user, password, _host: kerberos.client_context(
                user, password, "bi", service="MSOLAPSvc.3"
            ),
        )
        session.open()
        assert [r["CATALOG_NAME"] for r in session.discover("DBSCHEMA_CATALOGS")] == ["Sales"]
        assert [r["CATALOG_NAME"] for r in session.discover("DBSCHEMA_CATALOGS")] == ["Sales"]
        session.close()
    assert server.contexts[0].fake.target == SPN
    assert server.client_token_sizes and set(server.client_token_sizes) == {60}
    assert "Kerberos" in session.sign_in_note and "got with the password" in session.sign_in_note


def test_the_default_kerberos_factory_names_the_instance(monkeypatch):
    seen = {}

    def fake_client_context(username, password, host, **kwargs):
        seen.update(kwargs, host=host)
        return FakeKerberosClient(f"{kwargs['service']}/bi:{kwargs['instance']}@X")

    monkeypatch.setattr(kerberos, "client_context", fake_client_context)
    with FakeSsas({"DBSCHEMA_CATALOGS": CATALOGS}, acceptor=FakeKerberosAcceptor) as server:
        session = Session("127.0.0.1", server.port, "CORP\\dana", PASSWORD, sign_in="kerberos", instance="TAB", connect_timeout=5)
        session.open()
        assert len(session.discover("DBSCHEMA_CATALOGS")) == 1
        session.close()
    assert seen == {"service": "MSOLAPSvc.3", "instance": "TAB", "host": "127.0.0.1"}


def test_a_server_that_refuses_the_kerberos_ticket_is_explained(kdc):
    class Refusing(FakeKerberosAcceptor):
        def step(self, in_token=None, *, channel_bindings=None):
            raise ValueError("KRB_AP_ERR_MODIFIED")

    with FakeSsas(acceptor=Refusing) as server:
        session = Session(
            "127.0.0.1", server.port, "CORP\\dana", PASSWORD, sign_in="kerberos", connect_timeout=5, io_timeout=5,
            context_factory=lambda u, p, _h: kerberos.client_context(u, p, "bi", service="MSOLAPSvc.3", wrap=False),
        )
        with pytest.raises(CubeError) as info:
            session.open()
    assert info.value.category == "auth" and "Kerberos" in info.value.message and "NTLM" in info.value.message


# -- the one piece that touches the real GSS library ---------------------------------
def test_the_gss_context_is_made_for_the_exact_service_principal(monkeypatch):
    gssapi = pytest.importorskip("gssapi")
    from spnego import _gss

    if not (_gss.HAS_GSSAPI and _gss.HAS_IOV):
        pytest.skip("gssapi without IOV")
    made = {}

    class Recorder:
        complete = False
        actual_flags = 0

        def __init__(self, **kwargs):
            made.update(kwargs)

        def step(self, in_token):
            return b"first token"

    monkeypatch.setattr(gssapi, "SecurityContext", Recorder)
    credential = object()
    target = f"MSOLAPSvc.3/{FQDN}:TABULAR@{REALM}"
    context = kerberos._KerberosContext(credential, target)
    assert context.step() == b"first token"
    assert str(made["name"]) == target
    assert made["creds"] is credential and made["usage"] == "initiate"
    assert made["mech"].dotted_form == KRB5_OID
