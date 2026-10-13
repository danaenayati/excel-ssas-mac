import pytest

USER = "CORP\\dana"
PASSWORD = "correct horse battery"


@pytest.fixture
def ntlm_users(tmp_path, monkeypatch):
    """Accounts the fake server's NTLM acceptor knows about."""
    users = tmp_path / "ntlm-users"
    users.write_text(f"CORP:dana:{PASSWORD}\n")
    monkeypatch.setenv("NTLM_USER_FILE", str(users))
    return USER, PASSWORD
