from __future__ import annotations

import json

import keyring
import pytest

from outlook_web_mcp.credentials import (
    SECRET_CHUNK_CHARS,
    SERVICE_NAME,
    KeyringCredentialStore,
)
from outlook_web_mcp.errors import AuthenticationError
from outlook_web_mcp.models import CredentialBundle

PROFILE = "cred-test"
METADATA_USER = f"{PROFILE}:metadata"


class FakeBackend:
    def __init__(self, priority: float = 5.0) -> None:
        self.priority = priority


@pytest.fixture
def vault(monkeypatch):
    store: dict[tuple[str, str], str] = {}

    def get_password(service: str, username: str) -> str | None:
        return store.get((service, username))

    def set_password(service: str, username: str, value: str) -> None:
        store[(service, username)] = value

    def delete_password(service: str, username: str) -> None:
        if (service, username) not in store:
            raise keyring.errors.PasswordDeleteError("missing")
        del store[(service, username)]

    monkeypatch.setattr(keyring, "get_password", get_password)
    monkeypatch.setattr(keyring, "set_password", set_password)
    monkeypatch.setattr(keyring, "delete_password", delete_password)
    monkeypatch.setattr(keyring, "get_keyring", lambda: FakeBackend(5.0))
    return store


def _bundle(token: str = "r" * (SECRET_CHUNK_CHARS * 2 + 17)) -> CredentialBundle:
    return CredentialBundle(
        refresh_token=token,
        client_id="cid",
        tenant_id="tid",
        puid="puid",
        display_name="Owner",
    )


def test_round_trip_splits_a_long_token(vault) -> None:
    store = KeyringCredentialStore(PROFILE)
    bundle = _bundle()
    store.save(bundle)

    chunks = [value for (service, user), value in vault.items() if ":refresh:" in user]
    assert len(chunks) == 3
    assert all(len(chunk) <= SECRET_CHUNK_CHARS for chunk in chunks)

    loaded = store.load()
    assert loaded is not None
    assert loaded.refresh_token == bundle.refresh_token
    assert loaded.tenant_id == "tid"
    assert loaded.puid == "puid"
    assert loaded.display_name == "Owner"


def test_metadata_entry_never_contains_the_token(vault) -> None:
    store = KeyringCredentialStore(PROFILE)
    token = "secret-token-value-" * 80
    store.save(_bundle(token))

    metadata = vault[(SERVICE_NAME, METADATA_USER)]
    assert "secret-token-value" not in metadata
    parsed = json.loads(metadata)
    assert "refresh_token" not in parsed
    assert parsed["schema_version"] == 1
    assert parsed["refresh_parts"] == -(-len(token) // SECRET_CHUNK_CHARS)


def test_no_single_vault_entry_holds_the_whole_token_when_long(vault) -> None:
    store = KeyringCredentialStore(PROFILE)
    bundle = _bundle()
    store.save(bundle)
    assert bundle.refresh_token not in vault.values()


def test_save_rotates_generation_and_drops_the_old_chunks(vault) -> None:
    store = KeyringCredentialStore(PROFILE)
    store.save(_bundle("a" * (SECRET_CHUNK_CHARS + 1)))
    first_users = {user for (_, user) in vault if ":refresh:" in user}

    store.save(_bundle("b" * (SECRET_CHUNK_CHARS + 1)))
    remaining = {user for (_, user) in vault if ":refresh:" in user}

    assert not (first_users & remaining)
    assert store.load().refresh_token == "b" * (SECRET_CHUNK_CHARS + 1)


def test_delete_removes_everything_for_the_profile(vault) -> None:
    store = KeyringCredentialStore(PROFILE)
    store.save(_bundle())
    store.delete()
    assert store.load() is None
    assert not [user for (_, user) in vault if user.startswith(PROFILE)]


def test_insecure_backend_refuses_to_store_or_read(vault, monkeypatch) -> None:
    monkeypatch.setattr(keyring, "get_keyring", lambda: FakeBackend(0.0))
    store = KeyringCredentialStore(PROFILE)

    with pytest.raises(AuthenticationError) as raised:
        store.save(_bundle())
    assert raised.value.code == "keyring_unavailable"
    assert vault == {}

    with pytest.raises(AuthenticationError) as raised:
        store.load()
    assert raised.value.code == "keyring_unavailable"


def test_missing_chunk_is_reported_not_silently_truncated(vault) -> None:
    store = KeyringCredentialStore(PROFILE)
    store.save(_bundle())
    victim = max(user for (_, user) in vault if ":refresh:" in user)
    del vault[(SERVICE_NAME, victim)]

    with pytest.raises(AuthenticationError) as raised:
        store.load()
    assert raised.value.code == "credentials_incomplete"


def test_corrupt_metadata_is_reported(vault) -> None:
    store = KeyringCredentialStore(PROFILE)
    vault[(SERVICE_NAME, METADATA_USER)] = "{not json"
    with pytest.raises(AuthenticationError) as raised:
        store.load()
    assert raised.value.code == "credentials_invalid"


def test_unsupported_schema_version_is_reported(vault) -> None:
    store = KeyringCredentialStore(PROFILE)
    vault[(SERVICE_NAME, METADATA_USER)] = json.dumps(
        {"schema_version": 2, "generation": "g", "refresh_parts": 1}
    )
    with pytest.raises(AuthenticationError) as raised:
        store.load()
    assert raised.value.code == "credentials_invalid"


def test_load_returns_none_when_nothing_was_saved(vault) -> None:
    assert KeyringCredentialStore("never-used").load() is None
