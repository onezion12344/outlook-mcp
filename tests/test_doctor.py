from __future__ import annotations

import json

import keyring
import pytest

from outlook_web_mcp import doctor
from outlook_web_mcp.config import Settings
from outlook_web_mcp.credentials import KeyringCredentialStore
from outlook_web_mcp.index import open_index
from outlook_web_mcp.models import CredentialBundle


class _Backend:
    def __init__(self, priority: float) -> None:
        self.priority = priority


@pytest.fixture
def settings(tmp_path, monkeypatch):
    monkeypatch.setattr(Settings, "index_path", property(lambda self: str(tmp_path / "index.db")))
    monkeypatch.setattr(
        Settings, "approval_dir", property(lambda self: str(tmp_path / "approvals"))
    )
    monkeypatch.setattr(Settings, "pixel_dir", property(lambda self: str(tmp_path / "pixels")))
    monkeypatch.setattr(keyring, "get_keyring", lambda: _Backend(5.0))
    monkeypatch.setattr(KeyringCredentialStore, "load", lambda self: None)
    instance = Settings(credential_profile="doctor-test")
    object.__setattr__(instance, "pixel_base", None)
    object.__setattr__(instance, "pixel_sync_token", None)
    return instance


def _by_name(checks):
    return {check.name: check for check in checks}


def test_unconfigured_machine_reports_warnings_not_failures(settings) -> None:
    report = doctor.summarise(doctor.run_checks(settings))
    checks = _by_name(doctor.run_checks(settings))
    assert report["ok"] is True
    assert checks["keyring"].status == "ok"
    assert checks["credentials"].status == "warn"
    assert checks["index"].status == "warn"
    assert checks["approvals"].status == "ok"
    assert checks["pixels"].status == "ok"


def test_insecure_keyring_backend_fails(settings, monkeypatch) -> None:
    monkeypatch.setattr(keyring, "get_keyring", lambda: _Backend(0.0))
    report = doctor.summarise(doctor.run_checks(settings))
    assert report["ok"] is False
    assert report["counts"]["fail"] == 1
    assert _by_name(doctor.run_checks(settings))["keyring"].status == "fail"


def test_credentials_are_reported_without_leaking_them(settings, monkeypatch) -> None:
    secret = "SECRET-REFRESH-TOKEN"
    monkeypatch.setattr(
        KeyringCredentialStore,
        "load",
        lambda self: CredentialBundle(secret, "cid", "tid", display_name="Owner"),
    )
    report = doctor.summarise(doctor.run_checks(settings))
    blob = json.dumps(report)
    assert secret not in blob
    assert "Owner" in blob
    assert _by_name(doctor.run_checks(settings))["credentials"].status == "ok"


def test_stale_index_is_a_warning(settings) -> None:
    connection = open_index(settings.index_path)
    connection.execute(
        "INSERT INTO items (email_id, folder, received_at, indexed_at) VALUES (?,?,?,?)",
        ("m1", "inbox", "2026-10-01T00:00:00Z", "2026-10-01T00:00:00Z"),
    )
    connection.commit()
    connection.close()

    checks = _by_name(doctor.run_checks(settings))
    assert checks["index"].status == "warn"
    assert "rows" in checks["index"].detail or "1" in checks["index"].detail


def test_pending_and_expired_approvals_are_counted(settings) -> None:
    import time

    from outlook_web_mcp import approval

    approval.approve(settings.approval_dir, "fresh")
    expired = approval.approval_path(settings.approval_dir, "old")
    expired.write_text(json.dumps({"at": int(time.time()) - 9999}), encoding="utf-8")

    checks = _by_name(doctor.run_checks(settings))
    assert checks["approvals"].status == "warn"
    assert "1 live" in checks["approvals"].detail
    assert "1 expired" in checks["approvals"].detail
