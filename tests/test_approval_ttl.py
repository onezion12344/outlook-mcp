from __future__ import annotations

import json
import os
import time

from outlook_web_mcp import approval


def _nonce(directory: str, digest: str):
    return approval.approval_path(directory, digest)


def test_a_fresh_nonce_is_consumed_exactly_once(tmp_path) -> None:
    directory = str(tmp_path / "approvals")
    approval.approve(directory, "abc123")
    assert approval.consume(directory, "abc123") is True
    assert approval.consume(directory, "abc123") is False
    assert not _nonce(directory, "abc123").exists()


def test_an_expired_nonce_is_refused_and_removed(tmp_path) -> None:
    directory = str(tmp_path / "approvals")
    path = approval.approve(directory, "abc123")
    path.write_text(json.dumps({"at": int(time.time()) - 7200, "ttl": 1800}), encoding="utf-8")

    assert approval.consume(directory, "abc123") is False
    assert not path.exists()


def test_a_legacy_plain_epoch_file_is_still_understood(tmp_path) -> None:
    directory = str(tmp_path / "approvals")
    path = _nonce(directory, "old")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(f"{int(time.time())}\n", encoding="utf-8")
    assert approval.consume(directory, "old") is True


def test_corrupt_nonce_is_refused_and_cleaned_up(tmp_path) -> None:
    directory = str(tmp_path / "approvals")
    path = _nonce(directory, "junk")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("not a timestamp", encoding="utf-8")
    assert approval.consume(directory, "junk") is False
    assert not path.exists()


def test_purge_drops_stale_nonces_and_keeps_fresh_ones(tmp_path) -> None:
    directory = str(tmp_path / "approvals")
    fresh = approval.approve(directory, "fresh")
    stale = _nonce(directory, "stale")
    stale.write_text(json.dumps({"at": int(time.time()) - 9999}), encoding="utf-8")

    removed = approval.purge_expired(directory)
    assert removed == 1
    assert fresh.exists()
    assert not stale.exists()


def test_nonce_file_is_owner_only(tmp_path) -> None:
    directory = str(tmp_path / "approvals")
    path = approval.approve(directory, "perm")
    if os.name == "posix":
        assert oct(os.stat(path).st_mode & 0o777) == "0o600"
