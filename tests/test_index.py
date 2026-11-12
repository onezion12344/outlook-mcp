from __future__ import annotations

import os
import sqlite3
from pathlib import Path

import pytest

from outlook_web_mcp.index import (
    match_expression,
    open_index,
    prune_folder,
    search,
    status,
    upsert_rows,
)


def _fts5_available() -> bool:
    try:
        probe = sqlite3.connect(":memory:")
        probe.execute("CREATE VIRTUAL TABLE t USING fts5(x)")
        return True
    except sqlite3.OperationalError:
        return False


pytestmark = pytest.mark.skipif(not _fts5_available(), reason="sqlite build lacks FTS5")


def _row(
    email_id: str,
    *,
    subject: str = "Quarterly update",
    sender: str = "registry@example.com",
    received: str = "2026-11-05T09:00:00Z",
    preview: str | None = None,
    message_class: str = "IPM.Note",
) -> dict[str, object]:
    return {
        "emailId": email_id,
        "subject": subject,
        "senderName": "University Registry",
        "senderAddress": sender,
        "receivedAt": received,
        "isUnread": True,
        "hasAttachments": False,
        "messageClass": message_class,
        "internetMessageId": f"<{email_id}@tenant>",
        "references": "",
        "bodyPreview": preview,
    }


@pytest.fixture
def conn(tmp_path):
    connection = open_index(str(tmp_path / "idx" / "index.db"))
    yield connection
    connection.close()


def test_upsert_is_idempotent_and_updates_fields(conn) -> None:
    assert upsert_rows(conn, [_row("m1")], "inbox") == 1
    assert upsert_rows(conn, [_row("m1", subject="Renamed")], "inbox") == 1
    assert status(conn)["rows"] == 1
    assert search(conn, "renamed")[0]["subject"] == "Renamed"


def test_search_matches_body_preview_not_only_subject(conn) -> None:
    upsert_rows(
        conn,
        [_row("m1", subject="Weekly digest", preview="the scholarship deadline moved")],
        "inbox",
    )
    hits = search(conn, "scholarship")
    assert [hit["emailId"] for hit in hits] == ["m1"]


def test_search_survives_fts_metacharacters(conn) -> None:
    upsert_rows(conn, [_row("m1", subject='He said "hi"')], "inbox")
    for query in ['"hi"', "NEAR(", "a OR", '""', "-*"]:
        assert isinstance(search(conn, query), list)


def test_folder_and_since_filters(conn) -> None:
    upsert_rows(conn, [_row("m1", received="2026-01-01T00:00:00Z")], "inbox")
    upsert_rows(conn, [_row("m2", received="2026-06-01T00:00:00Z")], "archive")
    assert {hit["emailId"] for hit in search(conn, "", folder="archive")} == {"m2"}
    assert {hit["emailId"] for hit in search(conn, "", since="2026-05-01")} == {"m2"}
    assert len(search(conn, "")) == 2


def test_status_reports_shape(conn) -> None:
    upsert_rows(conn, [_row("m1")], "inbox")
    upsert_rows(conn, [_row("m2")], "sent")
    info = status(conn)
    assert info["rows"] == 2
    assert info["folders"] == 2
    assert info["oldest"] and info["newest"] and info["lastSync"]


def test_message_class_is_queryable_directly(conn) -> None:
    upsert_rows(conn, [_row("m1", message_class="REPORT.IPM.Note.NDR")], "inbox")
    found = conn.execute(
        "SELECT email_id FROM items WHERE message_class LIKE 'REPORT.IPM.Note.NDR%'"
    ).fetchall()
    assert [row["email_id"] for row in found] == ["m1"]


def test_limit_is_clamped(conn) -> None:
    upsert_rows(conn, [_row(f"m{i}") for i in range(5)], "inbox")
    assert len(search(conn, "", limit=500)) == 5
    assert len(search(conn, "", limit=0)) == 1


def test_database_file_is_owner_only(tmp_path) -> None:
    path = tmp_path / "idx" / "index.db"
    connection = open_index(str(path))
    connection.close()
    if os.name == "posix":
        assert oct(os.stat(path).st_mode & 0o777) == "0o600"


def test_match_expression_quotes_every_token() -> None:
    assert match_expression('foo "bar"') == '"foo" """bar"""'
    assert match_expression("   ") == ""


def test_timestamps_are_stored_in_one_comparable_form(conn) -> None:
    upsert_rows(conn, [_row("m1", received="2026-11-06T09:00:00+08:00")], "inbox")
    stored = conn.execute("SELECT received_at FROM items WHERE email_id='m1'").fetchone()[0]
    assert stored == "2026-11-06T01:00:00Z"


def test_since_compares_instants_not_strings(conn) -> None:
    # 2026-11-05T23:30-05:00 == 2026-11-06T04:30Z, i.e. clearly after the boundary,
    # yet a raw string compare would put it before "2026-11-06T00:00:00Z".
    upsert_rows(conn, [_row("m1", received="2026-11-05T23:30:00-05:00")], "inbox")
    assert [hit["emailId"] for hit in search(conn, "", since="2026-11-06T00:00:00Z")] == ["m1"]


def test_prune_respects_the_normalised_boundary(conn) -> None:
    upsert_rows(conn, [_row("m1", received="2026-11-05T23:30:00-05:00")], "inbox")
    # Inside the window once normalised, so an unseen-but-present row must survive.
    removed = prune_folder(conn, "inbox", since="2026-11-06T00:00:00Z", keep_ids={"m1"})
    assert removed == 0
    assert search(conn, "")[0]["emailId"] == "m1"


def test_opening_the_index_migrates_legacy_offset_timestamps(tmp_path) -> None:
    path = str(tmp_path / "idx" / "legacy.db")
    connection = open_index(path)
    connection.execute(
        "INSERT INTO items (email_id, folder, received_at, indexed_at) VALUES (?,?,?,?)",
        ("legacy-1", "inbox", "2026-11-06T09:00:00+08:00", "2026-11-06T02:00:00Z"),
    )
    connection.commit()
    connection.close()

    reopened = open_index(path)
    stored = reopened.execute("SELECT received_at FROM items WHERE email_id='legacy-1'").fetchone()[
        0
    ]
    assert stored == "2026-11-06T01:00:00Z"
    reopened.close()


def test_old_index_is_upgraded_in_place_without_losing_rows(tmp_path) -> None:
    import sqlite3

    from outlook_web_mcp import index as index_store

    path = str(tmp_path / "old.db")
    legacy = sqlite3.connect(path)
    legacy.executescript(
        """
        CREATE TABLE items (
          email_id TEXT PRIMARY KEY, folder TEXT NOT NULL,
          subject TEXT NOT NULL DEFAULT '', sender_name TEXT NOT NULL DEFAULT '',
          sender_address TEXT NOT NULL DEFAULT '', received_at TEXT NOT NULL DEFAULT '',
          is_read INTEGER NOT NULL DEFAULT 0, has_attachments INTEGER NOT NULL DEFAULT 0,
          message_class TEXT NOT NULL DEFAULT '', internet_message_id TEXT NOT NULL DEFAULT '',
          refs TEXT NOT NULL DEFAULT '', body_preview TEXT NOT NULL DEFAULT '',
          indexed_at TEXT NOT NULL
        );
        CREATE VIRTUAL TABLE items_fts USING fts5(
          email_id UNINDEXED, subject, sender_name, sender_address, body_preview
        );
        INSERT INTO items (email_id, folder, subject, received_at, indexed_at)
          VALUES ('old-1', 'inbox', 'Legacy subject',
                  '2026-01-01T00:00:00Z', '2026-01-01T00:00:00Z');
        INSERT INTO items_fts VALUES ('old-1', 'Legacy subject', '', '', '');
        """
    )
    legacy.commit()
    legacy.close()

    connection = index_store.open_index(path)
    try:
        # Row survived, new columns exist, and the rebuilt FTS still answers.
        assert connection.execute("SELECT COUNT(*) FROM items").fetchone()[0] == 1
        row = connection.execute("SELECT body_text, body_fetched FROM items").fetchone()
        assert (row["body_text"], row["body_fetched"]) == ("", 0)
        assert index_store.search(connection, "Legacy", limit=5)
    finally:
        connection.close()


def test_full_body_is_stored_and_searchable(tmp_path) -> None:
    from outlook_web_mcp import index as index_store

    connection = index_store.open_index(str(tmp_path / "bodies.db"))
    try:
        index_store.upsert_rows(
            connection,
            [{"emailId": "m1", "subject": "Short", "bodyPreview": "first 255 chars"}],
            "inbox",
        )
        # A word that exists only deep in the body, past the preview.
        index_store.store_body(
            connection,
            "m1",
            body_text="preamble " + "filler " * 300 + "kubernetes",
            web_link="https://outlook.cloud.microsoft/owa/?ItemID=m1",
            attachments=[
                {
                    "name": "report.pdf",
                    "contentType": "application/pdf",
                    "size": 1234,
                    "summary": "application/pdf, 1234 bytes",
                }
            ],
        )
        hits = index_store.search(connection, "kubernetes", limit=5)
        assert [hit["emailId"] for hit in hits] == ["m1"]
        assert index_store.body_progress(connection) == {
            "total": 1,
            "fetched": 1,
            "chars": len("preamble " + "filler " * 300 + "kubernetes"),
        }
        attachment = connection.execute(
            "SELECT name, size FROM attachments WHERE email_id = 'm1'"
        ).fetchone()
        assert (attachment["name"], attachment["size"]) == ("report.pdf", 1234)
    finally:
        connection.close()


def test_durable_state_is_not_under_caches(tmp_path, monkeypatch) -> None:
    from outlook_web_mcp.config import Settings

    monkeypatch.setattr(Settings, "data_dir", property(lambda self: str(tmp_path / "data")))
    settings = Settings(credential_profile="path-test")
    assert "Caches" not in settings.index_path
    assert Path(settings.index_path).name == "index.db"
    assert "Caches" not in settings.pixel_dir
    assert "Caches" not in settings.sent_links_dir


def test_ocr_writeback_marks_provenance_and_is_searchable(tmp_path) -> None:
    from outlook_web_mcp import index as index_store

    connection = index_store.open_index(str(tmp_path / "ocr.db"))
    try:
        index_store.upsert_rows(connection, [{"emailId": "js-1", "subject": "Verify"}], "inbox")
        index_store.store_body(connection, "js-1", body_text="Your code is 424242", source="ocr")
        row = connection.execute(
            "SELECT body_source, body_fetched FROM items WHERE email_id='js-1'"
        ).fetchone()
        assert (row["body_source"], row["body_fetched"]) == ("ocr", 1)
        assert [h["emailId"] for h in index_store.search(connection, "424242")] == ["js-1"]
    finally:
        connection.close()
