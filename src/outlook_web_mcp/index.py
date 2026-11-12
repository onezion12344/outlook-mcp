from __future__ import annotations

import contextlib
import os
import sqlite3
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

SCHEMA = """
CREATE TABLE IF NOT EXISTS items (
  email_id TEXT PRIMARY KEY,
  folder TEXT NOT NULL,
  subject TEXT NOT NULL DEFAULT '',
  sender_name TEXT NOT NULL DEFAULT '',
  sender_address TEXT NOT NULL DEFAULT '',
  received_at TEXT NOT NULL DEFAULT '',
  is_read INTEGER NOT NULL DEFAULT 0,
  has_attachments INTEGER NOT NULL DEFAULT 0,
  message_class TEXT NOT NULL DEFAULT '',
  internet_message_id TEXT NOT NULL DEFAULT '',
  refs TEXT NOT NULL DEFAULT '',
  body_preview TEXT NOT NULL DEFAULT '',
  body_text TEXT NOT NULL DEFAULT '',
  body_fetched INTEGER NOT NULL DEFAULT 0,
  body_source TEXT NOT NULL DEFAULT '',
  web_link TEXT NOT NULL DEFAULT '',
  indexed_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS items_folder_received ON items(folder, received_at);
CREATE INDEX IF NOT EXISTS items_message_class ON items(message_class);
CREATE INDEX IF NOT EXISTS items_body_fetched ON items(body_fetched);
CREATE TABLE IF NOT EXISTS attachments (
  email_id TEXT NOT NULL,
  name TEXT NOT NULL DEFAULT '',
  content_type TEXT NOT NULL DEFAULT '',
  size INTEGER NOT NULL DEFAULT 0,
  summary TEXT NOT NULL DEFAULT '',
  attachment_id TEXT NOT NULL DEFAULT '',
  text_body TEXT NOT NULL DEFAULT '',
  text_method TEXT NOT NULL DEFAULT '',
  text_fetched INTEGER NOT NULL DEFAULT 0,
  PRIMARY KEY (email_id, name)
);
CREATE INDEX IF NOT EXISTS attachments_email ON attachments(email_id);
CREATE INDEX IF NOT EXISTS attachments_pending ON attachments(text_fetched, size);
CREATE VIRTUAL TABLE IF NOT EXISTS attachments_fts USING fts5(
  email_id UNINDEXED, name, text_body
);
CREATE VIRTUAL TABLE IF NOT EXISTS items_fts USING fts5(
  email_id UNINDEXED, subject, sender_name, sender_address, body_preview, body_text
);
"""

# Columns added after the first release. Kept as (name, DDL) so the upgrade path is
# explicit rather than "CREATE TABLE IF NOT EXISTS silently did nothing".
_MIGRATIONS: tuple[tuple[str, str], ...] = (
    ("body_text", "ALTER TABLE items ADD COLUMN body_text TEXT NOT NULL DEFAULT ''"),
    ("body_fetched", "ALTER TABLE items ADD COLUMN body_fetched INTEGER NOT NULL DEFAULT 0"),
    ("web_link", "ALTER TABLE items ADD COLUMN web_link TEXT NOT NULL DEFAULT ''"),
    ("body_source", "ALTER TABLE items ADD COLUMN body_source TEXT NOT NULL DEFAULT ''"),
)

_ATTACHMENT_MIGRATIONS: tuple[tuple[str, str], ...] = (
    ("attachment_id", "ALTER TABLE attachments ADD COLUMN attachment_id TEXT NOT NULL DEFAULT ''"),
    ("text_body", "ALTER TABLE attachments ADD COLUMN text_body TEXT NOT NULL DEFAULT ''"),
    ("text_method", "ALTER TABLE attachments ADD COLUMN text_method TEXT NOT NULL DEFAULT ''"),
    ("text_fetched", "ALTER TABLE attachments ADD COLUMN text_fetched INTEGER NOT NULL DEFAULT 0"),
)


def normalize_timestamp(value: str) -> str:
    """Store one comparable form (UTC, seconds, `Z`).

    The server reports timestamps with tenant offsets (e.g. `+08:00`). Comparing those
    as strings against a UTC `since` boundary is wrong across offsets, and that boundary
    drives both `search(since=...)` and the prune window.
    """
    text = (value or "").strip()
    if not text:
        return ""
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        return text
    if parsed.tzinfo is None:
        return text
    return parsed.astimezone(UTC).isoformat(timespec="seconds").replace("+00:00", "Z")


def open_index(path: str) -> sqlite3.Connection:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    connection = sqlite3.connect(str(target))
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA journal_mode=WAL")
    # Upgrade columns first: SCHEMA indexes a column (_body_fetched) that an old
    # database does not have yet, so it must run after the ALTERs.
    _migrate(connection)
    connection.executescript(SCHEMA)
    _rebuild_fts_if_needed(connection)
    _normalise_existing_timestamps(connection)
    with contextlib.suppress(OSError):
        os.chmod(target, 0o600)
    return connection


def _rebuild_fts_if_needed(connection: sqlite3.Connection) -> None:
    """The first FTS table had no body_text column; rebuild once and repopulate."""
    fts_sql = connection.execute(
        "SELECT sql FROM sqlite_master WHERE type='table' AND name='items_fts'"
    ).fetchone()
    if not fts_sql or "body_text" in str(fts_sql["sql"] or ""):
        return
    connection.execute("DROP TABLE items_fts")
    connection.executescript(SCHEMA)
    connection.execute(
        "INSERT INTO items_fts (email_id, subject, sender_name, sender_address,"
        " body_preview, body_text)"
        " SELECT email_id, subject, sender_name, sender_address, body_preview, body_text"
        " FROM items"
    )
    connection.commit()


def _migrate(connection: sqlite3.Connection) -> None:
    """Upgrade an index written by an older version in place, keeping every row."""
    exists = connection.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name='items'"
    ).fetchone()
    if not exists:
        return
    columns = {row["name"] for row in connection.execute("PRAGMA table_info(items)")}
    for name, ddl in _MIGRATIONS:
        if name not in columns:
            connection.execute(ddl)
    attachment_exists = connection.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name='attachments'"
    ).fetchone()
    if attachment_exists:
        attachment_columns = {
            row["name"] for row in connection.execute("PRAGMA table_info(attachments)")
        }
        for name, ddl in _ATTACHMENT_MIGRATIONS:
            if name not in attachment_columns:
                connection.execute(ddl)
    connection.commit()


def _normalise_existing_timestamps(connection: sqlite3.Connection) -> int:
    """Rewrite rows written before timestamps were normalised (mixed offsets)."""
    rows = connection.execute(
        "SELECT email_id, received_at FROM items"
        " WHERE received_at != '' AND received_at NOT LIKE '%Z'"
    ).fetchall()
    changed = 0
    for row in rows:
        fixed = normalize_timestamp(row["received_at"])
        if fixed and fixed != row["received_at"]:
            connection.execute(
                "UPDATE items SET received_at = ? WHERE email_id = ?",
                (fixed, row["email_id"]),
            )
            changed += 1
    if changed:
        connection.commit()
    return changed


def match_expression(query: str) -> str:
    """Never interpolate user text into FTS syntax: quote each token as a phrase."""
    tokens = [token for token in (query or "").split() if token]
    return " ".join('"' + token.replace('"', '""') + '"' for token in tokens)


def upsert_rows(
    connection: sqlite3.Connection,
    rows: list[dict[str, Any]],
    folder: str,
) -> int:
    indexed_at = datetime.now(UTC).isoformat(timespec="seconds")
    written = 0
    for row in rows:
        email_id = str(row.get("emailId") or "").strip()
        if not email_id:
            continue
        subject = str(row.get("subject") or "")
        sender_name = str(row.get("senderName") or "")
        sender_address = str(row.get("senderAddress") or "")
        preview = str(row.get("bodyPreview") or "")
        connection.execute(
            """
            INSERT INTO items (
              email_id, folder, subject, sender_name, sender_address, received_at,
              is_read, has_attachments, message_class, internet_message_id, refs,
              body_preview, indexed_at
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(email_id) DO UPDATE SET
              folder=excluded.folder, subject=excluded.subject,
              sender_name=excluded.sender_name, sender_address=excluded.sender_address,
              received_at=excluded.received_at, is_read=excluded.is_read,
              has_attachments=excluded.has_attachments, message_class=excluded.message_class,
              internet_message_id=excluded.internet_message_id, refs=excluded.refs,
              body_preview=excluded.body_preview, indexed_at=excluded.indexed_at
            """,
            (
                email_id,
                folder,
                subject,
                sender_name,
                sender_address,
                normalize_timestamp(str(row.get("receivedAt") or "")),
                0 if row.get("isUnread") else 1,
                1 if row.get("hasAttachments") else 0,
                str(row.get("messageClass") or ""),
                str(row.get("internetMessageId") or ""),
                str(row.get("references") or ""),
                preview,
                indexed_at,
            ),
        )
        connection.execute("DELETE FROM items_fts WHERE email_id = ?", (email_id,))
        connection.execute(
            "INSERT INTO items_fts"
            " (email_id, subject, sender_name, sender_address, body_preview, body_text)"
            " VALUES (?,?,?,?,?,?)",
            (email_id, subject, sender_name, sender_address, preview, ""),
        )
        written += 1
    connection.commit()
    return written


def rows_needing_body(
    connection: sqlite3.Connection,
    *,
    limit: int = 200,
    folder: str | None = None,
) -> list[str]:
    """Ids whose full body has not been fetched yet, newest first."""
    sql = "SELECT email_id FROM items WHERE body_fetched = 0"
    params: list[Any] = []
    if folder:
        sql += " AND folder = ?"
        params.append(folder)
    sql += " ORDER BY received_at DESC LIMIT ?"
    params.append(max(1, int(limit)))
    return [str(row["email_id"]) for row in connection.execute(sql, params).fetchall()]


def store_body(
    connection: sqlite3.Connection,
    email_id: str,
    *,
    body_text: str,
    web_link: str = "",
    attachments: list[dict[str, Any]] | None = None,
    source: str = "getitem",
) -> None:
    """Persist the full body text (and attachment metadata) for one message.

    Bodies live on this machine only; attachments are never downloaded, so what is
    stored for them is a summary line plus the OWA link to open the message.
    """
    connection.execute(
        "UPDATE items SET body_text = ?, body_fetched = 1, web_link = ?, body_source = ?"
        " WHERE email_id = ?",
        (body_text, web_link, source, email_id),
    )
    row = connection.execute(
        "SELECT subject, sender_name, sender_address, body_preview FROM items WHERE email_id = ?",
        (email_id,),
    ).fetchone()
    if row is not None:
        connection.execute("DELETE FROM items_fts WHERE email_id = ?", (email_id,))
        connection.execute(
            "INSERT INTO items_fts"
            " (email_id, subject, sender_name, sender_address, body_preview, body_text)"
            " VALUES (?,?,?,?,?,?)",
            (
                email_id,
                row["subject"],
                row["sender_name"],
                row["sender_address"],
                row["body_preview"],
                body_text,
            ),
        )
    connection.execute("DELETE FROM attachments WHERE email_id = ?", (email_id,))
    for entry in attachments or []:
        connection.execute(
            "INSERT INTO attachments"
            " (email_id, name, content_type, size, summary, attachment_id)"
            " VALUES (?,?,?,?,?,?)"
            " ON CONFLICT(email_id, name) DO UPDATE SET"
            "  content_type=excluded.content_type, size=excluded.size,"
            "  summary=excluded.summary, attachment_id=excluded.attachment_id",
            (
                email_id,
                str(entry.get("name") or ""),
                str(entry.get("contentType") or ""),
                int(entry.get("size") or 0),
                str(entry.get("summary") or ""),
                str(entry.get("attachmentId") or ""),
            ),
        )
    connection.commit()


def attachments_needing_text(
    connection: sqlite3.Connection, *, limit: int = 20, max_bytes: int = 10_000_000
) -> list[dict[str, Any]]:
    rows = connection.execute(
        "SELECT email_id, name, content_type, size, attachment_id FROM attachments"
        " WHERE text_fetched = 0 AND attachment_id != '' AND size > 0 AND size <= ?"
        " ORDER BY size ASC LIMIT ?",
        (max_bytes, max(1, int(limit))),
    ).fetchall()
    return [dict(row) for row in rows]


def store_attachment_text(
    connection: sqlite3.Connection,
    email_id: str,
    name: str,
    *,
    text_body: str,
    method: str,
) -> None:
    connection.execute(
        "UPDATE attachments SET text_body = ?, text_method = ?, text_fetched = 1"
        " WHERE email_id = ? AND name = ?",
        (text_body, method, email_id, name),
    )
    connection.execute(
        "DELETE FROM attachments_fts WHERE email_id = ? AND name = ?", (email_id, name)
    )
    if text_body:
        connection.execute(
            "INSERT INTO attachments_fts (email_id, name, text_body) VALUES (?,?,?)",
            (email_id, name, text_body),
        )
    connection.commit()


def attachment_progress(connection: sqlite3.Connection) -> dict[str, Any]:
    row = connection.execute(
        "SELECT COUNT(*) AS total, SUM(text_fetched) AS fetched,"
        " SUM(LENGTH(text_body)) AS chars FROM attachments"
    ).fetchone()
    return {
        "total": int(row["total"] or 0),
        "fetched": int(row["fetched"] or 0),
        "chars": int(row["chars"] or 0),
    }


def body_progress(connection: sqlite3.Connection) -> dict[str, Any]:
    row = connection.execute(
        "SELECT COUNT(*) AS total, SUM(body_fetched) AS fetched,"
        " SUM(LENGTH(body_text)) AS chars FROM items"
    ).fetchone()
    return {
        "total": int(row["total"] or 0),
        "fetched": int(row["fetched"] or 0),
        "chars": int(row["chars"] or 0),
    }


def search(
    connection: sqlite3.Connection,
    query: str,
    *,
    folder: str | None = None,
    since: str | None = None,
    limit: int = 50,
) -> list[dict[str, Any]]:
    expression = match_expression(query)
    if expression:
        sql = (
            "SELECT items.* FROM items_fts"
            " JOIN items ON items.email_id = items_fts.email_id"
            " WHERE items_fts MATCH ?"
        )
        params: list[Any] = [expression]
    else:
        sql = "SELECT items.* FROM items WHERE 1=1"
        params = []
    if folder:
        sql += " AND items.folder = ?"
        params.append(folder)
    if since:
        sql += " AND items.received_at >= ?"
        params.append(since)
    sql += " ORDER BY items.received_at DESC LIMIT ?"
    params.append(max(1, min(int(limit), 200)))

    found = connection.execute(sql, params).fetchall()
    return [
        {
            "emailId": row["email_id"],
            "storeId": "outlook-web",
            "folder": row["folder"],
            "subject": row["subject"],
            "senderName": row["sender_name"],
            "senderAddress": row["sender_address"],
            "receivedAt": row["received_at"],
            "isUnread": not row["is_read"],
            "hasAttachments": bool(row["has_attachments"]),
            "messageClass": row["message_class"],
            "internetMessageId": row["internet_message_id"],
            "references": row["refs"],
            "bodyPreview": row["body_preview"],
            "indexedAt": row["indexed_at"],
        }
        for row in found
    ]


def prune_folder(
    connection: sqlite3.Connection,
    folder: str,
    *,
    since: str,
    keep_ids: set[str],
) -> int:
    """Remove rows the mailbox no longer has inside a fully scanned window.

    Only safe when the scan covering [since, now] was complete; a truncated scan would
    delete rows it simply never fetched. Mirrors what SyncFolderItems would tell us.
    """
    stale = [
        row["email_id"]
        for row in connection.execute(
            "SELECT email_id FROM items WHERE folder = ? AND received_at >= ?",
            (folder, since),
        ).fetchall()
        if row["email_id"] not in keep_ids
    ]
    for email_id in stale:
        connection.execute("DELETE FROM items WHERE email_id = ?", (email_id,))
        connection.execute("DELETE FROM items_fts WHERE email_id = ?", (email_id,))
    if stale:
        connection.commit()
    return len(stale)


def status(connection: sqlite3.Connection) -> dict[str, Any]:
    row = connection.execute(
        "SELECT COUNT(*) AS rows, COUNT(DISTINCT folder) AS folders,"
        " MIN(received_at) AS oldest, MAX(received_at) AS newest,"
        " MAX(indexed_at) AS last_sync FROM items"
    ).fetchone()
    return {
        "rows": int(row["rows"] or 0),
        "folders": int(row["folders"] or 0),
        "oldest": row["oldest"] or "",
        "newest": row["newest"] or "",
        "lastSync": row["last_sync"] or "",
    }
