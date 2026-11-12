from __future__ import annotations

from dataclasses import dataclass
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any

from . import approval, pixel_store
from .config import Settings
from .credentials import KeyringCredentialStore
from .errors import OutlookWebMcpError
from .index import open_index
from .index import status as index_status
from .tls import default_context

STALE_INDEX_DAYS = 7
STALE_PIXEL_HOURS = 48


@dataclass(slots=True)
class Check:
    name: str
    status: str  # "ok" | "warn" | "fail"
    detail: str

    def to_dict(self) -> dict[str, str]:
        return {"name": self.name, "status": self.status, "detail": self.detail}


def _keyring(settings: Settings) -> Check:
    import keyring

    backend = keyring.get_keyring()
    name = type(backend).__name__
    try:
        priority = float(backend.priority)
    except (TypeError, ValueError, RuntimeError):
        return Check("keyring", "fail", f"{name}: no usable priority")
    if priority <= 0:
        return Check("keyring", "fail", f"{name}: priority {priority} is not a secure store")
    return Check("keyring", "ok", f"{name} (priority {priority:g})")


def _credentials(settings: Settings) -> Check:
    try:
        bundle = KeyringCredentialStore(settings.credential_profile).load()
    except OutlookWebMcpError as exc:
        return Check("credentials", "fail", f"{exc.code}: {exc}")
    if bundle is None:
        return Check("credentials", "warn", "no stored credentials; run `outlook-web-mcp login`")
    return Check("credentials", "ok", f"signed in as {bundle.display_name}")


def _index(settings: Settings) -> Check:
    path = Path(settings.index_path)
    if not path.exists():
        return Check("index", "warn", "not built yet; run index_sync")
    try:
        connection = open_index(settings.index_path)
        try:
            info = index_status(connection)
        finally:
            connection.close()
    except Exception as exc:
        return Check("index", "fail", f"unreadable: {type(exc).__name__}")
    if info["rows"] == 0:
        return Check("index", "warn", "empty; run index_sync")
    last = str(info.get("lastSync") or "")
    try:
        age = datetime.now(UTC) - datetime.fromisoformat(last.replace("Z", "+00:00"))
    except ValueError:
        return Check("index", "warn", f"{info['rows']} rows; lastSync unreadable ({last!r})")
    stale = age > timedelta(days=STALE_INDEX_DAYS)
    detail = (
        f"{info['rows']} rows / {info['folders']} folders; "
        f"last sync {age.days}d {age.seconds // 3600}h ago"
    )
    return Check("index", "warn" if stale else "ok", detail)


def _approvals(settings: Settings) -> Check:
    directory = Path(settings.approval_dir)
    if not directory.exists():
        return Check("approvals", "ok", "no pending approvals")
    now = datetime.now(UTC).timestamp()
    pending = sorted(directory.glob("*.nonce"))
    expired = 0
    live = 0
    for path in pending:
        issued = approval._issued_at(path)
        if issued is None or now - issued > approval.APPROVAL_TTL_SECONDS:
            expired += 1
        else:
            live += 1
    if expired:
        return Check(
            "approvals", "warn", f"{live} live, {expired} expired (purged on next approve)"
        )
    return Check("approvals", "ok", f"{live} live")


def _pixels(settings: Settings) -> Check:
    if not settings.pixel_base:
        return Check("pixels", "ok", "tracking disabled (OUTLOOK_WEB_PIXEL_BASE unset)")
    if not settings.pixel_sync_token:
        return Check("pixels", "warn", "collector configured but no SYNC_TOKEN")
    directory = Path(settings.pixel_dir)
    last = pixel_store.last_event_at(str(directory))
    if not last:
        return Check("pixels", "warn", f"{settings.pixel_base}: no events synced yet")
    try:
        age = datetime.now(UTC) - datetime.fromisoformat(last.replace("Z", "+00:00"))
    except ValueError:
        return Check("pixels", "warn", f"{settings.pixel_base}: last event unreadable")
    stale = age > timedelta(hours=STALE_PIXEL_HOURS)
    return Check(
        "pixels",
        "warn" if stale else "ok",
        f"{settings.pixel_base}: last event {age.seconds // 3600}h ago",
    )


def _network(settings: Settings) -> Check:
    try:
        context = default_context()
    except Exception as exc:
        return Check("tls", "fail", f"cannot build trust context: {type(exc).__name__}")
    certs = len(context.get_ca_certs())
    proxy = settings.proxy or "direct"
    return Check("tls", "ok" if certs else "fail", f"{certs} CAs pinned to certifi; egress {proxy}")


def run_checks(settings: Settings) -> list[Check]:
    """Local-only health checks. Never prints tokens, message content or credentials."""
    checks = [
        _keyring(settings),
        _credentials(settings),
        _index(settings),
        _approvals(settings),
        _pixels(settings),
        _network(settings),
    ]
    return checks


def summarise(checks: list[Check]) -> dict[str, Any]:
    return {
        "ok": all(check.status != "fail" for check in checks),
        "counts": {
            status: sum(1 for check in checks if check.status == status)
            for status in ("ok", "warn", "fail")
        },
        "checks": [check.to_dict() for check in checks],
    }
