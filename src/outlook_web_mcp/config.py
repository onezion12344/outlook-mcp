from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

from platformdirs import user_cache_path, user_config_path, user_data_path

from .proxy import resolve_proxy

DEFAULT_OUTLOOK_CLIENT_ID = "9199bf20-a13f-4107-85dc-02114787ef48"
OUTLOOK_AUDIENCE = "https://outlook.office.com"
OUTLOOK_SCOPE = f"{OUTLOOK_AUDIENCE}/.default openid profile offline_access"
OWA_ORIGIN = "https://outlook.cloud.microsoft"
OWA_ENDPOINT = f"{OWA_ORIGIN}/owa/service.svc"
STORE_ID = "outlook-web"


def _int_env(name: str, default: int) -> int:
    value = os.getenv(name)
    if value is None:
        return default
    try:
        return int(value)
    except ValueError as exc:
        raise ValueError(f"{name} must be an integer") from exc


@dataclass(frozen=True, slots=True)
class Settings:
    credential_profile: str = field(
        default_factory=lambda: (
            os.getenv("OUTLOOK_WEB_CREDENTIAL_PROFILE", "default").strip() or "default"
        )
    )
    client_id: str = field(
        default_factory=lambda: (
            os.getenv("OUTLOOK_WEB_CLIENT_ID", DEFAULT_OUTLOOK_CLIENT_ID).strip()
            or DEFAULT_OUTLOOK_CLIENT_ID
        )
    )
    time_zone_id: str = field(
        default_factory=lambda: (
            os.getenv("OUTLOOK_WEB_TIME_ZONE_ID", "Tokyo Standard Time").strip()
            or "Tokyo Standard Time"
        )
    )
    request_timeout_seconds: int = field(
        default_factory=lambda: _int_env("OUTLOOK_WEB_REQUEST_TIMEOUT_SECONDS", 30)
    )
    max_response_bytes: int = field(
        default_factory=lambda: _int_env("OUTLOOK_WEB_MAX_RESPONSE_BYTES", 10_000_000)
    )
    max_scanned_items: int = field(
        default_factory=lambda: _int_env("OUTLOOK_WEB_MAX_SCANNED_ITEMS", 20_000)
    )
    max_scanned_folders: int = field(
        default_factory=lambda: _int_env("OUTLOOK_WEB_MAX_SCANNED_FOLDERS", 100)
    )
    log_level: str = field(
        default_factory=lambda: os.getenv("OUTLOOK_WEB_LOG_LEVEL", "INFO").strip().upper()
    )
    verbose_errors: bool = field(
        default_factory=lambda: (
            os.getenv("OUTLOOK_WEB_VERBOSE_ERRORS", "").strip() in {"1", "true", "TRUE", "yes"}
        )
    )
    proxy: str | None = field(default_factory=lambda: resolve_proxy(os.getenv("OUTLOOK_WEB_PROXY")))
    pixel_base: str | None = field(
        default_factory=lambda: (os.getenv("OUTLOOK_WEB_PIXEL_BASE") or "").strip() or None
    )
    track_pixels: bool = field(
        default_factory=lambda: (
            os.getenv("OUTLOOK_WEB_TRACK_PIXEL", "1").strip() not in {"0", "false", "FALSE", "no"}
        )
    )
    pixel_sync_token: str | None = field(
        default_factory=lambda: (os.getenv("OUTLOOK_WEB_PIXEL_SYNC_TOKEN") or "").strip() or None
    )

    @property
    def lock_path(self) -> str:
        root = user_cache_path("outlook-web-mcp-server", "ma-nakaya", ensure_exists=True)
        return str(root / f"auth-{self.credential_profile}.lock")

    @property
    def error_log_path(self) -> str:
        root = user_cache_path("outlook-web-mcp-server", "ma-nakaya", ensure_exists=True)
        return str(root / "errors.log")

    @property
    def approval_dir(self) -> str:
        root = user_config_path("outlook-web-mcp-server", "example", ensure_exists=True)
        return str(root / "approvals")

    @property
    def data_dir(self) -> str:
        """Durable state lives in Application Support, never in Caches.

        Caches are fair game for the OS and for every "clean my Mac" tool — and this
        index did get purged from there once, which is exactly why it moved.
        """
        root = user_data_path("outlook-web-mcp", "example", ensure_exists=True)
        return str(root)

    @property
    def index_path(self) -> str:
        return str(Path(self.data_dir) / "index.db")

    @property
    def sent_links_dir(self) -> str:
        return str(Path(self.data_dir) / "sent-links")

    @property
    def pixel_dir(self) -> str:
        return str(Path(self.data_dir) / "pixels")

    def validate(self) -> None:
        if not self.client_id:
            raise ValueError("OUTLOOK_WEB_CLIENT_ID must not be empty")
        if not self.time_zone_id:
            raise ValueError("OUTLOOK_WEB_TIME_ZONE_ID must not be empty")
        if self.request_timeout_seconds < 1:
            raise ValueError("OUTLOOK_WEB_REQUEST_TIMEOUT_SECONDS must be positive")
        if self.max_response_bytes < 1024:
            raise ValueError("OUTLOOK_WEB_MAX_RESPONSE_BYTES must be at least 1024")
        if not 1 <= self.max_scanned_items <= 50_000:
            raise ValueError("OUTLOOK_WEB_MAX_SCANNED_ITEMS must be between 1 and 50000")
        if not 1 <= self.max_scanned_folders <= 1_000:
            raise ValueError("OUTLOOK_WEB_MAX_SCANNED_FOLDERS must be between 1 and 1000")


def migrate_legacy_state(settings: Settings) -> list[str]:
    """Move durable state out of the old Caches location, once.

    Returns the names moved. Safe to call on every start: it only acts when the
    legacy directory still exists and the new one does not have that entry.
    """
    import shutil

    legacy = user_cache_path("outlook-web-mcp", "example", ensure_exists=False)
    if not legacy.exists():
        return []
    moved: list[str] = []
    target_root = Path(settings.data_dir)
    for name in ("index.db", "sent-links", "pixels"):
        source = legacy / name
        target = target_root / name
        if source.exists() and not target.exists():
            shutil.move(str(source), str(target))
            moved.append(name)
    for suffix in ("-wal", "-shm"):
        stale = legacy / f"index.db{suffix}"
        if stale.exists():
            stale.unlink()
    return moved
