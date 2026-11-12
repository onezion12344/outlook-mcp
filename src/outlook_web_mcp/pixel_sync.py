from __future__ import annotations

import ssl
from typing import Any

import httpx

from . import pixel_store
from .config import Settings
from .errors import OutlookWebMcpError
from .tls import default_context


async def fetch_events(
    client: httpx.AsyncClient,
    *,
    base: str,
    sync_token: str,
    since: str,
) -> list[dict[str, Any]]:
    """Pull collected open events. The collector never returns IP addresses."""
    response = await client.get(
        f"{base.rstrip('/')}/events",
        params={"since": since},
        headers={"Authorization": f"Bearer {sync_token}"},
    )
    if response.status_code == 401:
        raise OutlookWebMcpError(
            "The collector rejected the sync token.", "pixel_sync_unauthorized"
        )
    response.raise_for_status()
    payload: object = response.json()
    if not isinstance(payload, list):
        raise ValueError("collector returned an unexpected payload")
    events: list[dict[str, Any]] = []
    for item in payload:
        if not isinstance(item, dict):
            continue
        token = str(item.get("token") or "")
        at = str(item.get("at") or "")
        if len(token) != 32 or not at:
            continue
        events.append({"token": token, "at": at, "ua": str(item.get("ua") or "")[:120]})
    return events


def make_client(
    *,
    context: ssl.SSLContext,
    proxy: str | None,
    timeout: float = 20.0,
) -> httpx.AsyncClient:
    """Same posture as the mail client: pinned trust root, no env proxy parsing."""
    return httpx.AsyncClient(
        timeout=timeout,
        follow_redirects=False,
        verify=context,
        trust_env=False,
        proxy=proxy,
    )


async def sync_events(settings: Settings, *, since: str = "") -> dict[str, Any]:
    """Fetch collected events and write them where track_status reads them."""
    base = (settings.pixel_base or "").rstrip("/")
    token = settings.pixel_sync_token
    if not base:
        raise OutlookWebMcpError(
            "OUTLOOK_WEB_PIXEL_BASE is not configured; the collector is not wired up.",
            "pixel_not_configured",
        )
    if not token:
        raise OutlookWebMcpError(
            "OUTLOOK_WEB_PIXEL_SYNC_TOKEN is not configured.", "pixel_not_configured"
        )
    since_value = since or pixel_store.last_event_at(settings.pixel_dir) or "1970-01-01T00:00:00Z"
    async with make_client(context=default_context(), proxy=settings.proxy) as client:
        events = await fetch_events(client, base=base, sync_token=token, since=since_value)
    pixel_store.replace_events(settings.pixel_dir, events)
    return {
        "base": base,
        "since": since_value,
        "events": len(events),
        "lastAt": pixel_store.last_event_at(settings.pixel_dir),
    }
