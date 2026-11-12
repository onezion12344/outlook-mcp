from __future__ import annotations

import json
import os

import httpx
import pytest

from outlook_web_mcp import pixel_store, pixel_sync
from outlook_web_mcp.config import Settings
from outlook_web_mcp.errors import OutlookWebMcpError

BASE = "https://pixel.example.com"
TOKEN = "d" * 32


def _client(handler):
    return httpx.AsyncClient(transport=httpx.MockTransport(handler))


@pytest.mark.asyncio
async def test_fetch_events_keeps_only_well_formed_rows() -> None:
    payload = [
        {"token": TOKEN, "at": "2026-11-06T10:00:00Z", "ua": "Apple Mail", "ip": "1.2.3.4"},
        {"token": "short", "at": "2026-11-06T10:00:00Z"},
        {"token": "e" * 32},
        "garbage",
    ]

    def handler(request: httpx.Request) -> httpx.Response:
        assert request.headers["authorization"] == "Bearer secret"
        assert request.url.params["since"] == "2026-01-01T00:00:00Z"
        return httpx.Response(200, json=payload)

    async with _client(handler) as client:
        events = await pixel_sync.fetch_events(
            client, base=BASE, sync_token="secret", since="2026-01-01T00:00:00Z"
        )
    assert events == [{"token": TOKEN, "at": "2026-11-06T10:00:00Z", "ua": "Apple Mail"}]
    assert "ip" not in events[0]


@pytest.mark.asyncio
async def test_fetch_events_rejects_a_bad_sync_token() -> None:
    async with _client(lambda _: httpx.Response(401)) as client:
        with pytest.raises(OutlookWebMcpError) as raised:
            await pixel_sync.fetch_events(client, base=BASE, sync_token="nope", since="x")
    assert raised.value.code == "pixel_sync_unauthorized"


@pytest.mark.asyncio
async def test_fetch_events_rejects_unexpected_payload() -> None:
    async with _client(lambda _: httpx.Response(200, json={"nope": True})) as client:
        with pytest.raises(ValueError):
            await pixel_sync.fetch_events(client, base=BASE, sync_token="s", since="x")


@pytest.mark.asyncio
async def test_sync_events_requires_configuration(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(Settings, "pixel_dir", property(lambda self: str(tmp_path / "pixels")))
    settings = Settings(credential_profile="sync-none")
    object.__setattr__(settings, "pixel_base", None)
    with pytest.raises(OutlookWebMcpError) as raised:
        await pixel_sync.sync_events(settings)
    assert raised.value.code == "pixel_not_configured"
    assert "OUTLOOK_WEB_PIXEL_BASE" in str(raised.value)

    object.__setattr__(settings, "pixel_base", BASE)
    object.__setattr__(settings, "pixel_sync_token", None)
    with pytest.raises(OutlookWebMcpError) as raised:
        await pixel_sync.sync_events(settings)
    assert raised.value.code == "pixel_not_configured"
    assert "OUTLOOK_WEB_PIXEL_SYNC_TOKEN" in str(raised.value)


@pytest.mark.asyncio
async def test_sync_events_writes_owner_only_store_and_reports(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(Settings, "pixel_dir", property(lambda self: str(tmp_path / "pixels")))
    settings = Settings(credential_profile="sync-ok")
    object.__setattr__(settings, "pixel_base", BASE)
    object.__setattr__(settings, "pixel_sync_token", "secret")

    def handler(_: httpx.Request) -> httpx.Response:
        return httpx.Response(200, json=[{"token": TOKEN, "at": "2026-11-06T12:00:00Z"}])

    def fake_client(**_kwargs):
        return _client(handler)

    monkeypatch.setattr(pixel_sync, "make_client", fake_client)

    out = await pixel_sync.sync_events(settings)
    assert out["events"] == 1
    assert out["lastAt"] == "2026-11-06T12:00:00Z"
    path = tmp_path / "pixels" / "events.json"
    assert json.loads(path.read_text())[0]["token"] == TOKEN
    if os.name == "posix":
        assert oct(os.stat(path).st_mode & 0o777) == "0o600"

    # Second run asks only for newer events.
    seen: list[str] = []

    def handler2(request: httpx.Request) -> httpx.Response:
        seen.append(request.url.params["since"])
        return httpx.Response(200, json=[])

    def fake_client2(**_kwargs):
        return _client(handler2)

    monkeypatch.setattr(pixel_sync, "make_client", fake_client2)
    await pixel_sync.sync_events(settings)
    assert seen == ["2026-11-06T12:00:00Z"]


def test_store_helpers_handle_missing_and_corrupt_files(tmp_path) -> None:
    directory = str(tmp_path / "pixels")
    assert pixel_store.last_event_at(directory) == ""
    path = tmp_path / "pixels" / "events.json"
    path.parent.mkdir(parents=True)
    path.write_text("{not json", encoding="utf-8")
    assert pixel_store.last_event_at(directory) == ""
