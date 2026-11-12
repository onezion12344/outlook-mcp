from __future__ import annotations

import json
import os
import shutil
import subprocess
from collections.abc import Iterator
from dataclasses import replace
from pathlib import Path

import httpx
import pytest

from outlook_web_mcp import pixel_store, pixel_sync
from outlook_web_mcp.config import Settings
from outlook_web_mcp.errors import OutlookWebMcpError

HARNESS = Path(__file__).with_name("pixel_worker_harness.mjs")
TOKEN = "a" * 32
SYNC_TOKEN = "test-sync-token"


def _node_supports_sqlite() -> bool:
    if shutil.which("node") is None:
        return False
    probe = subprocess.run(
        ["node", "-e", "require('node:sqlite')"],
        capture_output=True,
        check=False,
    )
    return probe.returncode == 0


_NODE_READY = _node_supports_sqlite()
if os.getenv("PIXEL_WORKER_REQUIRE_NODE") == "1" and not _NODE_READY:
    # Skipping here would turn the CI job green while these tests cover nothing.
    raise RuntimeError(
        "PIXEL_WORKER_REQUIRE_NODE=1 but node with node:sqlite (>= 22.5) is unavailable"
    )

pytestmark = pytest.mark.skipif(
    not _NODE_READY, reason="needs Node with the built-in node:sqlite module"
)


@pytest.fixture()
def collector(tmp_path: Path) -> Iterator[tuple[str, Path]]:
    db_path = tmp_path / "events.db"
    child_env = {**os.environ, "SYNC_TOKEN": SYNC_TOKEN}
    process = subprocess.Popen(
        ["node", str(HARNESS), str(db_path)],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        env=child_env,
    )
    try:
        assert process.stdout is not None
        line = process.stdout.readline()
        assert line, f"harness died: {process.stderr.read() if process.stderr else ''}"
        port = json.loads(line)["port"]
        yield f"http://127.0.0.1:{port}", db_path
    finally:
        process.terminate()
        try:
            process.wait(timeout=10)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait(timeout=10)


def _get(base: str, path: str, headers: dict[str, str] | None = None) -> httpx.Response:
    with httpx.Client(trust_env=False, timeout=10.0) as client:
        return client.get(f"{base}{path}", headers=headers or {})


def _events(base: str, query: str = "") -> list[dict[str, object]]:
    response = _get(base, f"/events{query}", {"authorization": f"Bearer {SYNC_TOKEN}"})
    assert response.status_code == 200
    payload = response.json()
    assert isinstance(payload, list)
    return payload


def test_worker_serves_a_gif_and_records_one_event(collector: tuple[str, Path]) -> None:
    base, _ = collector
    response = _get(
        base, f"/o/{TOKEN}.gif", {"user-agent": "Apple Mail", "cf-connecting-ip": "1.2.3.4"}
    )
    assert response.status_code == 200
    assert response.headers["content-type"] == "image/gif"
    # Gmail's image proxy caches; without no-store a second open would be invisible.
    assert "no-store" in response.headers["cache-control"]
    assert response.content.startswith(b"GIF89a")
    # A 1x1 transparent GIF: tiny, and ending in the trailer byte.
    assert response.content.endswith(b";")
    assert len(response.content) < 100

    events = _events(base)
    assert [event["token"] for event in events] == [TOKEN]
    assert events[0]["ua"] == "Apple Mail"
    # The collector never returns the address, only a 4-byte truncated hash.
    assert "1.2.3.4" not in json.dumps(events)


def test_worker_rejects_unknown_paths_and_unauthenticated_reads(
    collector: tuple[str, Path],
) -> None:
    base, _ = collector
    assert _get(base, "/o/short.gif").status_code == 404
    assert _get(base, "/o/" + "b" * 31 + ".gif").status_code == 404
    assert _get(base, "/events").status_code == 401
    assert _get(base, "/events", {"authorization": "Bearer wrong"}).status_code == 401


def test_repeat_fetches_accumulate_and_since_filters(collector: tuple[str, Path]) -> None:
    base, _ = collector
    _get(base, f"/o/{TOKEN}.gif")
    assert len(_events(base)) == 1
    _get(base, f"/o/{TOKEN}.gif", {"user-agent": "second open"})
    both = _events(base)
    assert len(both) == 2

    newest = str(both[-1]["at"])
    assert _events(base, f"?since={newest}") == []


@pytest.mark.asyncio
async def test_sync_events_pulls_from_the_real_worker(
    collector: tuple[str, Path], tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    base, _ = collector
    _get(base, f"/o/{TOKEN}.gif", {"user-agent": "Outlook"})

    pixels = tmp_path / "pixels"
    monkeypatch.setattr(Settings, "pixel_dir", property(lambda self: str(pixels)))
    settings = replace(Settings(), pixel_base=base, pixel_sync_token=SYNC_TOKEN, proxy=None)

    pixel_store.register(
        str(pixels), token=TOKEN, draft_id="draft-1", url="u", at="2026-11-06T00:00:00Z"
    )
    result = await pixel_sync.sync_events(settings)
    assert result["events"] == 1

    stored = json.loads((pixels / "events.json").read_text())
    attributed = pixel_store.events_for(str(pixels), TOKEN, stored)
    assert len(attributed) == 1
    assert attributed[0]["ua"] == "Outlook"
    assert pixel_store.for_draft(str(pixels), "draft-1")["token"] == TOKEN


@pytest.mark.asyncio
async def test_sync_refuses_a_bad_token_against_the_real_worker(
    collector: tuple[str, Path], tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    base, _ = collector
    monkeypatch.setattr(Settings, "pixel_dir", property(lambda self: str(tmp_path / "pixels")))
    settings = replace(Settings(), pixel_base=base, pixel_sync_token="wrong", proxy=None)
    with pytest.raises(OutlookWebMcpError) as raised:
        await pixel_sync.sync_events(settings)
    assert raised.value.code == "pixel_sync_unauthorized"
