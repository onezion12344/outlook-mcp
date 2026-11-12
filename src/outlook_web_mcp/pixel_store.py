from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

REGISTRY_FILE = "pixels.json"


def _path(directory: str) -> Path:
    return Path(directory) / REGISTRY_FILE


def _load(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}
    return data if isinstance(data, dict) else {}


def _save(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, separators=(",", ":"), ensure_ascii=True), encoding="utf-8")
    os.chmod(tmp, 0o600)
    tmp.replace(path)


def register(directory: str, *, token: str, draft_id: str, url: str, at: str) -> None:
    path = _path(directory)
    data = _load(path)
    data[token] = {"draftId": draft_id, "url": url, "createdAt": at, "sentAt": None}
    _save(path, data)


def mark_sent(directory: str, token: str, at: str) -> bool:
    path = _path(directory)
    data = _load(path)
    entry = data.get(token)
    if not isinstance(entry, dict):
        return False
    entry["sentAt"] = at
    _save(path, data)
    return True


def known(directory: str, token: str) -> bool:
    return isinstance(_load(_path(directory)).get(token), dict)


def for_draft(directory: str, draft_id: str) -> dict[str, Any] | None:
    for token, entry in _load(_path(directory)).items():
        if isinstance(entry, dict) and entry.get("draftId") == draft_id:
            return {"token": token, **entry}
    return None


def events_for(directory: str, token: str, events: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Filter collected open events down to one token; keeps the payload minimal."""
    return [
        {"at": str(event.get("at") or ""), "ua": str(event.get("ua") or "")[:120]}
        for event in events
        if str(event.get("token") or "") == token
    ]


def replace_events(directory: str, events: list[dict[str, Any]]) -> None:
    """Write the collected event set atomically, owner-only."""
    path = Path(directory) / "events.json"
    path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(events, separators=(",", ":"), ensure_ascii=True), encoding="utf-8")
    os.chmod(tmp, 0o600)
    tmp.replace(path)


def last_event_at(directory: str) -> str:
    path = Path(directory) / "events.json"
    if not path.exists():
        return ""
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return ""
    if not isinstance(data, list):
        return ""
    stamps = sorted(str(item.get("at") or "") for item in data if isinstance(item, dict))
    return stamps[-1] if stamps else ""
