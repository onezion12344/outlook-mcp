from __future__ import annotations

import json
import os
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any

FILE_NAME = "sent-links.json"
RETENTION_DAYS = 180


def _path(directory: str) -> Path:
    return Path(directory) / FILE_NAME


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


def _prune(data: dict[str, Any]) -> dict[str, Any]:
    cutoff = (datetime.now(UTC) - timedelta(days=RETENTION_DAYS)).isoformat(timespec="seconds")
    return {
        key: value
        for key, value in data.items()
        if isinstance(value, dict) and str(value.get("at") or "") >= cutoff
    }


def record(
    directory: str,
    *,
    draft_id: str,
    sent_id: str,
    internet_message_id: str,
    at: str,
) -> None:
    """Remember which sent item came from which draft, so tracking stops guessing."""
    path = _path(directory)
    data = _prune(_load(path))
    data[draft_id] = {
        "sentId": sent_id,
        "internetMessageId": internet_message_id,
        "at": at,
    }
    _save(path, data)


def lookup(directory: str, draft_id: str) -> dict[str, Any] | None:
    entry = _load(_path(directory)).get(draft_id)
    return entry if isinstance(entry, dict) else None
