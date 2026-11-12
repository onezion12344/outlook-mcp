from __future__ import annotations

import re
import subprocess

_SCUTIL_TIMEOUT_SECONDS = 2.0
_PATTERN = re.compile(r"^\s*(HTTPEnable|HTTPProxy|HTTPSPort)\s*:\s*(.+?)\s*$", re.MULTILINE)


def parse_scutil_proxy(text: str) -> str | None:
    """Read the HTTP/HTTPS proxy out of `scutil --proxy` output. Never raises."""
    try:
        values = {match.group(1): match.group(2) for match in _PATTERN.finditer(text)}
    except Exception:
        return None
    if values.get("HTTPEnable") != "1":
        return None
    host = values.get("HTTPProxy", "").strip()
    port = values.get("HTTPSPort", "").strip()
    if not host or not port.isdigit():
        return None
    return f"http://{host}:{port}"


def _scutil_output() -> str:
    try:
        completed = subprocess.run(
            ["scutil", "--proxy"],
            capture_output=True,
            text=True,
            timeout=_SCUTIL_TIMEOUT_SECONDS,
            check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return ""
    return completed.stdout if completed.returncode == 0 else ""


def resolve_proxy(explicit: str | None = None) -> str | None:
    """Explicit value wins; otherwise the macOS system proxy; otherwise direct."""
    candidate = (explicit or "").strip()
    if candidate:
        return candidate
    return parse_scutil_proxy(_scutil_output())
