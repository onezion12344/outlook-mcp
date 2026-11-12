from __future__ import annotations

import contextlib
import hashlib
import json
import os
import re
import time
from pathlib import Path

DIGEST_LENGTH = 16
# An approved-but-unsent draft is an armed action; it must not stay armed forever.
APPROVAL_TTL_SECONDS = 30 * 60
_ADDRESS = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)


def _normalise_recipients(value: str) -> str:
    return ",".join(sorted({address.casefold() for address in _ADDRESS.findall(value or "")}))


def draft_digest(
    *,
    draft_id: str,
    subject: str,
    recipients: str,
    body: str,
    bcc: str = "",
    attachments: tuple[str, ...] | list[str] = (),
) -> str:
    """Bind an approval to the exact draft content the user reviewed.

    Covers visible recipients, blind recipients, attachment fingerprints and the body,
    because any of those can be changed after approval without touching the others.
    """
    material = "\x1f".join(
        [
            draft_id,
            (subject or "").strip(),
            _normalise_recipients(recipients),
            _normalise_recipients(bcc),
            "|".join(sorted(str(item) for item in attachments)),
            hashlib.sha256((body or "").encode("utf-8")).hexdigest(),
        ]
    )
    return hashlib.sha256(material.encode("utf-8")).hexdigest()[:DIGEST_LENGTH]


def approval_path(directory: str, digest: str) -> Path:
    return Path(directory) / f"{digest}.nonce"


def approve(directory: str, digest: str, *, ttl_seconds: int = APPROVAL_TTL_SECONDS) -> Path:
    """Create the one-time token the user authorises in a terminal."""
    path = approval_path(directory, digest)
    path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    path.write_text(
        json.dumps({"at": int(time.time()), "ttl": int(ttl_seconds)}),
        encoding="utf-8",
    )
    os.chmod(path, 0o600)
    purge_expired(directory, ttl_seconds=ttl_seconds)
    return path


def _issued_at(path: Path) -> int | None:
    try:
        raw = path.read_text(encoding="utf-8").strip()
    except OSError:
        return None
    try:
        return int(raw.split("\n")[0])  # legacy plain epoch
    except ValueError:
        pass
    try:
        return int(json.loads(raw).get("at"))
    except (json.JSONDecodeError, TypeError, ValueError, AttributeError):
        return None


def purge_expired(directory: str, *, ttl_seconds: int = APPROVAL_TTL_SECONDS) -> int:
    """Drop approvals that were never used in time."""
    removed = 0
    for path in Path(directory).glob("*.nonce"):
        issued = _issued_at(path)
        if issued is None or time.time() - issued > ttl_seconds:
            with contextlib.suppress(OSError):
                path.unlink()
                removed += 1
    return removed


def consume(directory: str, digest: str, *, ttl_seconds: int = APPROVAL_TTL_SECONDS) -> bool:
    """Return True exactly once per approved digest, and only while it is fresh."""
    path = approval_path(directory, digest)
    issued = _issued_at(path)
    if issued is None:
        with contextlib.suppress(OSError):
            path.unlink()
        return False
    with contextlib.suppress(OSError):
        path.unlink()
    return time.time() - issued <= ttl_seconds
