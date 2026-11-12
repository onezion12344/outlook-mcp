from __future__ import annotations

import re
from typing import Any

_PREFIX = re.compile(r"^\s*(?:(?:re|fw|fwd)\s*:\s*)+", re.IGNORECASE)
_BRACKET = re.compile(r"^\s*(?:\[[^\]]*\]\s*)+")
_BOUNCE_SUBJECT = re.compile(r"undeliverable|delivery status notification", re.IGNORECASE)
_ADDRESS = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)
_BOUNCE_CLASS = "REPORT.IPM.Note.NDR"


def normalize_subject(subject: str) -> str:
    """Strip leading [tags] and Re:/Fw: prefixes, casefold, for cross-folder matching."""
    text = (subject or "").strip()
    previous = None
    while previous != text:
        previous = text
        text = _BRACKET.sub("", text)
        text = _PREFIX.sub("", text)
    return text.strip().casefold()


def recipient_addresses(draft: dict[str, Any]) -> set[str]:
    found: set[str] = set()
    for key in ("to", "cc"):
        found.update(_ADDRESS.findall(str(draft.get(key) or "")))
    return {address.casefold() for address in found}


def _is_bounce(row: dict[str, Any]) -> bool:
    """Trust `ItemClass` whenever it is present; the subject is only a fallback."""
    message_class = str(row.get("messageClass") or "")
    if message_class:
        return message_class.casefold().startswith(_BOUNCE_CLASS.casefold())
    return bool(_BOUNCE_SUBJECT.search(str(row.get("subject") or "")))


def looks_like_bounce(row: dict[str, Any]) -> bool:
    """Cheap screen; the authoritative MessageClass only arrives via GetItem."""
    return _is_bounce(row)


def _subject_of(row: dict[str, Any]) -> str:
    return normalize_subject(str(row.get("subject") or ""))
    return normalize_subject(str(row.get("subject") or ""))


def classify(
    draft: dict[str, Any],
    sent_rows: list[dict[str, Any]],
    inbox_rows: list[dict[str, Any]],
    *,
    coverage: dict[str, Any] | None = None,
    linked_sent: dict[str, Any] | None = None,
) -> dict[str, object]:
    """Decide sent / replied / bounced for one draft from already-scanned rows.

    Pure and offline: never claims "not replied" when a scan was truncated, because
    the matching message may simply have been beyond the cap.
    """
    target = normalize_subject(str(draft.get("subject") or ""))
    recipients = recipient_addresses(draft)
    sent_message_id = str(draft.get("internetMessageId") or "")

    if linked_sent is not None:
        # We sent this draft ourselves and recorded which item it became: no guessing.
        sent_matches = [linked_sent]
    else:
        sent_matches = [row for row in sent_rows if _subject_of(row) == target]
    # Two sends with the same subject cannot be told apart, so a reply cannot be
    # attributed to this draft either. Refuse to conclude instead of guessing.
    ambiguous = len(sent_matches) > 1
    sent_match = sent_matches[0] if sent_matches else None
    sent_at = str(sent_match.get("receivedAt") or "") if sent_match else ""
    sent_reference = str(sent_match.get("internetMessageId") or "") if sent_match else ""

    bounced_match = next((row for row in inbox_rows if _is_bounce(row)), None)

    replied_match: dict[str, Any] | None = None
    for row in [] if ambiguous else inbox_rows:
        if _is_bounce(row) or _subject_of(row) != target:
            continue
        sender = str(row.get("senderAddress") or "").casefold()
        references = str(row.get("references") or "")
        linked = sender in recipients
        if sent_reference and sent_reference in references:
            linked = True
        if sent_message_id and sent_message_id in references:
            linked = True
        if not linked:
            continue
        received = str(row.get("receivedAt") or "")
        if sent_at and received and received <= sent_at:
            continue
        replied_match = row
        break

    scans = coverage or {}
    truncated = any(
        bool(block.get("truncated")) for block in scans.values() if isinstance(block, dict)
    )
    return {
        "sent": sent_match is not None,
        "replied": replied_match is not None,
        "bounced": bounced_match is not None,
        "evidence": {"sent": sent_match, "replied": replied_match, "bounced": bounced_match},
        "linked": linked_sent is not None,
        "ambiguous": ambiguous,
        "sentCandidates": [str(row.get("emailId") or "") for row in sent_matches[:5]],
        "coverage": {"truncated": truncated, **scans},
        "conclusive": not truncated and not ambiguous,
    }
