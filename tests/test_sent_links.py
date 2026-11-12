from __future__ import annotations

import os

import pytest
from test_service import StubApi, StubAuth, _folder, _message, _success

from outlook_web_mcp import sent_links
from outlook_web_mcp.config import Settings
from outlook_web_mcp.track import classify

DRAFT = {
    "emailId": "draft-1",
    "subject": "JUPAS enquiry",
    "to": "Registry <registry@example.com>",
    "cc": "",
    "internetMessageId": "<draft-1@tenant>",
}


def test_record_and_lookup_round_trip(tmp_path) -> None:
    directory = str(tmp_path / "sent-links")
    sent_links.record(
        directory,
        draft_id="draft-1",
        sent_id="sent-9",
        internet_message_id="<sent-9@tenant>",
        at="2026-11-06T09:00:00Z",
    )
    entry = sent_links.lookup(directory, "draft-1")
    assert entry is not None
    assert entry["sentId"] == "sent-9"
    if os.name == "posix":
        assert oct(os.stat(tmp_path / "sent-links" / "sent-links.json").st_mode & 0o777) == "0o600"


def test_missing_and_corrupt_store_degrade_to_none(tmp_path) -> None:
    directory = str(tmp_path / "sent-links")
    assert sent_links.lookup(directory, "nope") is None
    path = tmp_path / "sent-links" / "sent-links.json"
    path.parent.mkdir(parents=True)
    path.write_text("{not json", encoding="utf-8")
    assert sent_links.lookup(directory, "nope") is None


def test_a_linked_send_beats_subject_matching_and_kills_ambiguity() -> None:
    sent_rows = [
        {"emailId": "s-1", "subject": "JUPAS enquiry", "receivedAt": "2026-11-04T09:00:00Z"},
        {"emailId": "s-2", "subject": "JUPAS enquiry", "receivedAt": "2026-11-03T09:00:00Z"},
    ]
    unlinked = classify(DRAFT, sent_rows, [])
    assert unlinked["ambiguous"] is True and unlinked["conclusive"] is False

    linked = {
        "emailId": "s-2",
        "subject": "JUPAS enquiry",
        "receivedAt": "2026-11-05T09:00:00Z",
        "internetMessageId": "<s-2@tenant>",
        "messageClass": "IPM.Note",
        "senderAddress": "",
        "references": "",
    }
    result = classify(DRAFT, sent_rows, [], linked_sent=linked)
    assert result["linked"] is True
    assert result["ambiguous"] is False
    assert result["conclusive"] is True
    assert result["sentCandidates"] == ["s-2"]


def test_a_linked_reply_is_matched_by_reference() -> None:
    linked = {
        "emailId": "s-2",
        "subject": "JUPAS enquiry",
        "receivedAt": "2026-11-05T09:00:00Z",
        "internetMessageId": "<s-2@tenant>",
        "messageClass": "IPM.Note",
        "senderAddress": "",
        "references": "",
    }
    reply = {
        "emailId": "i-1",
        "subject": "Re: JUPAS enquiry",
        "senderAddress": "someone-else@example.org",  # not in To/Cc
        "receivedAt": "2026-11-06T01:00:00Z",
        "references": "<s-2@tenant>",
        "messageClass": "IPM.Note",
    }
    result = classify(DRAFT, [linked], [reply], linked_sent=linked)
    assert result["replied"] is True


@pytest.mark.asyncio
async def test_track_status_uses_the_recorded_link(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(
        Settings, "sent_links_dir", property(lambda self: str(tmp_path / "sent-links"))
    )
    sent_links.record(
        str(tmp_path / "sent-links"),
        draft_id="draft-1",
        sent_id="s-2",
        internet_message_id="<s-2@tenant>",
        at="2026-11-05T09:00:00Z",
    )

    draft_item = {
        "ItemId": {"Id": "draft-1", "ChangeKey": "k"},
        "Subject": "JUPAS enquiry",
        "ToRecipients": [{"Mailbox": {"Name": "Registry", "EmailAddress": "registry@example.com"}}],
        "CcRecipients": [],
        "InternetMessageId": "<draft-1@tenant>",
        "Body": {"BodyType": "Text", "Value": "hello"},
        "DateTimeReceived": "2026-11-05T08:00:00Z",
        "IsRead": True,
        "HasAttachments": False,
    }
    same_subject = _message("s-1")
    same_subject["Subject"] = "JUPAS enquiry"
    same_subject["DateTimeReceived"] = "2026-11-04T09:00:00Z"
    reply = _message("i-1")
    reply["Subject"] = "Re: JUPAS enquiry"
    reply["From"] = {"Mailbox": {"Name": "Registry", "EmailAddress": "registry@example.com"}}
    reply["DateTimeReceived"] = "2026-11-06T02:00:00Z"

    empty = _success(RootFolder={"Items": [], "IncludesLastItemInRange": True})
    api = StubApi(
        [
            _success(Items=[draft_item]),
            _success(Folders=[_folder("s-1", "Sent")]),
            _success(RootFolder={"Items": [same_subject], "IncludesLastItemInRange": True}),
            _success(Folders=[_folder("i-1", "Inbox")]),
            _success(RootFolder={"Items": [reply], "IncludesLastItemInRange": True}),
            _success(Folders=[_folder("a-1", "Archive")]),
            empty,
        ]
    )
    from outlook_web_mcp.service import OutlookService

    service = OutlookService(  # type: ignore[arg-type]
        StubAuth(), api, Settings(credential_profile="link-test")
    )
    out = await service.track_status(draft_id="draft-1", store_id="outlook-web")
    assert out["linked"] is True
    assert out["ambiguous"] is False
    assert out["replied"] is True


@pytest.mark.asyncio
async def test_window_extends_to_cover_the_send_date(tmp_path, monkeypatch) -> None:
    from datetime import UTC, datetime, timedelta

    monkeypatch.setattr(
        Settings, "sent_links_dir", property(lambda self: str(tmp_path / "sent-links"))
    )
    sent_at = (datetime.now(UTC) - timedelta(days=60)).isoformat(timespec="seconds")
    sent_links.record(
        str(tmp_path / "sent-links"),
        draft_id="draft-1",
        sent_id="s-9",
        internet_message_id="<s-9@tenant>",
        at=sent_at,
    )
    draft_item = {
        "ItemId": {"Id": "draft-1", "ChangeKey": "k"},
        "Subject": "JUPAS enquiry",
        "ToRecipients": [{"Mailbox": {"EmailAddress": "registry@example.com"}}],
        "CcRecipients": [],
        "InternetMessageId": "<draft-1@tenant>",
        "Body": {"BodyType": "Text", "Value": "hello"},
        "DateTimeReceived": "2026-11-05T08:00:00Z",
        "IsRead": True,
        "HasAttachments": False,
    }
    empty = _success(RootFolder={"Items": [], "IncludesLastItemInRange": True})
    api = StubApi(
        [
            _success(Items=[draft_item]),
            _success(Folders=[_folder("s-1", "Sent")]),
            empty,
            _success(Folders=[_folder("i-1", "Inbox")]),
            empty,
            _success(Folders=[_folder("a-1", "Archive")]),
            empty,
        ]
    )
    from outlook_web_mcp.service import OutlookService

    service = OutlookService(  # type: ignore[arg-type]
        StubAuth(), api, Settings(credential_profile="window-test")
    )
    out = await service.track_status(draft_id="draft-1", store_id="outlook-web")
    assert out["windowDays"] >= 67  # 60 days since sending + 7 days of slack
