from __future__ import annotations

import pytest
from test_service import StubApi, StubAuth, _folder, _message, _success

from outlook_web_mcp.config import Settings
from outlook_web_mcp.service import OutlookService
from outlook_web_mcp.track import classify


def _row(
    email_id: str,
    *,
    subject: str,
    sender: str = "student@example.com",
    received: str = "2026-11-05T09:00:00Z",
    message_class: str = "IPM.Note",
    references: str = "",
    internet_message_id: str = "",
) -> dict[str, object]:
    return {
        "emailId": email_id,
        "subject": subject,
        "senderAddress": sender,
        "receivedAt": received,
        "messageClass": message_class,
        "references": references,
        "internetMessageId": internet_message_id,
    }


DRAFT = {
    "emailId": "draft-1",
    "subject": "JUPAS enquiry",
    "to": "Registry <registry@example.com>; CDS <cdsug@example.edu>",
    "cc": "",
    "internetMessageId": "<draft-1@tenant>",
}


def test_classify_reports_all_three_when_they_happened() -> None:
    sent = _row(
        "s-1",
        subject="JUPAS enquiry",
        internet_message_id="<sent-1@tenant>",
        received="2026-11-05T09:00:00Z",
    )
    inbox = [
        _row(
            "i-1",
            subject="Re: JUPAS enquiry",
            sender="registry@example.com",
            received="2026-11-06T02:00:00Z",
            references="<sent-1@tenant>",
        )
    ]
    result = classify(DRAFT, [sent], inbox)
    assert result["sent"] is True
    assert result["replied"] is True
    assert result["bounced"] is False
    assert result["evidence"]["sent"]["emailId"] == "s-1"
    assert result["evidence"]["replied"]["emailId"] == "i-1"


def test_classify_is_honest_when_nothing_matched() -> None:
    result = classify(DRAFT, [], [])
    assert result["sent"] is False
    assert result["replied"] is False
    assert result["bounced"] is False
    assert result["evidence"] == {"sent": None, "replied": None, "bounced": None}


def test_classify_detects_bounce_from_report_class() -> None:
    sent = _row("s-1", subject="JUPAS enquiry")
    bounce = _row(
        "i-9",
        subject="Undeliverable: JUPAS enquiry",
        sender="postmaster@example.edu",
        message_class="REPORT.IPM.Note.NDR",
    )
    result = classify(DRAFT, [sent], [bounce])
    assert result["bounced"] is True
    assert result["replied"] is False
    assert result["evidence"]["bounced"]["emailId"] == "i-9"


def test_reply_must_come_from_a_recipient_not_a_stranger() -> None:
    sent = _row("s-1", subject="JUPAS enquiry")
    inbox = [_row("i-2", subject="Re: JUPAS enquiry", sender="spammer@example.org")]
    result = classify(DRAFT, [sent], inbox)
    assert result["replied"] is False


def test_reply_before_send_does_not_count() -> None:
    sent = _row("s-1", subject="JUPAS enquiry", received="2026-11-05T09:00:00Z")
    inbox = [
        _row(
            "i-3",
            subject="Re: JUPAS enquiry",
            sender="registry@example.com",
            received="2026-11-04T01:00:00Z",
        )
    ]
    result = classify(DRAFT, [sent], inbox)
    assert result["replied"] is False


def test_subject_normalisation_ignores_re_and_brackets() -> None:
    draft = {
        "emailId": "d",
        "subject": "[Urgent] Re: JUPAS enquiry",
        "to": "alice@example.com",
        "cc": "",
        "internetMessageId": "",
    }
    sent = _row("s-2", subject="JUPAS enquiry")
    assert classify(draft, [sent], [])["sent"] is True


def test_coverage_flags_when_scans_truncated() -> None:
    result = classify(
        DRAFT,
        [],
        [],
        coverage={"sent": {"truncated": True}, "inbox": {"truncated": False}},
    )
    assert result["coverage"]["truncated"] is True
    assert result["conclusive"] is False


@pytest.mark.asyncio
async def test_track_status_scans_sent_and_inbox_and_reports() -> None:
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
    sent_row = _message("s-1")
    sent_row["Subject"] = "JUPAS enquiry"
    sent_row["MessageClass"] = "IPM.Note"
    sent_row["DateTimeReceived"] = "2026-11-04T09:00:00Z"
    reply_row = _message("i-1")
    reply_row["Subject"] = "Re: JUPAS enquiry"
    reply_row["From"] = {"Mailbox": {"Name": "Registry", "EmailAddress": "registry@example.com"}}
    reply_row["MessageClass"] = "IPM.Note"
    reply_row["DateTimeReceived"] = "2026-11-05T02:00:00Z"

    api = StubApi(
        [
            _success(Items=[draft_item]),
            _success(Folders=[_folder("sent-1", "Sent Items")]),
            _success(RootFolder={"Items": [sent_row], "IncludesLastItemInRange": True}),
            _success(Folders=[_folder("in-1", "Inbox")]),
            _success(RootFolder={"Items": [reply_row], "IncludesLastItemInRange": True}),
            _success(Folders=[_folder("ar-1", "Archive")]),
            _success(RootFolder={"Items": [], "IncludesLastItemInRange": True}),
        ]
    )
    service = OutlookService(  # type: ignore[arg-type]
        StubAuth(), api, Settings(credential_profile="track-test")
    )

    out = await service.track_status(draft_id="draft-1", store_id="outlook-web")
    assert out["sent"] is True
    assert out["replied"] is True
    assert out["bounced"] is False
    assert out["conclusive"] is True
    assert out["draftId"] == "draft-1"
    actions = [call[0] for call in api.calls]
    assert actions == [
        "GetItem",
        "GetFolder",
        "FindItem",
        "GetFolder",
        "FindItem",
        "GetFolder",
        "FindItem",
    ]


@pytest.mark.asyncio
async def test_reply_filed_into_archive_still_counts(tmp_path, monkeypatch) -> None:
    draft_item = {
        "ItemId": {"Id": "draft-2", "ChangeKey": "k"},
        "Subject": "JUPAS enquiry",
        "ToRecipients": [{"Mailbox": {"Name": "Registry", "EmailAddress": "registry@example.com"}}],
        "CcRecipients": [],
        "InternetMessageId": "<draft-2@tenant>",
        "Body": {"BodyType": "Text", "Value": "hello"},
        "DateTimeReceived": "2026-11-05T08:00:00Z",
        "IsRead": True,
        "HasAttachments": False,
    }
    archived_reply = _message("a-1")
    archived_reply["Subject"] = "Re: JUPAS enquiry"
    archived_reply["From"] = {
        "Mailbox": {"Name": "Registry", "EmailAddress": "registry@example.com"}
    }
    archived_reply["DateTimeReceived"] = "2026-11-06T02:00:00Z"

    api = StubApi(
        [
            _success(Items=[draft_item]),
            _success(Folders=[_folder("sent-2", "Sent Items")]),
            _success(RootFolder={"Items": [], "IncludesLastItemInRange": True}),
            _success(Folders=[_folder("in-2", "Inbox")]),
            _success(RootFolder={"Items": [], "IncludesLastItemInRange": True}),
            _success(Folders=[_folder("ar-2", "Archive")]),
            _success(RootFolder={"Items": [archived_reply], "IncludesLastItemInRange": True}),
        ]
    )
    service = OutlookService(  # type: ignore[arg-type]
        StubAuth(), api, Settings(credential_profile="track-archive")
    )

    out = await service.track_status(draft_id="draft-2", store_id="outlook-web")
    assert out["replied"] is True
    assert out["evidence"]["replied"]["emailId"] == "a-1"


@pytest.mark.asyncio
async def test_opened_is_reported_only_from_collected_events(tmp_path, monkeypatch) -> None:
    import json

    from test_service import StubApi as Api

    from outlook_web_mcp import pixel, pixel_store
    from outlook_web_mcp.service import OutlookService as Service

    token = pixel.new_token()
    monkeypatch.setattr(Settings, "pixel_dir", property(lambda self: str(tmp_path / "pixels")))
    settings = Settings(credential_profile="track-open")
    pixel_store.register(
        settings.pixel_dir,
        token=token,
        draft_id="draft-1",
        url=pixel.tracking_url("https://pixel.example.com", token),
        at="2026-11-06T10:00:00Z",
    )
    events_file = tmp_path / "pixels" / "events.json"
    events_file.write_text(
        json.dumps([{"token": token, "at": "2026-11-06T12:00:00Z", "ua": "Apple Mail"}]),
        encoding="utf-8",
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
    empty = _success(RootFolder={"Items": [], "IncludesLastItemInRange": True})
    api = Api(
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
    service = Service(StubAuth(), api, settings)  # type: ignore[arg-type]

    out = await service.track_status(draft_id="draft-1", store_id="outlook-web")
    assert out["tracked"] is True
    assert out["opened"]["count"] == 1
    assert "low" in out["opened"]["confidence"]
    assert "ip" not in json.dumps(out["opened"]).lower()


def test_two_sends_with_the_same_subject_are_flagged_as_ambiguous() -> None:
    sent = [
        _row("s-1", subject="JUPAS enquiry", received="2026-11-05T09:00:00Z"),
        _row("s-2", subject="JUPAS enquiry", received="2026-11-04T09:00:00Z"),
    ]
    reply = _row(
        "i-1",
        subject="Re: JUPAS enquiry",
        sender="registry@example.com",
        received="2026-11-06T02:00:00Z",
    )
    result = classify(DRAFT, sent, [reply])
    assert result["ambiguous"] is True
    assert len(result["sentCandidates"]) == 2
    # A reply exists, but it cannot be tied to this draft: refuse to claim it.
    assert result["replied"] is False
    assert result["conclusive"] is False


def test_single_send_is_not_ambiguous() -> None:
    sent = [_row("s-1", subject="JUPAS enquiry")]
    result = classify(DRAFT, sent, [])
    assert result["ambiguous"] is False
    assert result["sentCandidates"] == ["s-1"]
    assert result["conclusive"] is True


@pytest.mark.asyncio
async def test_bounce_class_comes_from_the_search_row_without_extra_reads(
    tmp_path, monkeypatch
) -> None:
    draft_item = {
        "ItemId": {"Id": "draft-4", "ChangeKey": "k"},
        "Subject": "JUPAS enquiry",
        "ToRecipients": [{"Mailbox": {"Name": "Registry", "EmailAddress": "registry@example.com"}}],
        "CcRecipients": [],
        "InternetMessageId": "<draft-4@tenant>",
        "Body": {"BodyType": "Text", "Value": "hello"},
        "DateTimeReceived": "2026-11-05T08:00:00Z",
        "IsRead": True,
        "HasAttachments": False,
    }
    bounce = _message("b-1")
    bounce["Subject"] = "Undeliverable: JUPAS enquiry"
    bounce["MessageClass"] = "REPORT.IPM.Note.NDR"

    empty = _success(RootFolder={"Items": [], "IncludesLastItemInRange": True})
    api = StubApi(
        [
            _success(Items=[draft_item]),
            _success(Folders=[_folder("s-4", "Sent")]),
            empty,
            _success(Folders=[_folder("i-4", "Inbox")]),
            _success(RootFolder={"Items": [bounce], "IncludesLastItemInRange": True}),
            _success(Folders=[_folder("a-4", "Archive")]),
            empty,
        ]
    )
    service = OutlookService(  # type: ignore[arg-type]
        StubAuth(), api, Settings(credential_profile="track-bounce-row")
    )

    out = await service.track_status(draft_id="draft-4", store_id="outlook-web")
    assert out["bounced"] is True
    # Only the draft itself is read; the class already arrived with the search row.
    assert len([call for call in api.calls if call[0] == "GetItem"]) == 1


def test_an_unrelated_message_mentioning_undeliverable_is_not_a_bounce() -> None:
    impostor = _row(
        "i-9",
        subject="Undeliverable: how should we word this?",
        sender="colleague@example.edu",
        message_class="IPM.Note",
    )
    result = classify(DRAFT, [], [impostor])
    assert result["bounced"] is False
