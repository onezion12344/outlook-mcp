from __future__ import annotations

from datetime import UTC, datetime, timedelta
from typing import Any

import pytest

from outlook_web_mcp.config import STORE_ID, Settings
from outlook_web_mcp.service import OutlookService


def _success(**values: object) -> dict[str, Any]:
    return {
        "Body": {
            "ResponseMessages": {
                "Items": [
                    {
                        "ResponseClass": "Success",
                        "ResponseCode": "NoError",
                        **values,
                    }
                ]
            }
        }
    }


def _folder(folder_id: str = "folder-1", name: str = "Inbox") -> dict[str, object]:
    return {
        "FolderId": {"Id": folder_id, "ChangeKey": "folder-key"},
        "DisplayName": name,
        "ParentFolderId": {"Id": "parent"},
        "ChildFolderCount": 1,
        "TotalCount": 12,
        "UnreadCount": 3,
    }


def _message(email_id: str = "mail-1") -> dict[str, object]:
    return {
        "ItemId": {"Id": email_id, "ChangeKey": "mail-key"},
        "Subject": "Quarterly update",
        "From": {"Mailbox": {"Name": "Example Sender", "EmailAddress": "sender@example.com"}},
        "DisplayTo": "Mailbox User",
        # Relative to now: search_emails filters out messages older than
        # days_back, so a fixed date turns this fixture into a time bomb.
        "DateTimeReceived": (datetime.now(UTC) - timedelta(days=1)).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "IsRead": False,
        "HasAttachments": True,
        "Preview": "short preview",
        "Body": {"BodyType": "HTML", "Value": "<p>Hello<br>world and more</p>"},
        "ToRecipients": [{"Mailbox": {"Name": "Mailbox User", "EmailAddress": "user@example.com"}}],
        "CcRecipients": [],
    }


class StubAuth:
    settings = Settings(credential_profile="service-test")

    async def ensure_status(self) -> dict[str, object]:
        return {"authenticated": True, "usable": True}


class StubApi:
    def __init__(self, responses: list[dict[str, Any]]) -> None:
        self.responses = responses
        self.calls: list[tuple[str, dict[str, Any], str]] = []

    async def call(
        self,
        action: str,
        body: dict[str, Any],
        *,
        app: str = "Mail",
    ) -> dict[str, Any]:
        self.calls.append((action, body, app))
        return self.responses.pop(0)


def _service(responses: list[dict[str, Any]]) -> tuple[OutlookService, StubApi]:
    api = StubApi(responses)
    service = OutlookService(  # type: ignore[arg-type]
        StubAuth(),
        api,
        Settings(credential_profile="service-test"),
    )
    return service, api


@pytest.mark.asyncio
async def test_search_emails_preserves_com_contract_and_does_not_fetch_body() -> None:
    service, api = _service(
        [
            _success(Folders=[_folder()]),
            _success(
                RootFolder={
                    "Items": [_message()],
                    "IncludesLastItemInRange": True,
                }
            ),
        ]
    )

    result = await service.search_emails(
        query="quarterly",
        days_back=7,
        max_results=5,
        include_body_preview=False,
    )

    rows = result["result"]
    assert isinstance(rows, list)
    assert rows[0]["emailId"] == "mail-1"
    assert rows[0]["storeId"] == STORE_ID
    assert rows[0]["isUnread"] is True
    assert rows[0]["bodyPreview"] is None
    find_body = api.calls[1][1]
    properties = find_body["ItemShape"]["AdditionalProperties"]
    assert not any(item["FieldURI"] == "Preview" for item in properties)
    assert "Restriction" not in find_body
    assert find_body["ViewFilter"] == "All"


@pytest.mark.asyncio
async def test_list_folders_and_get_email_shapes() -> None:
    service, _ = _service([_success(Folders=[_folder()])])
    listed = await service.list_mail_folders(recursive=False)
    row = listed["result"][0]
    assert row == {
        "folderId": "folder-1",
        "storeId": STORE_ID,
        "name": "Inbox",
        "folderPath": "Inbox",
        "parentFolderId": "parent",
        "unreadItemCount": 3,
        "totalItemCount": 12,
        "hasChildren": True,
    }

    service, _ = _service([_success(Items=[_message()])])
    detail = await service.get_email(
        email_id="mail-1",
        store_id=STORE_ID,
        max_body_characters=11,
    )
    assert detail["body"] == "Hello\nworld"
    assert detail["bodyTruncated"] is True
    assert detail["to"] == "Mailbox User <user@example.com>"


@pytest.mark.asyncio
async def test_read_state_update_touches_only_is_read_field() -> None:
    service, api = _service([_success(Items=[_message()]), _success()])

    result = await service.set_email_read_state(
        email_id="mail-1",
        store_id=STORE_ID,
        is_read=True,
    )

    assert result["status"] == "read_state_updated"
    action, payload, _ = api.calls[1]
    assert action == "UpdateItem"
    assert payload["MessageDisposition"] == "SaveOnly"
    update = payload["ItemChanges"][0]["Updates"][0]
    assert update["Path"]["FieldURI"] == "IsRead"
    assert update["Item"] == {"__type": "Message:#Exchange", "IsRead": True}


@pytest.mark.asyncio
async def test_reply_is_saved_as_draft_and_never_sent() -> None:
    draft = _message("draft-1")
    draft["Subject"] = "Re: Quarterly update"
    service, api = _service(
        [
            _success(Items=[_message()]),
            _success(Items=[{"ItemId": {"Id": "draft-1"}}]),
            _success(Items=[draft]),
        ]
    )

    result = await service.create_reply_draft(
        email_id="mail-1",
        store_id=STORE_ID,
        body="Reviewed response",
        reply_all=True,
    )

    assert result["draftId"] == "draft-1"
    assert result["status"] == "saved_to_drafts"
    action, payload, _ = api.calls[1]
    assert action == "CreateItem"
    assert payload["MessageDisposition"] == "SaveOnly"
    assert payload["Items"][0]["__type"] == "ReplyAllToItem:#Exchange"
    assert "Send" not in repr(payload)


@pytest.mark.asyncio
async def test_calendar_uses_calendar_view_and_app() -> None:
    event = {
        "ItemId": {"Id": "event-1"},
        "Subject": "Planning",
        "Start": "2026-08-04T10:00:00+09:00",
        "End": "2026-08-04T10:30:00+09:00",
        "Location": {"DisplayName": "Room A"},
        "Organizer": {"Mailbox": {"EmailAddress": "owner@example.com"}},
        "IsAllDayEvent": False,
        "CalendarItemType": "Single",
        "FreeBusyType": "Busy",
    }
    service, api = _service(
        [
            _success(Folders=[_folder("calendar-id", "Calendar")]),
            {"Body": {"Items": [event]}},
        ]
    )

    result = await service.list_calendar_events(
        starts_after="2026-08-04T00:00:00+09:00",
        ends_before="2026-08-05T00:00:00+09:00",
        max_results=10,
    )

    assert result["result"][0]["eventId"] == "event-1"
    action, payload, app = api.calls[1]
    assert (action, app) == ("GetCalendarView", "Calendar")
    assert payload["CalendarId"]["BaseFolderId"]["Id"] == "calendar-id"
    assert payload["RangeStart"] == "2026-08-03T15:00:00.000"
    assert payload["RangeEnd"] == "2026-08-04T15:00:00.000"


@pytest.mark.asyncio
async def test_index_bodies_stores_text_and_attachment_metadata(tmp_path, monkeypatch) -> None:
    from outlook_web_mcp import index as index_store

    monkeypatch.setattr(Settings, "index_path", property(lambda self: str(tmp_path / "index.db")))
    settings = Settings(credential_profile="bodies-test")
    item = _message("m-42")
    item["Body"] = {"BodyType": "HTML", "Value": "<p>Dear Registry,</p><p>Fees.</p>"}
    item["WebClientReadFormQueryString"] = "?ItemID=abc&viewmodel=ReadMessageItem"
    item["Attachments"] = [{"Name": "offer.pdf", "ContentType": "application/pdf", "Size": 2048}]
    api = StubApi([_success(Items=[item])])
    service = OutlookService(StubAuth(), api, settings)  # type: ignore[arg-type]

    connection = index_store.open_index(settings.index_path)
    index_store.upsert_rows(connection, [{"emailId": "m-42", "subject": "s"}], "inbox")
    connection.close()

    result = await service.index_bodies(max_fetches=10)
    assert result["fetched"] == 1
    assert result["remaining"] == 0

    connection = index_store.open_index(settings.index_path)
    try:
        row = connection.execute(
            "SELECT body_text, web_link FROM items WHERE email_id = 'm-42'"
        ).fetchone()
        assert "Fees." in row["body_text"]
        assert row["web_link"].startswith("https://outlook.cloud.microsoft/owa/")
        attachment = connection.execute(
            "SELECT name, content_type, size FROM attachments WHERE email_id = 'm-42'"
        ).fetchone()
        assert (attachment["name"], attachment["content_type"], attachment["size"]) == (
            "offer.pdf",
            "application/pdf",
            2048,
        )
        assert index_store.search(connection, "Registry", limit=5)
    finally:
        connection.close()
