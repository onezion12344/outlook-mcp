from __future__ import annotations

import pytest
from test_service import StubApi, StubAuth, _success

from outlook_web_mcp import pixel, pixel_store
from outlook_web_mcp.approval import approve, draft_digest
from outlook_web_mcp.config import Settings
from outlook_web_mcp.errors import OutlookWebMcpError
from outlook_web_mcp.service import OutlookService

PIXEL_BASE = "https://pixel.example.com"

DRAFT_ITEM = {
    "ItemId": {"Id": "draft-1", "ChangeKey": "k"},
    "Subject": "JUPAS enquiry",
    "ToRecipients": [{"Mailbox": {"Name": "Registry", "EmailAddress": "registry@example.com"}}],
    "CcRecipients": [],
    "Body": {"BodyType": "Text", "Value": "Dear Registry"},
    "DateTimeReceived": "2026-11-05T08:00:00Z",
    "IsRead": True,
    "HasAttachments": False,
}


def _with_html(html: str) -> dict[str, object]:
    item = dict(DRAFT_ITEM)
    item["Body"] = {"BodyType": "HTML", "Value": html}
    return item


def _settings(tmp_path, monkeypatch) -> Settings:
    monkeypatch.setattr(
        Settings, "approval_dir", property(lambda self: str(tmp_path / "approvals"))
    )
    monkeypatch.setattr(Settings, "pixel_dir", property(lambda self: str(tmp_path / "pixels")))
    settings = Settings(credential_profile="send-test")
    # pixel_base is a dataclass field, not a property: override it on the instance.
    object.__setattr__(settings, "pixel_base", None)
    return settings


def _recipientless_draft() -> dict[str, object]:
    """Same DRAFT_ITEM shape but with no recipients at all."""
    item = dict(DRAFT_ITEM)
    item["ToRecipients"] = []
    item["DisplayTo"] = ""
    return item


def _service(responses, tmp_path, monkeypatch):
    api = StubApi(responses)
    service = OutlookService(  # type: ignore[arg-type]
        StubAuth(), api, _settings(tmp_path, monkeypatch)
    )
    return service, api


@pytest.mark.asyncio
async def test_send_refuses_without_approval(tmp_path, monkeypatch) -> None:
    service, api = _service([_success(Items=[DRAFT_ITEM]) for _ in range(2)], tmp_path, monkeypatch)
    with pytest.raises(OutlookWebMcpError) as raised:
        await service.send_draft(draft_id="draft-1", store_id="outlook-web")
    assert raised.value.code == "send_not_approved"
    assert "approve draft-1" in str(raised.value)
    assert "SendItem" not in [call[0] for call in api.calls]


@pytest.mark.asyncio
async def test_send_succeeds_once_with_approval_and_consumes_it(tmp_path, monkeypatch) -> None:
    service, api = _service(
        [
            _success(Items=[DRAFT_ITEM]),  # approval_digest (this test)
            _success(Items=[DRAFT_ITEM]),  # send_draft -> approval_digest
            _success(Items=[DRAFT_ITEM]),  # send_draft -> HTML pixel check
            _success(Items=[DRAFT_ITEM]),  # send_draft -> ChangeKey read
            _success(),  # SendItem
            _success(Folders=[{"FolderId": {"Id": "sent-1"}, "DisplayName": "Sent Items"}]),
            _success(RootFolder={"Items": [], "IncludesLastItemInRange": True}),
            _success(Items=[DRAFT_ITEM]),  # second attempt: approval read
            _success(Items=[DRAFT_ITEM]),  # second attempt: pixel read, then refusal
        ],
        tmp_path,
        monkeypatch,
    )
    info = await service.approval_digest("draft-1")
    approve(service.settings.approval_dir, info["digest"])

    result = await service.send_draft(draft_id="draft-1", store_id="outlook-web")
    assert result["status"] == "sent"
    assert result["draftId"] == "draft-1"
    assert result["tracked"] is False
    action, payload, _ = api.calls[4]
    assert action == "SendItem"
    assert payload["ItemIds"][0]["Id"] == "draft-1"
    assert payload["SaveItemToFolder"] is True
    assert "MessageDisposition" not in payload

    # second attempt must refuse: the token is one-time
    with pytest.raises(OutlookWebMcpError) as raised:
        await service.send_draft(draft_id="draft-1", store_id="outlook-web")
    assert raised.value.code == "send_not_approved"


@pytest.mark.asyncio
async def test_content_change_after_approval_invalidates_it(tmp_path, monkeypatch) -> None:
    service, api = _service([_success(Items=[DRAFT_ITEM]) for _ in range(2)], tmp_path, monkeypatch)
    info = await service.approval_digest("draft-1")
    approve(service.settings.approval_dir, info["digest"])

    edited = dict(DRAFT_ITEM)
    edited["Body"] = {"BodyType": "Text", "Value": "Dear Registry, one more line"}
    api.responses.insert(0, _success(Items=[edited]))
    api.responses.insert(0, _success(Items=[edited]))

    with pytest.raises(OutlookWebMcpError) as raised:
        await service.send_draft(draft_id="draft-1", store_id="outlook-web")
    assert raised.value.code == "send_not_approved"


def test_digest_is_stable_and_recipient_order_insensitive() -> None:
    first = draft_digest(
        draft_id="draft-1",
        subject="Hi",
        recipients="bob@example.com; alice@example.com",
        body="x",
    )
    second = draft_digest(
        draft_id="draft-1",
        subject="Hi",
        recipients="alice@example.com; bob@example.com",
        body="x",
    )
    assert first == second
    assert len(first) == 16


def test_digest_changes_with_body_subject_draft_or_recipient() -> None:
    base = draft_digest(draft_id="d", subject="s", recipients="alice@example.com", body="b")
    assert base != draft_digest(
        draft_id="d2", subject="s", recipients="alice@example.com", body="b"
    )
    assert base != draft_digest(
        draft_id="d", subject="s2", recipients="alice@example.com", body="b"
    )
    assert base != draft_digest(draft_id="d", subject="s", recipients="bob@example.com", body="b")
    assert base != draft_digest(
        draft_id="d", subject="s", recipients="alice@example.com", body="b2"
    )


def _html_item(html: str) -> dict[str, object]:
    item = dict(DRAFT_ITEM)
    item["Body"] = {"BodyType": "HTML", "Value": html}
    return item


@pytest.mark.asyncio
async def test_draft_with_a_foreign_pixel_is_refused(tmp_path, monkeypatch) -> None:
    foreign = pixel.new_token()
    html = f"<p>hi</p>{pixel.build_img(pixel.tracking_url(PIXEL_BASE, foreign), foreign)}"
    service, api = _service(
        [_success(Items=[_html_item(html)]) for _ in range(3)], tmp_path, monkeypatch
    )
    info = await service.approval_digest("draft-1")
    approve(service.settings.approval_dir, info["digest"])

    with pytest.raises(OutlookWebMcpError) as raised:
        await service.send_draft(draft_id="draft-1", store_id="outlook-web")
    assert raised.value.code == "pixel_unknown"
    assert "SendItem" not in [call[0] for call in api.calls]


@pytest.mark.asyncio
async def test_tracked_send_marks_the_token_as_sent(tmp_path, monkeypatch) -> None:
    token = pixel.new_token()
    html = f"<p>hi</p>{pixel.build_img(pixel.tracking_url(PIXEL_BASE, token), token)}"
    service, _ = _service(
        [_success(Items=[_html_item(html)]) for _ in range(4)] + [_success()],
        tmp_path,
        monkeypatch,
    )
    pixel_store.register(
        service.settings.pixel_dir,
        token=token,
        draft_id="draft-1",
        url=pixel.tracking_url(PIXEL_BASE, token),
        at="2026-11-06T10:00:00Z",
    )
    info = await service.approval_digest("draft-1")
    approve(service.settings.approval_dir, info["digest"])

    result = await service.send_draft(draft_id="draft-1", store_id="outlook-web")
    assert result["tracked"] is True
    entry = pixel_store.for_draft(service.settings.pixel_dir, "draft-1")
    assert entry is not None and entry["sentAt"]


@pytest.mark.asyncio
async def test_approval_digest_binds_to_html_not_the_text_rendering(tmp_path, monkeypatch) -> None:
    left = '<p>See <a href="https://example.com/1">here</a></p>'
    right = '<p>See <a href="https://example.com/2">here</a></p>'
    first, _ = _service([_success(Items=[_html_item(left)])], tmp_path, monkeypatch)
    second, _ = _service([_success(Items=[_html_item(right)])], tmp_path, monkeypatch)

    a = await first.approval_digest("draft-1")
    b = await second.approval_digest("draft-1")
    # Same visible text, different destination: the approval must not carry over.
    assert a["digest"] != b["digest"]


@pytest.mark.asyncio
async def test_approval_digest_is_stable_for_an_unchanged_draft(tmp_path, monkeypatch) -> None:
    service, _ = _service(
        [_success(Items=[_html_item("<p>hi</p>")]) for _ in range(2)], tmp_path, monkeypatch
    )
    first = await service.approval_digest("draft-1")
    second = await service.approval_digest("draft-1")
    assert first["digest"] == second["digest"]


@pytest.mark.asyncio
async def test_a_blind_recipient_added_after_approval_invalidates_it(tmp_path, monkeypatch) -> None:
    plain = _html_item("<p>hi</p>")
    with_bcc = _html_item("<p>hi</p>")
    with_bcc["BccRecipients"] = [{"Mailbox": {"EmailAddress": "hidden@example.edu"}}]

    first, _ = _service([_success(Items=[plain])], tmp_path, monkeypatch)
    second, _ = _service([_success(Items=[with_bcc])], tmp_path, monkeypatch)
    a = await first.approval_digest("draft-1")
    b = await second.approval_digest("draft-1")
    assert a["digest"] != b["digest"]
    assert a["bcc"] == ""
    assert "hidden@example.edu" in b["bcc"]


@pytest.mark.asyncio
async def test_an_attachment_added_after_approval_invalidates_it(tmp_path, monkeypatch) -> None:
    plain = _html_item("<p>hi</p>")
    with_file = _html_item("<p>hi</p>")
    with_file["Attachments"] = [{"Name": "transcript.pdf", "Size": 12345}]

    first, _ = _service([_success(Items=[plain])], tmp_path, monkeypatch)
    second, _ = _service([_success(Items=[with_file])], tmp_path, monkeypatch)
    a = await first.approval_digest("draft-1")
    b = await second.approval_digest("draft-1")
    assert a["digest"] != b["digest"]
    assert a["attachments"] == []
    assert b["attachments"] == ["transcript.pdf:12345"]


@pytest.mark.asyncio
async def test_send_refuses_a_recipient_less_draft(tmp_path, monkeypatch) -> None:
    service, api = _service(
        [
            _success(Items=[_recipientless_draft()]),  # approval_digest (this test)
        ],
        tmp_path,
        monkeypatch,
    )
    with pytest.raises(OutlookWebMcpError) as raised:
        await service.send_draft(draft_id="draft-1", store_id="outlook-web")
    assert raised.value.code == "draft_has_no_recipients"
    # Refusal happens before any SendItem is attempted.
    assert all(call[0] != "SendItem" for call in api.calls)


@pytest.mark.asyncio
async def test_senditem_500_falls_back_to_send_and_save_copy(tmp_path, monkeypatch) -> None:
    from outlook_web_mcp.errors import ApiError

    html = "<p>approved body</p>"
    service, api = _service(
        [
            _success(Items=[_with_html(html)]),  # approval_digest (this test)
            _success(Items=[_with_html(html)]),  # send_draft -> approval_digest
            _success(Items=[_with_html(html)]),  # pixel check
            _success(Items=[_with_html(html)]),  # ChangeKey read
            _success(Items=[_with_html(html)]),
            _success(Folders=[{"FolderId": {"Id": "sent-1"}, "DisplayName": "Sent Items"}]),
            _success(RootFolder={"Items": [], "IncludesLastItemInRange": True}),
        ],
        tmp_path,
        monkeypatch,
    )
    # Make SendItem raise the exact live failure: HTTP 500, non-JSON body.
    original_call = api.call

    async def failing_send(action, body, *, app="Mail"):
        if action == "SendItem":
            raise ApiError("non-JSON", "http_500", http_status=500)
        return await original_call(action, body, app=app)

    monkeypatch.setattr(api, "call", failing_send)
    info = await service.approval_digest("draft-1")
    approve(service.settings.approval_dir, info["digest"])
    result = await service.send_draft(draft_id="draft-1", store_id="outlook-web")
    assert result["status"] == "sent"
    fallback = [c for c in api.calls if c[0] == "CreateItem"]
    assert len(fallback) == 1
    payload = fallback[0][1]
    assert payload["MessageDisposition"] == "SendAndSaveCopy"
    assert payload["Items"][0]["Body"]["Value"] == html
