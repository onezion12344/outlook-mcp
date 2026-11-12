from __future__ import annotations

import pytest
from test_service import StubApi, StubAuth, _folder, _message, _success

from outlook_web_mcp.config import Settings
from outlook_web_mcp.errors import ApiError
from outlook_web_mcp.service import OutlookService


def _service(responses):
    api = StubApi(responses)
    return OutlookService(  # type: ignore[arg-type]
        StubAuth(), api, Settings(credential_profile="draft-test")
    ), api


def _draft_responses(draft_id: str = "draft-9", display_to: str = "registry@example.com"):
    saved = _message(draft_id)
    saved["DisplayTo"] = display_to
    return [
        _success(Folders=[_folder("d-1", "Drafts")]),
        _success(Items=[{"ItemId": {"Id": draft_id}}]),
        _success(Items=[saved]),
    ]


@pytest.mark.asyncio
async def test_create_draft_saves_only_and_never_sends() -> None:
    service, api = _service(_draft_responses())
    result = await service.create_draft(
        to=["registry@example.com"],
        subject="Enquiry",
        body="Dear colleague,\n\nText body.",
    )
    assert result["draftId"] == "draft-9"
    assert result["status"] == "saved_to_drafts"
    assert result["sent"] is False
    assert result["storeId"] == "outlook-web"
    action, payload, _ = api.calls[1]
    assert action == "CreateItem"
    assert payload["MessageDisposition"] == "SaveOnly"
    assert "Send" not in repr(payload)
    item = payload["Items"][0]
    assert item["__type"] == "Message:#Exchange"
    assert item["ToRecipients"] == [
        {
            "Name": "registry@example.com",
            "EmailAddress": "registry@example.com",
            "RoutingType": "SMTP",
            "MailboxType": "Mailbox",
        }
    ]
    assert item["Body"]["BodyType"] == "Text"
    # saved into the Drafts folder resolved by GetFolder, not a hardcoded id
    assert payload["SavedItemFolderId"]["BaseFolderId"]["Id"] == "d-1"


@pytest.mark.asyncio
async def test_cc_is_optional_and_separate() -> None:
    service, api = _service(_draft_responses("draft-cc"))
    await service.create_draft(
        to=["alice@example.com"], cc=["bob@example.com"], subject="S", body="B"
    )
    item = api.calls[1][1]["Items"][0]
    assert item["CcRecipients"] == [
        {
            "Name": "bob@example.com",
            "EmailAddress": "bob@example.com",
            "RoutingType": "SMTP",
            "MailboxType": "Mailbox",
        }
    ]


@pytest.mark.asyncio
async def test_html_body_requested_explicitly() -> None:
    service, api = _service(_draft_responses("draft-10"))
    await service.create_draft(to=["alice@example.com"], subject="S", body="<p>Hi</p>", html=True)
    item = api.calls[1][1]["Items"][0]
    assert item["Body"]["BodyType"] == "HTML"


@pytest.mark.asyncio
async def test_recipient_validation_rejects_junk() -> None:
    service, _ = _service(_draft_responses())
    with pytest.raises(ValueError, match="to"):
        await service.create_draft(to=[], subject="S", body="B")
    with pytest.raises(ValueError, match="address"):
        await service.create_draft(to=["not an address"], subject="S", body="B")
    with pytest.raises(ValueError, match="address"):
        await service.create_draft(to=["alice@example.com", "bad;two@x"], subject="S", body="B")


@pytest.mark.asyncio
async def test_subject_and_body_limits() -> None:
    service, _ = _service(_draft_responses())
    with pytest.raises(ValueError, match="subject"):
        await service.create_draft(to=["alice@example.com"], subject="  ", body="B")
    with pytest.raises(ValueError, match="body"):
        await service.create_draft(to=["alice@example.com"], subject="S", body="   ")
    with pytest.raises(ValueError, match="20000"):
        await service.create_draft(to=["alice@example.com"], subject="S", body="x" * 20001)


@pytest.mark.asyncio
async def test_missing_draft_id_raises_instead_of_claiming_success() -> None:
    service, _ = _service(
        [
            _success(Folders=[_folder("d-1", "Drafts")]),
            _success(Items=[]),
        ]
    )
    with pytest.raises(Exception, match="draft"):
        await service.create_draft(to=["alice@example.com"], subject="S", body="B")


def _settings_with(monkeypatch, tmp_path, *, base: str | None):
    monkeypatch.setattr(Settings, "pixel_dir", property(lambda self: str(tmp_path / "pixels")))
    settings = Settings(credential_profile="draft-track")
    object.__setattr__(settings, "pixel_base", base)
    return settings


@pytest.mark.asyncio
async def test_tracked_html_draft_embeds_a_pixel_and_registers_it(tmp_path, monkeypatch) -> None:
    from test_service import StubApi as Api

    from outlook_web_mcp import pixel, pixel_store
    from outlook_web_mcp.service import OutlookService as Service

    settings = _settings_with(monkeypatch, tmp_path, base="https://pixel.example.com")
    api = Api(_draft_responses("draft-trk"))
    service = Service(StubAuth(), api, settings)  # type: ignore[arg-type]

    result = await service.create_draft(
        to=["alice@example.com"], subject="S", body="<p>hi</p>", html=True, track=True
    )
    assert result["tracked"] is True
    sent_html = api.calls[1][1]["Items"][0]["Body"]["Value"]
    token = pixel.find_token(sent_html)
    assert token is not None
    assert pixel_store.known(settings.pixel_dir, token)
    assert pixel_store.for_draft(settings.pixel_dir, "draft-trk")["token"] == token


@pytest.mark.asyncio
async def test_tracking_without_a_pixel_base_warns_instead_of_lying(tmp_path, monkeypatch) -> None:
    from test_service import StubApi as Api

    from outlook_web_mcp.service import OutlookService as Service

    settings = _settings_with(monkeypatch, tmp_path, base=None)
    api = Api(_draft_responses("draft-nobase"))
    service = Service(StubAuth(), api, settings)  # type: ignore[arg-type]

    result = await service.create_draft(
        to=["alice@example.com"], subject="S", body="<p>hi</p>", html=True, track=True
    )
    assert result["tracked"] is False
    assert "OUTLOOK_WEB_PIXEL_BASE" in result["trackingWarning"]
    assert "<img" not in api.calls[1][1]["Items"][0]["Body"]["Value"]


@pytest.mark.asyncio
async def test_plain_text_draft_still_gets_a_pixel_by_default(tmp_path, monkeypatch) -> None:
    from test_service import StubApi as Api

    from outlook_web_mcp import pixel
    from outlook_web_mcp.service import OutlookService as Service

    settings = _settings_with(monkeypatch, tmp_path, base="https://pixel.example.com")
    api = Api(_draft_responses("draft-text"))
    service = Service(StubAuth(), api, settings)  # type: ignore[arg-type]

    result = await service.create_draft(
        to=["alice@example.com"], subject="S", body="plain body\nsecond & <b>"
    )
    assert result["status"] == "saved_to_drafts"
    assert result["tracked"] is True
    sent = api.calls[1][1]["Items"][0]
    assert sent["Body"]["BodyType"] == "HTML"
    value = sent["Body"]["Value"]
    assert "plain body" in value and "<br>" in value
    assert "&amp;" in value and "&lt;b&gt;" in value
    assert pixel.find_token(value) is not None


@pytest.mark.asyncio
async def test_without_a_pixel_base_a_text_draft_stays_text(tmp_path, monkeypatch) -> None:
    from test_service import StubApi as Api

    from outlook_web_mcp.service import OutlookService as Service

    settings = _settings_with(monkeypatch, tmp_path, base=None)
    api = Api(_draft_responses("draft-plain"))
    service = Service(StubAuth(), api, settings)  # type: ignore[arg-type]

    result = await service.create_draft(to=["alice@example.com"], subject="S", body="plain body")
    assert result["tracked"] is False
    assert "OUTLOOK_WEB_PIXEL_BASE" in result["trackingWarning"]
    assert api.calls[1][1]["Items"][0]["Body"]["BodyType"] == "Text"


@pytest.mark.asyncio
async def test_explicit_track_without_html_is_an_error(tmp_path, monkeypatch) -> None:
    from test_service import StubApi as Api

    from outlook_web_mcp.service import OutlookService as Service

    settings = _settings_with(monkeypatch, tmp_path, base="https://pixel.example.com")
    api = Api(_draft_responses())
    service = Service(StubAuth(), api, settings)  # type: ignore[arg-type]

    with pytest.raises(ValueError, match="html=True"):
        await service.create_draft(
            to=["alice@example.com"], subject="S", body="plain", html=False, track=True
        )


@pytest.mark.asyncio
async def test_draft_without_recipients_on_save_fails_loud() -> None:
    # If OWA ever drops recipients again, saving must fail, not return success.
    service, _ = _service(_draft_responses("draft-quiet", display_to=""))
    with pytest.raises(ApiError) as raised:
        await service.create_draft(to=["alice@example.com"], subject="S", body="B")
    assert raised.value.code == "recipients_dropped"


@pytest.mark.asyncio
async def test_recipients_are_flat_entries_not_ews_wrappers() -> None:
    # The {Mailbox: {...}} wrapper is silently dropped by this dialect (PITFALLS #23).
    service, api = _service(_draft_responses())
    await service.create_draft(to=["alice@example.com"], subject="S", body="B")
    item = api.calls[1][1]["Items"][0]
    recipient = item["ToRecipients"][0]
    assert "Mailbox" not in recipient
    assert recipient["EmailAddress"] == "alice@example.com"
    assert recipient["RoutingType"] == "SMTP"
