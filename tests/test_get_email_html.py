"""Tests for the get_email_html tool: raw HTML body, extracted links, and text."""

from __future__ import annotations

import pytest
from test_service import STORE_ID, _message, _service, _success

from outlook_web_mcp.errors import OutlookWebMcpError


def _html_message() -> dict[str, object]:
    message = dict(_message())
    message["Body"] = {
        "BodyType": "HTML",
        "Value": (
            '<p>Click <a href="https://accounts.hcaptcha.com/verify_email/abc">here</a>'
            " to log in.</p>"
            '<a href="https://www.hcaptcha.com/accessibility">Docs</a>'
            '<a href="mailto:support@hcaptcha.com">mail</a>'
            '<a href="javascript:void(0)">nope</a>'
        ),
    }
    return message


@pytest.mark.asyncio
async def test_get_email_html_returns_links_and_text() -> None:
    service, api = _service([_success(Items=[_html_message()])])

    result = await service.get_email_html(
        email_id="mail-1",
        store_id=STORE_ID,
    )

    assert result["emailId"] == "mail-1"
    assert result["storeId"] == STORE_ID
    assert result["html"] == _html_message()["Body"]["Value"]
    assert result["links"] == [
        {"href": "https://accounts.hcaptcha.com/verify_email/abc", "text": "here"},
        {"href": "https://www.hcaptcha.com/accessibility", "text": "Docs"},
    ]
    # html_to_text keeps inline anchors inline; the repo's extractor separates
    # consecutive anchors with a newline so a link list stays readable.
    assert result["textContent"] == "Click here to log in.\nDocs\nmail\nnope"
    assert result["bodyTruncated"] is False
    # The whole point: the GetItem call must ask for HTML, not the default Text.
    _, body, _ = api.calls[0]
    assert body["ItemShape"]["BodyType"] == "HTML"


@pytest.mark.asyncio
async def test_get_email_html_truncates_all_text_fields() -> None:
    service, _ = _service([_success(Items=[_html_message()])])

    result = await service.get_email_html(
        email_id="mail-1",
        store_id=STORE_ID,
        max_body_characters=40,
    )

    assert len(result["html"]) == 40
    assert result["bodyTruncated"] is True


@pytest.mark.asyncio
async def test_get_email_html_on_plain_text_message_has_no_links() -> None:
    service, api = _service([_success(Items=[_message()])])
    # Force a plain-text body despite the shared fixture.
    # response_items unwraps one level: the mail lives at Items[0].Items[0].
    api.responses[0]["Body"]["ResponseMessages"]["Items"][0]["Items"][0]["Body"] = {
        "BodyType": "Text",
        "Value": "Hello\r\nworld",
    }

    result = await service.get_email_html(email_id="mail-1", store_id=STORE_ID)

    assert result["links"] == []
    assert result["html"] == "Hello\nworld"
    assert result["textContent"] == "Hello\nworld"


@pytest.mark.asyncio
async def test_get_email_html_rejects_bad_store() -> None:
    service, _ = _service([])

    with pytest.raises(OutlookWebMcpError):
        await service.get_email_html(email_id="mail-1", store_id="not-the-store")


@pytest.mark.asyncio
async def test_get_email_html_rejects_bad_max_body_characters() -> None:
    service, _ = _service([_success(Items=[_html_message()])])

    with pytest.raises(ValueError):
        await service.get_email_html(email_id="mail-1", store_id=STORE_ID, max_body_characters=0)
    with pytest.raises(ValueError):
        await service.get_email_html(
            email_id="mail-1", store_id=STORE_ID, max_body_characters=50_001
        )


@pytest.mark.asyncio
async def test_reply_draft_body_is_html_so_mobile_can_collapse_quotes() -> None:
    # A Text NewBodyContent makes Exchange flatten the quote into <br> runs
    # ("converted from text"), which Outlook mobile cannot fold. The reply body
    # must be HTML.
    service, api = _service(
        [
            _success(Items=[_html_message()]),  # get reply target
            _success(Items=[{"ItemId": {"Id": "reply-1"}}]),  # CreateItem save
            _success(Items=[_html_message()]),  # post-save read-back
        ]
    )
    await service.create_reply_draft(email_id="mail-1", store_id=STORE_ID, body="Thanks, will do.")
    reply_call = next(call for call in api.calls if call[0] == "CreateItem")
    item = reply_call[1]["Items"][0]
    assert item["NewBodyContent"]["BodyType"] == "HTML"
    assert "Thanks, will do." in item["NewBodyContent"]["Value"]
    assert "<" in item["NewBodyContent"]["Value"]  # actually HTML, not raw text
