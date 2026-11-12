from __future__ import annotations

import pytest
from test_service import StubApi, StubAuth, _folder, _message, _success

from outlook_web_mcp.config import Settings
from outlook_web_mcp.service import OutlookService


def _service(responses):
    api = StubApi(responses)
    return OutlookService(  # type: ignore[arg-type]
        StubAuth(), api, Settings(credential_profile="trace-test")
    ), api


def _report_item() -> dict[str, object]:
    """Shape a real search response would have: no MessageClass (FindItem rejects it)."""
    item = _message("r-1")
    item["InternetMessageId"] = "<draft-1@UniversityPR01000CBAAP0.example>"
    item["References"] = "<orig-2@example.com>"
    return item


def _detail_item() -> dict[str, object]:
    """Shape a GetItem(AllProperties) response has: MessageClass included."""
    item = _report_item()
    item["MessageClass"] = "REPORT.IPM.Note.IPNRN"
    return item


@pytest.mark.asyncio
async def test_search_rows_carry_trace_fields() -> None:
    service, api = _service(
        [
            _success(Folders=[_folder()]),
            _success(RootFolder={"Items": [_report_item()], "IncludesLastItemInRange": True}),
        ]
    )
    out = await service.search_emails(days_back=30, max_results=10)
    row = out["result"][0]
    assert row["internetMessageId"] == "<draft-1@UniversityPR01000CBAAP0.example>"
    assert row["references"] == "<orig-2@example.com>"
    # Exchange rejects MessageClass inside FindItem (BadRequest: Invalid argument),
    # so search rows always carry "" here; get_email returns the real value.
    assert row["messageClass"] == ""
    find_body = api.calls[1][1]
    uris = {prop["FieldURI"] for prop in find_body["ItemShape"]["AdditionalProperties"]}
    assert {"InternetMessageId", "References"} <= uris
    assert "MessageClass" not in uris


@pytest.mark.asyncio
async def test_missing_fields_are_empty_not_none() -> None:
    service, _ = _service(
        [
            _success(Folders=[_folder()]),
            _success(RootFolder={"Items": [_message("plain-1")], "IncludesLastItemInRange": True}),
        ]
    )
    out = await service.search_emails(days_back=30, max_results=10)
    row = out["result"][0]
    assert row["messageClass"] == ""
    assert row["internetMessageId"] == ""
    assert row["references"] == ""


@pytest.mark.asyncio
async def test_get_email_exposes_trace_fields() -> None:
    service, _ = _service([_success(Items=[_detail_item()])])
    detail = await service.get_email(email_id="r-1", store_id="outlook-web", max_body_characters=50)
    assert detail["messageClass"] == "REPORT.IPM.Note.IPNRN"
    assert detail["references"] == "<orig-2@example.com>"
