from __future__ import annotations

import pytest
from test_service import StubApi, StubAuth, _success

from outlook_web_mcp.config import Settings
from outlook_web_mcp.errors import ApiError
from outlook_web_mcp.index import open_index, upsert_rows
from outlook_web_mcp.service import OutlookService


def _service(responses, tmp_path, monkeypatch, profile: str = "authority-test"):
    monkeypatch.setattr(
        Settings, "index_path", property(lambda self: str(tmp_path / "idx" / "index.db"))
    )
    api = StubApi(responses)
    service = OutlookService(  # type: ignore[arg-type]
        StubAuth(), api, Settings(credential_profile=profile)
    )
    return service, api


@pytest.mark.asyncio
async def test_index_hit_for_a_deleted_message_must_fail_loudly(tmp_path, monkeypatch) -> None:
    """The index is a lead, never the truth: a stale id must not resolve from cache."""
    service, _ = _service([_success(Items=[])], tmp_path, monkeypatch)
    connection = open_index(service.settings.index_path)
    upsert_rows(
        connection,
        [
            {
                "emailId": "gone-1",
                "subject": "Deleted last week",
                "senderName": "x",
                "senderAddress": "x@example.edu",
                "receivedAt": "2026-11-01T00:00:00Z",
                "isUnread": False,
                "hasAttachments": False,
                "messageClass": "IPM.Note",
                "internetMessageId": "",
                "references": "",
                "bodyPreview": "cached preview that must not be returned as truth",
            }
        ],
        "inbox",
    )
    connection.close()

    hits = await service.search_index(query="deleted")
    assert [row["emailId"] for row in hits["result"]] == ["gone-1"]

    with pytest.raises(ApiError) as raised:
        await service.get_email(email_id="gone-1", store_id="outlook-web", max_body_characters=100)
    assert raised.value.code == "item_not_found"
    assert "cached preview" not in str(raised.value)


@pytest.mark.asyncio
async def test_search_index_meta_says_engine_and_coverage(tmp_path, monkeypatch) -> None:
    service, _ = _service([], tmp_path, monkeypatch)
    out = await service.search_index(query="x")
    assert out["meta"]["engine"] == "index"
    assert set(out["meta"]) >= {"engine", "rows", "folders", "oldest", "newest", "lastSync"}
