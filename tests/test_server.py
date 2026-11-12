from __future__ import annotations

import json

import pytest

from outlook_web_mcp.server import mcp


@pytest.mark.asyncio
async def test_expected_tools_and_annotations_are_registered() -> None:
    tools = await mcp.list_tools()
    by_name = {tool.name: tool for tool in tools}
    assert set(by_name) == {
        "outlook_auth_status",
        "search_emails",
        "list_mail_folders",
        "get_email",
        "get_email_html",
        "set_email_read_state",
        "list_calendar_events",
        "create_reply_draft",
        "create_draft",
        "send_draft",
        "track_status",
        "index_sync",
        "search_index",
        "index_status",
        "index_verify",
        "index_bodies",
        "index_attachment_text",
        "pixels_sync",
    }
    assert not [name for name in by_name if name.casefold() == "send_email"]
    assert by_name["search_emails"].annotations.readOnlyHint is True
    assert by_name["get_email"].annotations.readOnlyHint is True
    assert by_name["set_email_read_state"].annotations.readOnlyHint is False
    assert by_name["create_reply_draft"].annotations.readOnlyHint is False
    assert by_name["create_reply_draft"].annotations.idempotentHint is False
    assert by_name["index_bodies"].annotations.readOnlyHint is False
    assert by_name["create_draft"].annotations.readOnlyHint is False
    assert by_name["create_draft"].annotations.destructiveHint is False
    assert by_name["create_draft"].annotations.idempotentHint is False
    assert by_name["send_draft"].annotations.readOnlyHint is False
    assert by_name["send_draft"].annotations.openWorldHint is True
    assert by_name["send_draft"].annotations.idempotentHint is False
    assert by_name["search_index"].annotations.readOnlyHint is True
    assert by_name["index_status"].annotations.readOnlyHint is True
    assert by_name["index_sync"].annotations.readOnlyHint is False
    assert by_name["index_sync"].annotations.openWorldHint is False
    assert by_name["pixels_sync"].annotations.readOnlyHint is False
    assert by_name["pixels_sync"].annotations.openWorldHint is True

    search_properties = by_name["search_emails"].inputSchema["properties"]
    assert "daysBack" in search_properties
    assert "maxResults" in search_properties
    assert "includeBodyPreview" in search_properties
    assert "includeSubfolders" in search_properties
    assert "offset" in search_properties
    assert "continuationFolderId" in search_properties
    assert "continuationOffset" in search_properties


@pytest.mark.asyncio
async def test_auth_status_tool_hides_tenant_and_display_name(monkeypatch) -> None:
    from outlook_web_mcp import server

    async def fake_status() -> dict[str, object]:
        return {
            "authenticated": True,
            "profile": "default",
            "tenant_id": "00000000-0000-0000-0000-000000000000",
            "display_name": "Owner Real Name",
            "usable": True,
        }

    monkeypatch.setattr(server.service, "auth_status", fake_status)
    result = await server.mcp.call_tool("outlook_auth_status", {})
    payload = result[1] if isinstance(result, tuple) else result
    text = json.dumps(payload, default=str)
    assert "tenant_id" not in text and "display_name" not in text
    assert "00000000" not in text and "Owner Real Name" not in text
    assert "authenticated" in text
