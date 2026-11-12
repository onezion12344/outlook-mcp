from __future__ import annotations

import pytest
from test_service import StubApi, StubAuth, _folder, _message, _success

from outlook_web_mcp.config import Settings
from outlook_web_mcp.service import OutlookService


def _child(folder_id: str, name: str, folder_class: str = "IPF.Note") -> dict[str, object]:
    return {
        "FolderId": {"Id": folder_id, "ChangeKey": "k"},
        "DisplayName": name,
        "FolderClass": folder_class,
        "ParentFolderId": {"Id": "f-1"},
        "ChildFolderCount": 0,
        "TotalCount": 1,
        "UnreadCount": 0,
    }


def _service(responses: list[dict[str, object]], profile: str):
    api = StubApi(responses)
    return OutlookService(  # type: ignore[arg-type]
        StubAuth(), api, Settings(credential_profile=profile)
    ), api


@pytest.mark.asyncio
async def test_include_subfolders_scans_note_children_only() -> None:
    service, api = _service(
        [
            _success(Folders=[_folder("f-1", "Inbox")]),
            _success(
                RootFolder={
                    "Folders": [
                        _child("f-2", "Work"),
                        _child("f-3", "Appointments", "IPF.Appointment"),
                    ]
                }
            ),
            _success(RootFolder={"Folders": []}),
            _success(RootFolder={"Items": [_message("m1")], "IncludesLastItemInRange": True}),
            _success(RootFolder={"Items": [_message("m2")], "IncludesLastItemInRange": True}),
        ],
        "sub-on",
    )
    out = await service.search_emails(days_back=30, max_results=50, include_subfolders=True)

    assert "FindFolder" in [call[0] for call in api.calls]
    find_item_targets = [
        call[1]["ParentFolderIds"][0]["Id"] for call in api.calls if call[0] == "FindItem"
    ]
    assert find_item_targets == ["f-1", "f-2"]
    assert out["meta"]["scannedFolders"] == 2
    assert out["meta"]["scannedItems"] == 2


@pytest.mark.asyncio
async def test_include_subfolders_false_does_not_recurse() -> None:
    service, api = _service(
        [
            _success(Folders=[_folder("f-1", "Inbox")]),
            _success(RootFolder={"Items": [_message("m1")], "IncludesLastItemInRange": True}),
        ],
        "sub-off",
    )
    out = await service.search_emails(days_back=30, max_results=50, include_subfolders=False)
    assert "FindFolder" not in [call[0] for call in api.calls]
    assert out["meta"]["scannedFolders"] == 1
