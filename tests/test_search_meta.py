from __future__ import annotations

from typing import Any

import pytest
from test_service import StubApi, StubAuth, _folder, _message, _success

from outlook_web_mcp.config import Settings
from outlook_web_mcp.service import OutlookService


def _service(responses: list[dict[str, Any]], **settings_kw: Any):
    api = StubApi(responses)
    service = OutlookService(  # type: ignore[arg-type]
        StubAuth(),
        api,
        Settings(credential_profile="meta-test", **settings_kw),
    )
    return service, api


@pytest.mark.asyncio
async def test_meta_reports_coverage_and_no_truncation_when_server_finishes() -> None:
    service, _ = _service(
        [
            _success(Folders=[_folder()]),
            _success(
                RootFolder={
                    "Items": [_message("m1"), _message("m2")],
                    "IncludesLastItemInRange": True,
                }
            ),
        ]
    )
    out = await service.search_emails(days_back=30, max_results=50)
    meta = out["meta"]
    assert meta["engine"] == "local"
    assert meta["scannedItems"] == 2
    assert meta["scannedFolders"] == 1
    assert meta["matchCount"] == 2
    assert meta["returnedCount"] == 2
    assert meta["truncated"] is False
    assert meta["truncatedReason"] is None
    assert meta["hitCapReached"] is False
    assert set(meta["coverage"]) == {"from", "to"}


@pytest.mark.asyncio
async def test_meta_flags_scan_cap_as_truncated() -> None:
    service, _ = _service(
        [
            _success(Folders=[_folder()]),
            _success(
                RootFolder={
                    "Items": [_message("m1"), _message("m2")],
                    "IncludesLastItemInRange": False,
                    "IndexedPagingOffset": 2,
                }
            ),
        ],
        max_scanned_items=2,
    )
    out = await service.search_emails(days_back=30, max_results=50)
    meta = out["meta"]
    assert meta["truncated"] is True
    assert meta["truncatedReason"] == "scan_cap"


@pytest.mark.asyncio
async def test_exactly_a_full_page_is_not_reported_as_complete() -> None:
    """found == returned == maxResults, yet the server still has more: must flag it."""
    service, _ = _service(
        [
            _success(Folders=[_folder()]),
            _success(
                RootFolder={
                    "Items": [_message("m1"), _message("m2")],
                    "IncludesLastItemInRange": False,
                    "IndexedPagingOffset": 2,
                }
            ),
        ]
    )
    out = await service.search_emails(days_back=30, max_results=2)
    meta = out["meta"]
    assert meta["returnedCount"] == 2
    assert meta["matchCount"] == 2
    assert meta["hitCapReached"] is True
    assert meta["nextPage"] == {"folderId": meta["nextPage"]["folderId"], "offset": 2}


@pytest.mark.asyncio
async def test_meta_flags_hit_cap_when_more_matches_than_returned() -> None:
    service, _ = _service(
        [
            _success(Folders=[_folder()]),
            _success(
                RootFolder={
                    "Items": [_message("m1"), _message("m2"), _message("m3")],
                    "IncludesLastItemInRange": True,
                }
            ),
        ]
    )
    out = await service.search_emails(days_back=30, max_results=2)
    meta = out["meta"]
    assert meta["matchCount"] == 3
    assert meta["returnedCount"] == 2
    assert meta["hitCapReached"] is True
