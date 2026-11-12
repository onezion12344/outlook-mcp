from __future__ import annotations

from typing import Any

import pytest
from test_service import StubApi, StubAuth, _folder, _message, _success

from outlook_web_mcp.config import Settings
from outlook_web_mcp.service import OutlookService


def _service(responses: list[dict[str, Any]], **settings_kw: Any):
    api = StubApi(responses)
    service = OutlookService(  # type: ignore[arg-type]
        StubAuth(), api, Settings(credential_profile="page-test", **settings_kw)
    )
    return service, api


@pytest.mark.asyncio
async def test_offset_skips_matches_without_losing_coverage() -> None:
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
    out = await service.search_emails(days_back=30, max_results=2, offset=1)
    meta = out["meta"]
    assert len(out["result"]) == 2
    assert meta["matchCount"] == 3
    assert meta["offset"] == 1
    assert meta["returnedCount"] == 2
    # 1 skipped + 2 returned == 3 matches: nothing left behind, so no hit cap.
    assert meta["hitCapReached"] is False


@pytest.mark.asyncio
async def test_hit_cap_when_matches_remain_unreturned() -> None:
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


@pytest.mark.asyncio
async def test_continuation_is_returned_when_scan_cap_hits() -> None:
    service, _ = _service(
        [
            _success(Folders=[_folder("f-1", "Inbox")]),
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
    assert meta["continuation"] == {"folderId": "f-1", "offset": 2}


@pytest.mark.asyncio
async def test_continuation_resumes_from_the_given_cursor() -> None:
    service, api = _service(
        [
            _success(Folders=[_folder("f-1", "Inbox")]),
            _success(RootFolder={"Items": [_message("m3")], "IncludesLastItemInRange": True}),
        ],
        max_scanned_items=2,
    )
    out = await service.search_emails(
        days_back=30,
        max_results=50,
        continuation_folder_id="f-1",
        continuation_offset=2,
    )
    meta = out["meta"]
    assert meta["truncated"] is False
    assert meta["continuation"] is None
    assert api.calls[1][1]["Paging"]["Offset"] == 2


@pytest.mark.asyncio
async def test_continuation_folder_outside_the_set_is_rejected() -> None:
    service, _ = _service([_success(Folders=[_folder("f-1", "Inbox")])])
    with pytest.raises(ValueError, match="continuation"):
        await service.search_emails(days_back=30, max_results=10, continuation_folder_id="ghost")


@pytest.mark.asyncio
async def test_max_results_cap_raised_to_200() -> None:
    service, _ = _service([_success(Folders=[_folder()])])
    with pytest.raises(ValueError, match="1 and 200"):
        await service.search_emails(days_back=30, max_results=201)


@pytest.mark.asyncio
async def test_negative_offset_rejected() -> None:
    service, _ = _service([_success(Folders=[_folder()])])
    with pytest.raises(ValueError, match="offset"):
        await service.search_emails(days_back=30, max_results=10, offset=-1)


def test_default_scan_cap_is_twenty_thousand() -> None:
    import os

    os.environ.pop("OUTLOOK_WEB_MAX_SCANNED_ITEMS", None)
    assert Settings(credential_profile="cap-test").max_scanned_items == 20_000


@pytest.mark.asyncio
async def test_history_window_extends_past_one_year() -> None:
    service, _ = _service(
        [
            _success(Folders=[_folder()]),
            _success(RootFolder={"Items": [], "IncludesLastItemInRange": True}),
        ]
    )
    out = await service.search_emails(days_back=3650, max_results=10)
    assert out["result"] == []
    assert out["meta"]["coverage"]["from"] < out["meta"]["coverage"]["to"]


@pytest.mark.asyncio
async def test_absurd_window_is_rejected() -> None:
    service, _ = _service([_success(Folders=[_folder()])])
    with pytest.raises(ValueError, match="1 and 3650"):
        await service.search_emails(days_back=3651, max_results=10)
