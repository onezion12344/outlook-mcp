from __future__ import annotations

import pytest
from test_service import StubApi, StubAuth, _folder, _message, _success

from outlook_web_mcp.config import Settings
from outlook_web_mcp.service import OutlookService


def _at(email_id: str, stamp: str) -> dict[str, object]:
    item = _message(email_id)
    item["Subject"] = f"subject-{email_id}"
    item["DateTimeReceived"] = stamp
    return item


@pytest.mark.asyncio
async def test_results_are_ordered_by_instant_not_by_raw_string() -> None:
    # 2026-11-05T23:30-05:00 == 2026-11-06T04:30Z  (later)
    # 2026-11-06T09:00+08:00 == 2026-11-06T01:00Z  (earlier)
    # A raw string sort would put the second one first, because "2026-11-06" > "2026-11-05".
    api = StubApi(
        [
            _success(Folders=[_folder()]),
            _success(
                RootFolder={
                    "Items": [
                        _at("local-late", "2026-11-05T23:30:00-05:00"),
                        _at("local-early", "2026-11-06T09:00:00+08:00"),
                    ],
                    "IncludesLastItemInRange": True,
                }
            ),
        ]
    )
    service = OutlookService(  # type: ignore[arg-type]
        StubAuth(), api, Settings(credential_profile="order-test")
    )
    out = await service.search_emails(days_back=30, max_results=10)
    assert [row["emailId"] for row in out["result"]] == ["local-late", "local-early"]


@pytest.mark.asyncio
async def test_unparseable_timestamp_does_not_break_ordering() -> None:
    api = StubApi(
        [
            _success(Folders=[_folder()]),
            _success(
                RootFolder={
                    "Items": [
                        _at("odd", "not-a-date"),
                        _at("normal", "2026-11-06T09:00:00+08:00"),
                    ],
                    "IncludesLastItemInRange": True,
                }
            ),
        ]
    )
    service = OutlookService(  # type: ignore[arg-type]
        StubAuth(), api, Settings(credential_profile="order-odd")
    )
    out = await service.search_emails(days_back=30, max_results=10)
    ids = [row["emailId"] for row in out["result"]]
    assert set(ids) == {"odd", "normal"}
    assert ids[0] == "normal"
