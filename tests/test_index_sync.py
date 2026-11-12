from __future__ import annotations

import pytest
from test_service import StubApi, StubAuth, _folder, _message, _success

from outlook_web_mcp.config import Settings
from outlook_web_mcp.service import OutlookService


def _service(responses, tmp_path, monkeypatch, profile: str = "index-test"):
    monkeypatch.setattr(
        Settings, "index_path", property(lambda self: str(tmp_path / "idx" / "index.db"))
    )
    api = StubApi(responses)
    service = OutlookService(  # type: ignore[arg-type]
        StubAuth(), api, Settings(credential_profile=profile)
    )
    return service, api


def _hit(email_id: str, subject: str, preview: str = "") -> dict[str, object]:
    item = _message(email_id)
    item["Subject"] = subject
    item["Preview"] = preview
    return item


@pytest.mark.asyncio
async def test_index_sync_stores_rows_and_reports_counts(tmp_path, monkeypatch) -> None:
    service, _ = _service(
        [
            _success(Folders=[_folder("in-1", "Inbox")]),
            _success(
                RootFolder={"Items": [_hit("m1", "Weekly digest")], "IncludesLastItemInRange": True}
            ),
        ],
        tmp_path,
        monkeypatch,
    )
    out = await service.index_sync(folders=["inbox"], days_back=30)
    assert out["indexed"] == 1
    assert out["truncated"] is False
    assert out["folders"] == ["inbox"]

    info = await service.index_status()
    assert info["rows"] == 1
    assert info["folders"] == 1


@pytest.mark.asyncio
async def test_index_sync_follows_continuation_until_done(tmp_path, monkeypatch) -> None:
    service, api = _service(
        [
            _success(Folders=[_folder("in-1", "Inbox")]),
            _success(
                RootFolder={
                    "Items": [_hit("m1", "one")],
                    "IncludesLastItemInRange": False,
                    "IndexedPagingOffset": 1,
                }
            ),
            _success(
                RootFolder={
                    "Items": [_hit("m2", "two")],
                    "IncludesLastItemInRange": True,
                }
            ),
        ],
        tmp_path,
        monkeypatch,
    )
    out = await service.index_sync(folders=["inbox"], days_back=30)
    assert out["indexed"] == 2
    assert out["truncated"] is False
    find_items = [call for call in api.calls if call[0] == "FindItem"]
    assert len(find_items) == 2
    assert find_items[1][1]["Paging"]["Offset"] == 1


@pytest.mark.asyncio
async def test_index_sync_pages_by_offset_when_hit_cap_reached(tmp_path, monkeypatch) -> None:
    service, api = _service(
        [
            _success(Folders=[_folder("in-1", "Inbox")]),
            _success(
                RootFolder={
                    "Items": [_hit("m1", "a"), _hit("m2", "b"), _hit("m3", "c")],
                    "IncludesLastItemInRange": True,
                }
            ),
            _success(Folders=[_folder("in-1", "Inbox")]),
            _success(
                RootFolder={
                    "Items": [_hit("m3", "c")],
                    "IncludesLastItemInRange": True,
                }
            ),
        ],
        tmp_path,
        monkeypatch,
    )
    out = await service.index_sync(folders=["inbox"], days_back=30, page_size=2)
    assert out["indexed"] == 3
    find_items = [call for call in api.calls if call[0] == "FindItem"]
    assert len(find_items) == 2
    # The second page advances the server cursor (meta.nextPage, = Offset 0 + 3 rows
    # the stub returned) instead of re-scanning this window from Offset 0.
    assert find_items[1][1]["Paging"]["Offset"] == 3


@pytest.mark.asyncio
async def test_search_index_hits_body_preview(tmp_path, monkeypatch) -> None:
    service, _ = _service(
        [
            _success(Folders=[_folder("in-1", "Inbox")]),
            _success(
                RootFolder={
                    "Items": [_hit("m1", "Weekly digest", "scholarship deadline moved")],
                    "IncludesLastItemInRange": True,
                }
            ),
        ],
        tmp_path,
        monkeypatch,
    )
    await service.index_sync(folders=["inbox"], days_back=30)
    out = await service.search_index(query="scholarship")
    assert out["meta"]["engine"] == "index"
    assert [row["emailId"] for row in out["result"]] == ["m1"]


@pytest.mark.asyncio
async def test_search_index_on_empty_index_is_empty_not_an_error(tmp_path, monkeypatch) -> None:
    service, _ = _service([], tmp_path, monkeypatch)
    out = await service.search_index(query="anything")
    assert out["result"] == []
    assert out["meta"]["rows"] == 0


@pytest.mark.asyncio
async def test_second_sync_prunes_messages_deleted_from_the_mailbox(tmp_path, monkeypatch) -> None:
    first, _ = _service(
        [
            _success(Folders=[_folder("in-1", "Inbox")]),
            _success(
                RootFolder={
                    "Items": [_hit("m1", "one"), _hit("m2", "two")],
                    "IncludesLastItemInRange": True,
                }
            ),
        ],
        tmp_path,
        monkeypatch,
        "prune-first",
    )
    out = await first.index_sync(folders=["inbox"], days_back=30)
    assert out["indexed"] == 2
    assert out["pruned"] == 0

    # The mailbox no longer has m2 (deleted or moved).
    second, _ = _service(
        [
            _success(Folders=[_folder("in-1", "Inbox")]),
            _success(
                RootFolder={
                    "Items": [_hit("m1", "one")],
                    "IncludesLastItemInRange": True,
                }
            ),
        ],
        tmp_path,
        monkeypatch,
        "prune-second",
    )
    out2 = await second.index_sync(folders=["inbox"], days_back=30)
    assert out2["pruned"] == 1
    info = await second.index_status()
    assert info["rows"] == 1


@pytest.mark.asyncio
async def test_truncated_sync_never_prunes(tmp_path, monkeypatch) -> None:
    first, _ = _service(
        [
            _success(Folders=[_folder("in-1", "Inbox")]),
            _success(
                RootFolder={
                    "Items": [_hit("m1", "one"), _hit("m2", "two")],
                    "IncludesLastItemInRange": True,
                }
            ),
        ],
        tmp_path,
        monkeypatch,
        "prune-full",
    )
    await first.index_sync(folders=["inbox"], days_back=30)

    capped, _ = _service(
        [
            _success(Folders=[_folder("in-1", "Inbox")]),
            _success(
                RootFolder={
                    "Items": [_hit("m1", "one")],
                    "IncludesLastItemInRange": True,
                }
            ),
        ],
        tmp_path,
        monkeypatch,
        "prune-capped",
    )
    out = await capped.index_sync(folders=["inbox"], days_back=30, max_rows=1)
    assert out["truncated"] is True
    assert out["pruned"] == 0
    assert (await capped.index_status())["rows"] == 2


@pytest.mark.asyncio
async def test_default_sync_covers_every_mail_folder(tmp_path, monkeypatch) -> None:
    service, _ = _service([], tmp_path, monkeypatch, "default-folders")

    async def targets():
        return [
            ("inbox", "in-9", 1),
            ("sent", "s-9", 1),
            ("archive", "a-9", 1),
            ("deleteditems", "d-9", 1),
            ("drafts", "dr-9", 1),
            ("junk", "j-9", 1),
        ]

    seen: list[tuple[str, str | None]] = []

    async def search(**kwargs):
        seen.append((kwargs.get("folder"), kwargs.get("folder_id")))
        return {
            "result": [],
            "meta": {"scannedItems": 0, "hitCapReached": False, "truncated": False},
        }

    monkeypatch.setattr(service, "_mail_folder_targets", targets)
    monkeypatch.setattr(service, "search_emails", search)
    out = await service.index_sync()
    # Drafts and Junk are included now, not just the original four.
    assert out["folders"] == ["inbox", "sent", "archive", "deleteditems", "drafts", "junk"]
    assert [folder_id for _, folder_id in seen] == ["in-9", "s-9", "a-9", "d-9", "dr-9", "j-9"]
    assert out["truncated"] is False


@pytest.mark.asyncio
async def test_mail_folder_targets_exclude_calendar_subtrees(tmp_path, monkeypatch) -> None:
    service, _ = _service([], tmp_path, monkeypatch, "targets")

    async def listing(**kwargs):
        return {
            "result": [
                {
                    "name": "Inbox",
                    "folderId": "in-1",
                    "totalItemCount": 10,
                    "folderPath": "Top of Information Store\\Inbox",
                },
                {
                    "name": "Drafts",
                    "folderId": "dr-1",
                    "totalItemCount": 3,
                    "folderPath": "Top of Information Store\\Drafts",
                },
                {
                    "name": "Junk Email",
                    "folderId": "j-1",
                    "totalItemCount": 2,
                    "folderPath": "Top of Information Store\\Junk Email",
                },
                {
                    "name": "note",
                    "folderId": "n-1",
                    "totalItemCount": 3,
                    "folderPath": "Top of Information Store\\note",
                },
                {
                    "name": "Calendar",
                    "folderId": "c-1",
                    "totalItemCount": 43,
                    "folderPath": "Top of Information Store\\Calendar",
                },
                {
                    "name": "中国 节假日",
                    "folderId": "h-1",
                    "totalItemCount": 208,
                    "folderPath": "Top of Information Store\\Calendar\\中国 节假日",
                },
                {
                    "name": "Empty",
                    "folderId": "e-1",
                    "totalItemCount": 0,
                    "folderPath": "Top of Information Store\\Empty",
                },
            ]
        }

    monkeypatch.setattr(service, "list_mail_folders", listing)
    got = await service._mail_folder_targets()
    assert [label for label, _, _ in got] == ["inbox", "drafts", "junk", "note"]
    assert all(count > 0 for _, _, count in got)


@pytest.mark.asyncio
async def test_index_verify_reports_the_gap_per_folder(tmp_path, monkeypatch) -> None:
    service, _ = _service(
        [
            _success(Folders=[_folder("in-1", "Inbox")]),
            _success(RootFolder={"Items": [_hit("m1", "one")], "IncludesLastItemInRange": True}),
        ],
        tmp_path,
        monkeypatch,
        "verify-gap",
    )

    async def targets():
        return [
            ("inbox", "in-1", 12),
            ("sent", "se-1", 12),
            ("archive", "ar-1", 12),
            ("deleteditems", "de-1", 12),
        ]

    monkeypatch.setattr(service, "_mail_folder_targets", targets)
    await service.index_sync(folders=["inbox"], days_back=3650)
    report = await service.index_verify()
    by_folder = {item["folder"]: item for item in report["result"]}
    row = by_folder["inbox"]
    assert row["indexed"] == 1
    assert row["mailbox"] == 12  # the _folder() fixture reports TotalCount=12
    assert row["gap"] == 11
    assert row["complete"] is False
    # Folders with no indexed rows at all are reported too, not silently omitted.
    assert by_folder["sent"]["indexed"] == 0
    assert by_folder["sent"]["complete"] is False
    assert report["complete"] is False
