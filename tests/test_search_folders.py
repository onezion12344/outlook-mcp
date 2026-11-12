from __future__ import annotations

import pytest
from test_service import StubApi, StubAuth, _folder, _success

from outlook_web_mcp.config import Settings
from outlook_web_mcp.service import DISTINGUISHED_FOLDERS, OutlookService


def _service(responses: list[dict[str, object]]):
    api = StubApi(responses)
    return OutlookService(  # type: ignore[arg-type]
        StubAuth(), api, Settings(credential_profile="folder-test")
    ), api


@pytest.mark.asyncio
async def test_archive_resolves_to_distinguished_archive_id() -> None:
    service, api = _service([_success(Folders=[_folder("a-1", "Archive")])])
    await service.list_mail_folders(folder="archive", recursive=False)
    assert api.calls[0][1]["FolderIds"] == [
        {"__type": "DistinguishedFolderId:#Exchange", "Id": "archive"}
    ]


@pytest.mark.asyncio
async def test_root_resolves_to_msgfolderroot() -> None:
    service, api = _service([_success(Folders=[_folder("root-1", "Root")])])
    await service.list_mail_folders(folder="root", recursive=False)
    assert api.calls[0][1]["FolderIds"][0]["Id"] == "msgfolderroot"


@pytest.mark.asyncio
async def test_deleted_junk_and_outbox_resolve() -> None:
    cases = [
        ("deleteditems", "deleteditems"),
        ("deleted_mail", "deleteditems"),
        ("junk", "junkemail"),
        ("junkemail", "junkemail"),
        ("outbox", "outbox"),
    ]
    for name, expected in cases:
        service, api = _service([_success(Folders=[_folder("x-1", name)])])
        await service.list_mail_folders(folder=name, recursive=False)
        assert api.calls[0][1]["FolderIds"][0]["Id"] == expected, name


@pytest.mark.asyncio
async def test_unknown_folder_error_lists_every_supported_name() -> None:
    service, _ = _service([])
    with pytest.raises(ValueError) as raised:
        await service.list_mail_folders(folder="nope", recursive=False)
    message = str(raised.value)
    assert "Supported folders: inbox, sent, drafts" not in message
    for key in DISTINGUISHED_FOLDERS:
        assert key in message


@pytest.mark.asyncio
async def test_search_accepts_archive_folder_name() -> None:
    service, api = _service(
        [
            _success(Folders=[_folder("a-1", "Archive")]),
            _success(RootFolder={"Items": [], "IncludesLastItemInRange": True}),
        ]
    )
    out = await service.search_emails(folder="archive", days_back=30, max_results=10)
    assert out["meta"]["scannedFolders"] == 1
    assert api.calls[1][1]["ParentFolderIds"][0]["Id"] == "a-1"
