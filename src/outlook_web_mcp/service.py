from __future__ import annotations

import asyncio
import base64
import html as html_module
import json
from collections import deque
from contextlib import suppress
from datetime import UTC, datetime, timedelta
from pathlib import Path
from typing import Any

from . import index as index_store
from . import pixel, pixel_store, sent_links
from .api import OutlookWebApi, response_items, response_messages, root_folders, root_items
from .approval import consume, draft_digest
from .auth import AuthManager
from .config import OWA_ORIGIN, STORE_ID, Settings
from .errors import ApiError, OutlookWebMcpError
from .extract import text_from_attachment
from .models import CalendarEvent, FolderInfo, MailSummary
from .parsers import (
    change_key,
    html_links_and_text,
    item_id,
    looks_js_rendered,
    parse_calendar_event,
    parse_folder,
    parse_mail_detail,
    parse_mail_summary,
    recipient_text,
    record,
    records,
    string,
)
from .track import classify, normalize_subject

DISTINGUISHED_FOLDERS = {
    "inbox": "inbox",
    "sent": "sentitems",
    "sent_mail": "sentitems",
    "sentitems": "sentitems",
    "drafts": "drafts",
    "calendar": "calendar",
    "archive": "archive",
    "deleteditems": "deleteditems",
    "deleted_mail": "deleteditems",
    "junk": "junkemail",
    "junkemail": "junkemail",
    "outbox": "outbox",
    "root": "msgfolderroot",
}

# Folder roots that hold calendar/contact/task items, not mail.
NON_MAIL_ROOTS = {
    "calendar",
    "contacts",
    "tasks",
    "notes",
    "journal",
    "files",
    "conversation history",
    "conversation action settings",
    "yammer root",
    "birthdays",
    "companies",
    "gal contacts",
    "organizational contacts",
    "externalcontacts",
    "recipient cache",
    "team chat",
    "feeds",
    "peoplecentricconversation buddies",
    "top of information store",
}

# Stable index labels so rows written before this change keep their folder key.
FOLDER_LABELS = {
    "inbox": "inbox",
    "sent items": "sent",
    "archive": "archive",
    "deleted items": "deleteditems",
    "drafts": "drafts",
    "junk email": "junk",
    "outbox": "outbox",
}


def _property(field_uri: str) -> dict[str, str]:
    return {"__type": "PropertyUri:#Exchange", "FieldURI": field_uri}


def _mailbox_entry(address: str) -> dict[str, str]:
    """Flat recipient entry — the only shape this dialect actually stores.

    The EWS-style {Mailbox: {...}} wrapper is accepted with NoError and silently
    dropped; SendAndSaveCopy with it fails outright as ErrorInvalidRecipients.
    """
    return {
        "Name": address,
        "EmailAddress": address,
        "RoutingType": "SMTP",
        "MailboxType": "Mailbox",
    }


def _folder_id(folder_id: str, *, distinguished: bool = False) -> dict[str, str]:
    return {
        "__type": ("DistinguishedFolderId:#Exchange" if distinguished else "FolderId:#Exchange"),
        "Id": folder_id,
    }


def _target_folder(folder_id: str, *, distinguished: bool = False) -> dict[str, object]:
    return {
        "__type": "TargetFolderId:#Exchange",
        "BaseFolderId": _folder_id(folder_id, distinguished=distinguished),
    }


def _response_folders(data: object) -> list[dict[str, Any]]:
    for message in response_messages(data):
        folders = records(message.get("Folders"))
        if folders:
            return folders
    return []


def _parse_iso(value: str, parameter_name: str) -> datetime:
    normalized = value.strip().replace("Z", "+00:00")
    try:
        parsed = datetime.fromisoformat(normalized)
    except ValueError as exc:
        raise ValueError(
            f"{parameter_name} must be an ISO 8601 date and time with an offset"
        ) from exc
    if parsed.tzinfo is None:
        raise ValueError(f"{parameter_name} must include a timezone offset")
    return parsed


def _utc_string(value: datetime) -> str:
    return value.astimezone(UTC).isoformat(timespec="seconds").replace("+00:00", "Z")


def _calendar_string(value: datetime) -> str:
    return value.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%S.000")


def _text_to_html(text: str) -> str:
    """Escape a plain-text body and keep its line breaks, so it can carry a pixel."""
    escaped = html_module.escape(text, quote=False).replace("\r\n", "\n").replace("\r", "\n")
    paragraphs = [part.replace("\n", "<br>") for part in escaped.split("\n\n")]
    return "<div>" + "</div><div>".join(paragraphs) + "</div>"


class OutlookService:
    def __init__(
        self,
        auth: AuthManager,
        api: OutlookWebApi,
        settings: Settings | None = None,
    ) -> None:
        self.auth = auth
        self.api = api
        self.settings = settings or auth.settings

    async def auth_status(self) -> dict[str, object]:
        return await self.auth.ensure_status()

    @staticmethod
    def _validate_store(store_id: str) -> None:
        if store_id != STORE_ID:
            raise OutlookWebMcpError(
                "storeId is not an Outlook Web mailbox handle. Run the Outlook Web "
                "search or folder tool again and use the returned storeId.",
                "store_id_invalid",
            )

    async def _get_folder(
        self,
        folder_id: str,
        *,
        distinguished: bool,
        fallback_name: str = "",
    ) -> FolderInfo:
        data = await self.api.call(
            "GetFolder",
            {
                "__type": "GetFolderRequest:#Exchange",
                "FolderShape": {
                    "__type": "FolderResponseShape:#Exchange",
                    "BaseShape": "Default",
                    "AdditionalProperties": [
                        _property("ParentFolderId"),
                        _property("ChildFolderCount"),
                        _property("TotalCount"),
                        _property("UnreadCount"),
                    ],
                },
                "FolderIds": [_folder_id(folder_id, distinguished=distinguished)],
            },
        )
        folders = _response_folders(data)
        if not folders:
            raise ApiError("Outlook Web returned no folder.", "folder_not_found")
        raw = folders[0]
        name = string(raw.get("FolderDisplayName") or raw.get("DisplayName")) or fallback_name
        return parse_folder(raw, folder_path=name)

    async def _find_child_folders(self, parent_id: str, limit: int) -> list[dict[str, Any]]:
        data = await self.api.call(
            "FindFolder",
            {
                "__type": "FindFolderRequest:#Exchange",
                "FolderShape": {
                    "__type": "FolderResponseShape:#Exchange",
                    "BaseShape": "Default",
                    "AdditionalProperties": [
                        _property("ParentFolderId"),
                        _property("ChildFolderCount"),
                        _property("TotalCount"),
                        _property("UnreadCount"),
                        _property("FolderClass"),
                    ],
                },
                "ParentFolderIds": [_folder_id(parent_id)],
                "Traversal": "Shallow",
                "Paging": {
                    "__type": "IndexedPageView:#Exchange",
                    "BasePoint": "Beginning",
                    "Offset": 0,
                    "MaxEntriesReturned": min(max(limit, 1), 200),
                },
            },
        )
        return root_folders(data)

    async def _resolve_folder(self, name: str) -> FolderInfo:
        normalized = name.strip().casefold()
        distinguished = DISTINGUISHED_FOLDERS.get(normalized)
        if distinguished is None:
            raise ValueError(
                "Unsupported folder. Supported: "
                + ", ".join(sorted(DISTINGUISHED_FOLDERS))
                + ". For any other folder, pass folderId from "
                'list_mail_folders(folder="root").'
            )
        return await self._get_folder(
            distinguished,
            distinguished=True,
            fallback_name=normalized,
        )

    async def list_mail_folders(
        self,
        *,
        folder: str = "inbox",
        parent_folder_id: str | None = None,
        store_id: str | None = None,
        recursive: bool = True,
        max_results: int = 100,
        max_depth: int = 5,
    ) -> dict[str, object]:
        if not 1 <= max_results <= 200:
            raise ValueError("maxResults must be between 1 and 200")
        if not 1 <= max_depth <= 10:
            raise ValueError("maxDepth must be between 1 and 10")
        if parent_folder_id:
            if store_id is None:
                raise ValueError("storeId is required when parentFolderId is supplied")
            self._validate_store(store_id)
            root = await self._get_folder(
                parent_folder_id,
                distinguished=False,
                fallback_name="Selected folder",
            )
        else:
            root = await self._resolve_folder(folder)

        output: list[FolderInfo] = [root]
        queue: deque[tuple[FolderInfo, int]] = deque([(root, 0)])
        while recursive and queue and len(output) < max_results:
            parent, depth = queue.popleft()
            if depth >= max_depth:
                continue
            remaining = max_results - len(output)
            children = await self._find_child_folders(parent.folder_id, remaining)
            if children:
                parent.has_children = True
            for raw in children[:remaining]:
                name = string(raw.get("FolderDisplayName") or raw.get("DisplayName"))
                path = f"{parent.folder_path}\\{name}" if parent.folder_path else name
                child = parse_folder(
                    raw,
                    folder_path=path,
                    fallback_parent_id=parent.folder_id,
                )
                output.append(child)
                queue.append((child, depth + 1))
                if len(output) >= max_results:
                    break
        return {"result": [item.to_dict(STORE_ID) for item in output]}

    async def _mail_page(
        self,
        folder_id: str,
        *,
        unread_only: bool,
        include_preview: bool,
        offset: int,
        limit: int,
    ) -> tuple[list[dict[str, Any]], bool, int]:
        properties = [
            _property("Subject"),
            _property("From"),
            _property("Sender"),
            _property("DateTimeReceived"),
            _property("IsRead"),
            _property("HasAttachments"),
            _property("ItemClass"),
            _property("InternetMessageId"),
            _property("References"),
        ]
        if include_preview:
            properties.append(_property("Preview"))

        data = await self.api.call(
            "FindItem",
            {
                "__type": "FindItemRequest:#Exchange",
                "ItemShape": {
                    "__type": "ItemResponseShape:#Exchange",
                    "BaseShape": "IdOnly",
                    "AdditionalProperties": properties,
                },
                "ParentFolderIds": [_folder_id(folder_id)],
                "Traversal": "Shallow",
                "Paging": {
                    "__type": "IndexedPageView:#Exchange",
                    "BasePoint": "Beginning",
                    "Offset": offset,
                    "MaxEntriesReturned": limit,
                },
                "ViewFilter": "Unread" if unread_only else "All",
                "SortOrder": [
                    {
                        "__type": "SortResults:#Exchange",
                        "Order": "Descending",
                        "Path": _property("DateTimeReceived"),
                    }
                ],
            },
        )
        page = root_items(data)
        includes_last = True
        next_offset = offset + len(page)
        for message in response_messages(data):
            root = record(message.get("RootFolder"))
            if root:
                includes_last = bool(root.get("IncludesLastItemInRange", True))
                with suppress(TypeError, ValueError):
                    next_offset = int(root.get("IndexedPagingOffset") or next_offset)
                break
        return page, includes_last, next_offset

    async def _search_folder_ids(
        self,
        root: FolderInfo,
        *,
        include_subfolders: bool,
    ) -> list[str]:
        ids = [root.folder_id]
        if not include_subfolders:
            return ids
        queue: deque[tuple[str, int]] = deque([(root.folder_id, 0)])
        while queue and len(ids) < self.settings.max_scanned_folders:
            parent_id, depth = queue.popleft()
            if depth >= 10:
                continue
            children = await self._find_child_folders(
                parent_id,
                self.settings.max_scanned_folders - len(ids),
            )
            for raw in children:
                folder_class = string(raw.get("FolderClass"))
                if folder_class and not folder_class.startswith("IPF.Note"):
                    continue
                child_id = string(record(raw.get("FolderId")).get("Id"))
                if not child_id or child_id in ids:
                    continue
                ids.append(child_id)
                queue.append((child_id, depth + 1))
                if len(ids) >= self.settings.max_scanned_folders:
                    break
        return ids

    async def search_emails(
        self,
        *,
        folder: str = "inbox",
        query: str | None = None,
        days_back: int = 30,
        max_results: int = 20,
        include_body_preview: bool = False,
        folder_id: str | None = None,
        store_id: str | None = None,
        include_subfolders: bool = False,
        unread_only: bool = False,
        offset: int = 0,
        continuation_folder_id: str | None = None,
        continuation_offset: int | None = None,
    ) -> dict[str, object]:
        # The server has no such limit; ten years covers any mailbox we care about,
        # and long windows are walked page by page with the cursor instead of re-scanning.
        if not 1 <= days_back <= 3650:
            raise ValueError("daysBack must be between 1 and 3650")
        if not 1 <= max_results <= 200:
            raise ValueError("maxResults must be between 1 and 200")
        if offset < 0:
            raise ValueError("offset must be zero or positive")
        if folder_id:
            if store_id is None:
                raise ValueError("storeId is required when folderId is supplied")
            self._validate_store(store_id)
            root = await self._get_folder(
                folder_id,
                distinguished=False,
                fallback_name="Selected folder",
            )
        else:
            root = await self._resolve_folder(folder)
        folder_ids = await self._search_folder_ids(
            root,
            include_subfolders=include_subfolders,
        )
        cutoff_dt = datetime.now(UTC) - timedelta(days=days_back)
        started_at = datetime.now(UTC)
        normalized_query = (query or "").strip().casefold()
        found: list[MailSummary] = []
        seen: set[str] = set()
        scanned = 0
        page_size = 200
        scan_capped = False
        folder_capped = False
        more_available = False

        if len(folder_ids) >= self.settings.max_scanned_folders:
            folder_capped = True

        resume_offset = 0
        if continuation_folder_id is not None:
            if continuation_folder_id not in folder_ids:
                raise ValueError(
                    "continuation folderId is not in the current folder set; "
                    "re-run the search without continuation to restart"
                )
            folder_ids = folder_ids[folder_ids.index(continuation_folder_id) :]
            resume_offset = max(int(continuation_offset or 0), 0)
        last_folder_id: str | None = None
        last_offset = 0
        last_scanned_folder_id: str | None = None
        next_offset = 0

        for current_folder_id in folder_ids:
            last_scanned_folder_id = current_folder_id
            page_offset = resume_offset if current_folder_id == folder_ids[0] else 0
            includes_last = True
            next_offset = page_offset
            while scanned < self.settings.max_scanned_items:
                request_limit = min(page_size, self.settings.max_scanned_items - scanned)
                page, includes_last, next_offset = await self._mail_page(
                    current_folder_id,
                    unread_only=unread_only,
                    include_preview=include_body_preview,
                    offset=page_offset,
                    limit=request_limit,
                )
                scanned += len(page)
                reached_cutoff = False
                for raw in page:
                    summary = parse_mail_summary(
                        raw,
                        include_preview=include_body_preview,
                    )
                    if not summary.email_id or summary.email_id in seen:
                        continue
                    try:
                        received = _parse_iso(summary.received_at, "DateTimeReceived")
                    except ValueError:
                        received = None
                    if received is not None and received < cutoff_dt:
                        reached_cutoff = True
                        continue
                    if unread_only and not summary.is_unread:
                        continue
                    searchable = "\n".join(
                        [summary.subject, summary.sender_name, summary.sender_address]
                    ).casefold()
                    if normalized_query and normalized_query not in searchable:
                        continue
                    seen.add(summary.email_id)
                    found.append(summary)
                if reached_cutoff or includes_last or not page or next_offset <= page_offset:
                    break
                page_offset = next_offset
                if not normalized_query and not include_subfolders and len(found) >= max_results:
                    # We stopped on the result cap, not because the mailbox ran out.
                    more_available = True
                    break
            if scanned >= self.settings.max_scanned_items:
                scan_capped = True
                if not includes_last:
                    last_folder_id = current_folder_id
                    last_offset = next_offset
                break

        def _received_key(item: MailSummary) -> tuple[int, str]:
            try:
                return (1, _utc_string(_parse_iso(item.received_at, "DateTimeReceived")))
            except ValueError:
                return (0, item.received_at)

        found.sort(key=_received_key, reverse=True)
        selected = found[offset : offset + max_results]
        meta: dict[str, object] = {
            "engine": "local",
            "scannedItems": scanned,
            "scannedFolders": len(folder_ids),
            "coverage": {
                "from": _utc_string(cutoff_dt),
                "to": _utc_string(started_at),
            },
            "matchCount": len(found),
            "returnedCount": len(selected),
            "truncated": scan_capped or folder_capped,
            "truncatedReason": (
                "scan_cap" if scan_capped else ("folder_cap" if folder_capped else None)
            ),
            "hitCapReached": len(found) > offset + len(selected) or more_available,
            "offset": offset,
            "continuation": (
                {"folderId": last_folder_id, "offset": last_offset}
                if last_folder_id is not None
                else None
            ),
            "nextPage": (
                {"folderId": last_scanned_folder_id, "offset": next_offset}
                if last_scanned_folder_id is not None
                else None
            ),
        }
        return {"result": [item.to_dict(STORE_ID) for item in selected], "meta": meta}

    async def get_email_html(
        self,
        *,
        email_id: str,
        store_id: str,
        max_body_characters: int = 20_000,
    ) -> dict[str, object]:
        """Raw HTML body plus extracted http(s) links and readable text.

        Mail bodies are untrusted: links are returned as data for the caller to
        judge, never followed. Non-http schemes (mailto, javascript) are dropped.
        """
        self._validate_store(store_id)
        if not email_id.strip():
            raise ValueError("emailId is required")
        if not 1 <= max_body_characters <= 50_000:
            raise ValueError("maxBodyCharacters must be between 1 and 50000")
        raw = await self._get_item_raw(email_id, body_type="HTML")
        body = record(raw.get("Body"))
        value = string(body.get("Value"))
        is_html = string(body.get("BodyType")).casefold() == "html"
        if is_html:
            links, text_content = html_links_and_text(value)
        else:
            links = []
            text_content = value.replace("\r\n", "\n").replace("\r", "\n")
            value = text_content
        truncated = len(value) > max_body_characters
        return {
            "emailId": email_id,
            "storeId": STORE_ID,
            "html": value[:max_body_characters],
            "links": links,
            "textContent": text_content,
            "bodyTruncated": truncated,
            # Hint only: take the screenshot+OCR runbook when this is true.
            "likelyJsRendered": is_html and looks_js_rendered(value, text_content),
        }

    async def _get_item_raw(self, email_id: str, *, body_type: str = "Text") -> dict[str, Any]:
        data = await self.api.call(
            "GetItem",
            {
                "__type": "GetItemRequest:#Exchange",
                "ItemShape": {
                    "__type": "ItemResponseShape:#Exchange",
                    "BaseShape": "AllProperties",
                    "BodyType": body_type,
                },
                "ItemIds": [{"__type": "ItemId:#Exchange", "Id": email_id}],
            },
        )
        items = response_items(data)
        if not items:
            raise ApiError("The Outlook item was not found.", "item_not_found")
        return items[0]

    async def get_email(
        self,
        *,
        email_id: str,
        store_id: str,
        max_body_characters: int = 20_000,
    ) -> dict[str, object]:
        self._validate_store(store_id)
        if not email_id.strip():
            raise ValueError("emailId is required")
        if not 1 <= max_body_characters <= 50_000:
            raise ValueError("maxBodyCharacters must be between 1 and 50000")
        raw = await self._get_item_raw(email_id)
        detail = parse_mail_detail(raw, max_body_characters=max_body_characters)
        return detail.to_dict(STORE_ID)

    async def set_email_read_state(
        self,
        *,
        email_id: str,
        store_id: str,
        is_read: bool,
    ) -> dict[str, object]:
        self._validate_store(store_id)
        raw = await self._get_item_raw(email_id)
        key = change_key(raw)
        item_ref: dict[str, object] = {
            "__type": "ItemId:#Exchange",
            "Id": email_id,
        }
        if key:
            item_ref["ChangeKey"] = key
        await self.api.call(
            "UpdateItem",
            {
                "__type": "UpdateItemRequest:#Exchange",
                "ItemChanges": [
                    {
                        "__type": "ItemChange:#Exchange",
                        "ItemId": item_ref,
                        "Updates": [
                            {
                                "__type": "SetItemField:#Exchange",
                                "Path": _property("IsRead"),
                                "Item": {
                                    "__type": "Message:#Exchange",
                                    "IsRead": is_read,
                                },
                            }
                        ],
                    }
                ],
                "ConflictResolution": "AutoResolve",
                "MessageDisposition": "SaveOnly",
            },
        )
        summary = parse_mail_summary(raw, include_preview=False)
        return {
            "emailId": email_id,
            "storeId": STORE_ID,
            "subject": summary.subject,
            "receivedAt": summary.received_at,
            "isRead": is_read,
            "status": "read_state_updated",
        }

    async def create_draft(
        self,
        *,
        to: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        html: bool = False,
        track: bool | None = None,
    ) -> dict[str, object]:
        if not to:
            raise ValueError("to must contain at least one address")
        for address in [*to, *(cc or [])]:
            if "@" not in address or any(char in address for char in " ,;<>"):
                raise ValueError(f"invalid recipient address: {address}")
        if not subject.strip():
            raise ValueError("subject is required")
        if not body.strip():
            raise ValueError("body is required")
        if len(body) > 20_000:
            raise ValueError("body cannot exceed 20000 characters")
        tracked = self.settings.track_pixels if track is None else track
        token: str | None = None
        tracking_warning: str | None = None
        if tracked and not self.settings.pixel_base:
            tracking_warning = (
                "OUTLOOK_WEB_PIXEL_BASE is not configured; the draft was saved "
                "without a tracking pixel"
            )
            tracked = False
        if tracked and not html:
            if track is True:
                raise ValueError("track=True requires html=True")
            # Default-on tracking must actually happen: carry the text body as escaped
            # HTML with <br> breaks, which is how Outlook renders a text draft anyway.
            body = _text_to_html(body)
            html = True
        if tracked and self.settings.pixel_base:
            token = pixel.new_token()
            body = pixel.inject(body, base=self.settings.pixel_base, token=token)
        drafts = await self._get_folder("drafts", distinguished=True, fallback_name="Drafts")
        # Recipients must be flat entries ({EmailAddress, RoutingType...}), not the
        # EWS-style {Mailbox: {...}} wrapper: the server accepts the wrapper, returns
        # NoError, and silently drops every recipient. See PITFALLS #23.
        message: dict[str, object] = {
            "__type": "Message:#Exchange",
            "Subject": subject,
            "ToRecipients": [_mailbox_entry(address) for address in to],
            "Body": {
                "__type": "BodyContentType:#Exchange",
                "BodyType": "HTML" if html else "Text",
                "Value": body,
            },
        }
        if cc:
            message["CcRecipients"] = [_mailbox_entry(address) for address in cc]
        data = await self.api.call(
            "CreateItem",
            {
                "__type": "CreateItemRequest:#Exchange",
                "MessageDisposition": "SaveOnly",
                "SavedItemFolderId": _target_folder(drafts.folder_id),
                "Items": [message],
            },
        )
        created = response_items(data)
        draft_id = item_id(created[0]) if created else ""
        if not draft_id:
            raise ApiError("Outlook Web saved no identifiable draft.", "draft_id_missing")
        saved = await self._get_item_raw(draft_id, body_type="Text" if not html else "HTML")
        display_to = string(saved.get("DisplayTo"))
        if not display_to:
            # Fail loud: a draft that cannot say who it is for must never look saved.
            raise ApiError(
                "Outlook Web saved the draft without its recipients.",
                "recipients_dropped",
            )
        if token is not None and self.settings.pixel_base:
            pixel_store.register(
                self.settings.pixel_dir,
                token=token,
                draft_id=draft_id,
                url=pixel.tracking_url(self.settings.pixel_base, token),
                at=datetime.now(UTC).isoformat(timespec="seconds"),
            )
        result: dict[str, object] = {
            "draftId": draft_id,
            "storeId": STORE_ID,
            "folder": "drafts",
            "status": "saved_to_drafts",
            "sent": False,
            "tracked": token is not None,
        }
        if tracking_warning:
            result["trackingWarning"] = tracking_warning
        return result

    async def _mail_folder_targets(self) -> list[tuple[str, str, int]]:
        """Every folder that actually holds mail: (label, folderId, itemCount).

        Calendar/Contacts/Tasks subtrees are excluded by their path root, so a
        holidays calendar with 208 entries is never mistaken for 208 emails.
        """
        listing = await self.list_mail_folders(folder="root", recursive=True, max_results=200)
        targets: list[tuple[str, str, int]] = []
        rows: list[dict[str, Any]] = listing.get("result") or []  # type: ignore[assignment]
        for row in rows:
            count = int(row.get("totalItemCount") or 0)
            if count <= 0:
                continue
            path = str(row.get("folderPath") or "")
            parts = [part for part in path.replace("/", "\\").split("\\") if part]
            root = parts[1].casefold() if len(parts) > 1 else str(row.get("name") or "").casefold()
            if root in NON_MAIL_ROOTS:
                continue
            name = str(row.get("name") or "")
            label = FOLDER_LABELS.get(name.casefold(), name.casefold())
            targets.append((label, str(row.get("folderId") or ""), count))
        return targets

    async def index_sync(
        self,
        *,
        folders: list[str] | None = None,
        days_back: int = 30,
        page_size: int = 200,
        max_rows: int = 20_000,
    ) -> dict[str, object]:
        if not 1 <= page_size <= 200:
            raise ValueError("page_size must be between 1 and 200")
        if folders is None:
            # Every mail-bearing folder, including Drafts, Junk and custom folders.
            targets = [
                (label, folder_id) for label, folder_id, _ in await self._mail_folder_targets()
            ]
            if not targets:
                raise ValueError("no mail-bearing folders were found")
        else:
            if not folders:
                raise ValueError("folders must contain at least one folder name")
            targets = []
            for name in folders:
                key = name.strip().casefold()
                if key not in DISTINGUISHED_FOLDERS:
                    raise ValueError(
                        f"Unknown folder: {name}. Omit folders to index every mail folder."
                    )
                targets.append((FOLDER_LABELS.get(key, key), ""))
        folders = [label for label, _ in targets]
        connection = index_store.open_index(self.settings.index_path)
        indexed_ids: set[str] = set()
        scanned = 0
        pruned = 0
        truncated = False
        window_start = _utc_string(datetime.now(UTC) - timedelta(days=days_back))
        try:
            for folder, folder_id in targets:
                offset = 0
                seen: set[str] = set()
                folder_complete = True
                continuation_folder: str | None = None
                continuation_offset: int | None = None
                while True:
                    page: dict[str, Any] = await self.search_emails(
                        folder=folder,
                        folder_id=folder_id or None,
                        store_id=STORE_ID if folder_id else None,
                        days_back=days_back,
                        max_results=page_size,
                        include_body_preview=True,
                        offset=offset,
                        continuation_folder_id=continuation_folder,
                        continuation_offset=continuation_offset,
                    )
                    rows: list[dict[str, Any]] = page.get("result") or []
                    new_ids = {
                        str(row.get("emailId"))
                        for row in rows
                        if str(row.get("emailId") or "") not in seen
                    }
                    index_store.upsert_rows(connection, rows, folder)
                    indexed_ids.update(new_ids)
                    seen.update(new_ids)
                    meta: dict[str, Any] = page.get("meta") or {}
                    scanned += int(meta["scannedItems"])
                    if len(indexed_ids) >= max_rows:
                        truncated = True
                        folder_complete = False
                        break
                    continuation = meta.get("continuation")
                    if continuation:
                        continuation_folder = str(continuation["folderId"])
                        continuation_offset = int(continuation["offset"])
                        offset = 0
                        continue
                    if meta["hitCapReached"]:
                        cursor = meta.get("nextPage") or {}
                        next_page = int(cursor.get("offset") or 0)
                        if not new_ids or next_page <= 0:
                            truncated = True
                            folder_complete = False
                            break
                        # Reuse the continuation mechanism to advance the server-side
                        # cursor instead of re-scanning this window from the top.
                        continuation_folder = str(cursor.get("folderId") or "")
                        continuation_offset = next_page
                        offset = 0
                        continue
                    if meta["truncated"]:
                        truncated = True
                        folder_complete = False
                    break
                if folder_complete:
                    # The whole window was walked, so rows we did not see are gone from
                    # the mailbox (deleted or moved). SyncFolderItems is unavailable here.
                    pruned += index_store.prune_folder(
                        connection, folder, since=window_start, keep_ids=seen
                    )
        finally:
            connection.close()
        return {
            "indexed": len(indexed_ids),
            "scanned": scanned,
            "pruned": pruned,
            "folders": folders,
            "truncated": truncated,
        }

    async def search_index(
        self,
        *,
        query: str,
        limit: int = 50,
        folder: str | None = None,
        since: str | None = None,
    ) -> dict[str, object]:
        connection = index_store.open_index(self.settings.index_path)
        try:
            rows = index_store.search(connection, query, folder=folder, since=since, limit=limit)
            info = index_store.status(connection)
            bodies = index_store.body_progress(connection)
        finally:
            connection.close()
        return {"result": rows, "meta": {"engine": "index", **info, "bodies": bodies}}

    async def index_bodies(
        self,
        *,
        max_fetches: int = 100,
        folder: str | None = None,
        concurrency: int = 4,
    ) -> dict[str, object]:
        """Pull full body text for messages that have not been fetched yet.

        One GetItem per message, newest first, bounded per call so a large backfill is
        a series of resumable steps rather than one long request. Attachment files are
        never downloaded: only a summary line and the OWA link are stored.
        """
        if not 1 <= max_fetches <= 1000:
            raise ValueError("maxFetches must be between 1 and 1000")
        if not 1 <= concurrency <= 8:
            raise ValueError("concurrency must be between 1 and 8")
        connection = index_store.open_index(self.settings.index_path)
        fetched = 0
        failed = 0
        try:
            ids = index_store.rows_needing_body(connection, limit=max_fetches, folder=folder)
            gate = asyncio.Semaphore(concurrency)

            async def fetch(email_id: str) -> dict[str, Any] | None:
                # Network is the slow half; writes stay serialised on one connection.
                async with gate:
                    try:
                        return await self._get_item_raw(email_id, body_type="HTML")
                    except (ApiError, OutlookWebMcpError):
                        return None

            raws = await asyncio.gather(*(fetch(email_id) for email_id in ids))
            for email_id, raw in zip(ids, raws, strict=True):
                if raw is None:
                    failed += 1
                    continue
                html = string(record(raw.get("Body")).get("Value"))
                detail = parse_mail_detail(raw, max_body_characters=200_000)
                attachments = [
                    {
                        "name": string(entry.get("Name")),
                        "contentType": string(entry.get("ContentType")),
                        "size": int(entry.get("Size") or 0),
                        "attachmentId": string(record(entry.get("AttachmentId")).get("Id")),
                        "summary": (
                            f"{string(entry.get('ContentType')) or 'file'}, "
                            f"{int(entry.get('Size') or 0)} bytes — not downloaded; "
                            "the file stays on the server"
                        ),
                    }
                    for entry in records(record(raw).get("Attachments"))
                ]
                link = string(raw.get("WebClientReadFormQueryString"))
                if link and not link.startswith("http"):
                    link = f"{OWA_ORIGIN}/owa/{link.lstrip('/')}"
                index_store.store_body(
                    connection,
                    email_id,
                    body_text=detail.body or html,
                    web_link=link,
                    attachments=attachments,
                )
                fetched += 1
            progress = index_store.body_progress(connection)
        finally:
            connection.close()
        return {
            "fetched": fetched,
            "failed": failed,
            "remaining": progress["total"] - progress["fetched"],
            "bodies": progress,
        }

    async def _get_attachment(self, attachment_id: str) -> bytes:
        data = await self.api.call(
            "GetAttachment",
            {
                "__type": "GetAttachmentRequest:#Exchange",
                "AttachmentIds": [{"__type": "AttachmentIdType:#Exchange", "Id": attachment_id}],
                "BodyType": "Best",
            },
        )
        for message in response_items(data):
            for entry in records(record(message).get("Attachments")):
                content = string(entry.get("Content"))
                if content:
                    return base64.b64decode(content)
        raise ApiError("The attachment returned no content.", "attachment_empty")

    async def index_attachment_text(
        self,
        *,
        max_files: int = 20,
        max_bytes: int = 10_000_000,
    ) -> dict[str, object]:
        """Extract searchable text from attachments without keeping the file.

        Only text-bearing types under `max_bytes` are attempted; anything else stays
        metadata-only. The bytes are decoded, turned into text, and dropped — the
        file itself remains on the server.
        """
        if not 1 <= max_files <= 200:
            raise ValueError("maxFiles must be between 1 and 200")
        if not 1_000 <= max_bytes <= 50_000_000:
            raise ValueError("maxBytes must be between 1000 and 50000000")
        connection = index_store.open_index(self.settings.index_path)
        fetched = 0
        failed = 0
        skipped = 0
        try:
            rows = index_store.attachments_needing_text(
                connection, limit=max_files, max_bytes=max_bytes
            )
            for row in rows:
                try:
                    data = await self._get_attachment(str(row["attachment_id"]))
                except (ApiError, OutlookWebMcpError):
                    failed += 1
                    continue
                text, method = text_from_attachment(
                    str(row["name"]), str(row["content_type"]), data
                )
                if not text and method.startswith("unsupported"):
                    skipped += 1
                index_store.store_attachment_text(
                    connection,
                    str(row["email_id"]),
                    str(row["name"]),
                    text_body=text,
                    method=method,
                )
                fetched += 1
            progress = index_store.attachment_progress(connection)
        finally:
            connection.close()
        return {
            "fetched": fetched,
            "failed": failed,
            "skipped": skipped,
            "remaining": progress["total"] - progress["fetched"],
            "attachments": progress,
        }

    async def index_verify(self) -> dict[str, object]:
        """Compare indexed row counts with the mailbox's own folder totals.

        A gap means the index is incomplete for that folder (usually a shorter window)
        or that mail arrived since the last sync; the index cannot tell those apart,
        which is why this check exists.
        """
        connection = index_store.open_index(self.settings.index_path)
        try:
            rows = connection.execute(
                "SELECT folder, COUNT(*) AS c FROM items GROUP BY folder"
            ).fetchall()
        finally:
            connection.close()
        counts = {str(row["folder"]): int(row["c"]) for row in rows}
        # The universe is every folder that holds mail, not a hardcoded four: a folder
        # with mail but no indexed rows must be reported, or "complete" would be a lie.
        discovered = {label: total for label, _, total in await self._mail_folder_targets()}
        expected = sorted(set(counts) | set(discovered))
        report: list[dict[str, object]] = []
        for folder in expected:
            indexed = counts.get(folder, 0)
            if folder in discovered:
                total = discovered[folder]
            else:
                try:
                    total = (await self._resolve_folder(folder)).total_item_count
                except ValueError:
                    report.append(
                        {
                            "folder": folder,
                            "indexed": indexed,
                            "mailbox": None,
                            "gap": None,
                            "complete": False,
                            "note": "folder name is no longer resolvable",
                        }
                    )
                    continue
            report.append(
                {
                    "folder": folder,
                    "indexed": indexed,
                    "mailbox": total,
                    "gap": max(total - indexed, 0),
                    "complete": total == indexed,
                }
            )
        return {
            "result": report,
            "expectedFolders": expected,
            "complete": bool(report) and all(bool(item["complete"]) for item in report),
        }

    async def index_status(self) -> dict[str, object]:
        connection = index_store.open_index(self.settings.index_path)
        try:
            info = index_store.status(connection)
        finally:
            connection.close()
        return {"path": self.settings.index_path, **info}

    async def approval_digest(self, draft_id: str) -> dict[str, object]:
        raw = await self._get_item_raw(draft_id, body_type="HTML")
        html_body = string(record(raw.get("Body")).get("Value"))
        detail = parse_mail_detail(raw, max_body_characters=200_000)
        # Bind the approval to what is actually sent: the HTML body. Hashing the
        # plain-text rendering would ignore hrefs and embedded images, so a link swap
        # could ride on an old approval.
        body_for_digest = html_body or detail.body
        bcc = recipient_text(record(raw).get("BccRecipients"))
        attachments = tuple(
            f"{string(entry.get('Name'))}:{int(entry.get('Size') or 0)}"
            for entry in records(record(raw).get("Attachments"))
        )
        digest = draft_digest(
            draft_id=draft_id,
            subject=detail.subject,
            recipients=f"{detail.to}; {detail.cc}",
            body=body_for_digest,
            bcc=bcc,
            attachments=attachments,
        )
        return {
            "draftId": draft_id,
            "digest": digest,
            "subject": detail.subject,
            "to": detail.to,
            "cc": detail.cc,
            "bcc": bcc,
            "attachments": list(attachments),
            "bodyCharacters": len(body_for_digest),
        }

    async def send_draft(self, *, draft_id: str, store_id: str) -> dict[str, object]:
        """Send an existing draft, only after the user approved this exact content."""
        self._validate_store(store_id)
        info = await self.approval_digest(draft_id)
        if not str(info.get("to") or "") and not str(info.get("cc") or ""):
            # A draft the digest cannot attribute to anyone must never send: with the
            # wrapper-shape bug this was a real mode, not a hypothetical one.
            raise OutlookWebMcpError(
                "This draft has no recipients; sending it would deliver nothing. "
                "Re-create the draft.",
                "draft_has_no_recipients",
            )
        digest = str(info["digest"])
        token = await self._draft_pixel_token(draft_id)
        if token is not None and not pixel_store.known(self.settings.pixel_dir, token):
            raise OutlookWebMcpError(
                "This draft carries a tracking pixel this server did not create; "
                "refusing to send it. Remove the pixel, or re-create the draft.",
                "pixel_unknown",
            )
        if not consume(self.settings.approval_dir, digest):
            raise OutlookWebMcpError(
                "Sending this draft requires the user's one-time approval of its exact "
                "content, and approvals expire after 30 minutes. Run: "
                f"outlook-web-mcp approve {draft_id}  (digest {digest})",
                "send_not_approved",
            )
        # OWA's SendItem needs a ChangeKey matching the current item version.
        # 与同文件 set_email_read_state 的做法保持一致。
        raw = await self._get_item_raw(draft_id)
        key = change_key(raw)
        item_ref: dict[str, object] = {
            "__type": "ItemId:#Exchange",
            "Id": draft_id,
        }
        if key:
            item_ref["ChangeKey"] = key
        try:
            await self.api.call(
                "SendItem",
                {
                    "__type": "SendItemRequest:#Exchange",
                    "ItemIds": [item_ref],
                    "SaveItemToFolder": True,
                },
            )
        except ApiError as exc:
            if exc.http_status != 500:
                raise
            # SendItem 500s on this dialect even with a fresh ChangeKey. The
            # proven path is CreateItem + SendAndSaveCopy with flat recipient
            # entries, re-creating the exact approved content (PITFALLS #24).
            # The draft itself stays in Drafts — there is no delete tool.
            html_body = string(record(raw.get("Body")).get("Value"))
            if not html_body:
                raise ApiError(
                    "The draft has no readable body; cannot re-create it for sending.",
                    "draft_body_missing",
                ) from exc
            message: dict[str, object] = {
                "__type": "Message:#Exchange",
                "Subject": string(raw.get("Subject")),
                "Body": {
                    "__type": "BodyContentType:#Exchange",
                    "BodyType": "HTML"
                    if record(raw.get("Body")).get("BodyType") == "HTML"
                    else "Text",
                    "Value": html_body,
                },
            }
            for field, target in (
                ("ToRecipients", "ToRecipients"),
                ("CcRecipients", "CcRecipients"),
                ("BccRecipients", "BccRecipients"),
            ):
                entries = records(raw.get(field))
                if entries:
                    message[target] = [
                        {k: v for k, v in entry.items() if k != "__type"} for entry in entries
                    ]
            await self.api.call(
                "CreateItem",
                {
                    "__type": "CreateItemRequest:#Exchange",
                    "MessageDisposition": "SendAndSaveCopy",
                    "Items": [message],
                },
            )
        await self._link_sent_item(draft_id=draft_id, subject=str(info.get("subject") or ""))
        if token is not None:
            pixel_store.mark_sent(
                self.settings.pixel_dir,
                token,
                datetime.now(UTC).isoformat(timespec="seconds"),
            )
        return {
            "draftId": draft_id,
            "storeId": STORE_ID,
            "status": "sent",
            "digest": digest,
            "tracked": token is not None,
        }

    async def _draft_pixel_token(self, draft_id: str) -> str | None:
        raw = await self._get_item_raw(draft_id, body_type="HTML")
        html_body = string(record(raw.get("Body")).get("Value"))
        return pixel.find_token(html_body)

    async def _link_sent_item(self, *, draft_id: str, subject: str) -> None:
        """Remember which sent item came from this draft. Best effort; never blocks a send."""
        try:
            page: dict[str, Any] = await self.search_emails(
                folder="sent",
                query=subject.strip() or None,
                days_back=1,
                max_results=50,
            )
            rows: list[dict[str, Any]] = page.get("result") or []
        except Exception:
            # A failed lookup must not undo a message that already went out.
            return
        wanted = normalize_subject(subject)
        now = datetime.now(UTC)
        best: dict[str, Any] | None = None
        for row in rows:
            if normalize_subject(str(row.get("subject") or "")) != wanted:
                continue
            try:
                received = datetime.fromisoformat(
                    str(row.get("receivedAt") or "").replace("Z", "+00:00")
                )
            except ValueError:
                continue
            if abs((now - received).total_seconds()) > 900:
                continue
            if best is None or str(row.get("receivedAt")) > str(best.get("receivedAt")):
                best = row
        if best is None:
            return
        sent_links.record(
            self.settings.sent_links_dir,
            draft_id=draft_id,
            sent_id=str(best.get("emailId") or ""),
            internet_message_id=str(best.get("internetMessageId") or ""),
            at=str(best.get("receivedAt") or now.isoformat(timespec="seconds")),
        )

    async def track_status(
        self,
        *,
        draft_id: str,
        store_id: str,
        days_back: int = 30,
        max_rows: int = 200,
    ) -> dict[str, object]:
        self._validate_store(store_id)
        draft = parse_mail_detail(
            await self._get_item_raw(draft_id),
            max_body_characters=1,
        )
        query = draft.subject.strip() or None
        link = sent_links.lookup(self.settings.sent_links_dir, draft.email_id)
        linked_row: dict[str, Any] | None = None
        if link and link.get("sentId"):
            linked_row = {
                "emailId": str(link.get("sentId")),
                "subject": draft.subject,
                "senderAddress": "",
                "receivedAt": str(link.get("at") or ""),
                "messageClass": "IPM.Note",
                "references": "",
                "internetMessageId": str(link.get("internetMessageId") or ""),
            }
        # A reply that predates the default window is invisible, which would read as
        # "no reply". When we know when this was sent, look back far enough to include it.
        scan_days = days_back
        if linked_row and linked_row["receivedAt"]:
            try:
                sent_dt = _parse_iso(str(linked_row["receivedAt"]), "sentAt")
                scan_days = max(days_back, min((datetime.now(UTC) - sent_dt).days + 7, 3650))
            except ValueError:
                scan_days = days_back
        rows_by_folder: dict[str, list[dict[str, Any]]] = {}
        scans: dict[str, dict[str, Any]] = {}
        # Replies and bounces get filed away, so inbox alone would report false negatives.
        for name in ("sent", "inbox", "archive"):
            page: dict[str, Any] = await self.search_emails(
                folder=name,
                query=query,
                days_back=scan_days,
                max_results=max_rows,
            )
            rows_by_folder[name] = page.get("result") or []
            scans[name] = page.get("meta") or {}

        reply_rows: list[dict[str, Any]] = []
        seen_ids: set[str] = set()
        for name in ("inbox", "archive"):
            for row in rows_by_folder[name]:
                email_id = str(row.get("emailId") or "")
                if email_id and email_id in seen_ids:
                    continue
                seen_ids.add(email_id)
                reply_rows.append(row)

        verdict = classify(
            {
                "emailId": draft.email_id,
                "subject": draft.subject,
                "to": draft.to,
                "cc": draft.cc,
                "internetMessageId": draft.internet_message_id,
            },
            rows_by_folder["sent"],
            reply_rows,
            linked_sent=linked_row,
            coverage={
                name: {"truncated": bool(meta.get("truncated") or meta.get("hitCapReached"))}
                for name, meta in scans.items()
            },
        )
        pixel_entry = pixel_store.for_draft(self.settings.pixel_dir, draft.email_id)
        opened: dict[str, object] | None = None
        if pixel_entry is not None:
            events_file = Path(self.settings.pixel_dir) / "events.json"
            raw_events: object = []
            if events_file.exists():
                try:
                    raw_events = json.loads(events_file.read_text(encoding="utf-8"))
                except (json.JSONDecodeError, OSError):
                    raw_events = []
            if isinstance(raw_events, list):
                picked = pixel_store.events_for(
                    self.settings.pixel_dir, str(pixel_entry["token"]), raw_events
                )
                if picked:
                    opened = {
                        "count": len(picked),
                        "lastAt": picked[-1]["at"],
                        "confidence": "low - an image fetch can be a proxy or a prefetch",
                    }
        verdict["tracked"] = pixel_entry is not None
        verdict["opened"] = opened
        verdict["draftId"] = draft.email_id
        verdict["storeId"] = STORE_ID
        verdict["windowDays"] = scan_days
        return verdict

    async def create_reply_draft(
        self,
        *,
        email_id: str,
        store_id: str,
        body: str,
        reply_all: bool = False,
    ) -> dict[str, object]:
        self._validate_store(store_id)
        if not body.strip():
            raise ValueError("body is required")
        if len(body) > 20_000:
            raise ValueError("body cannot exceed 20000 characters")
        original = await self._get_item_raw(email_id)
        key = change_key(original)
        reference: dict[str, object] = {
            "__type": "ItemId:#Exchange",
            "Id": email_id,
        }
        if key:
            reference["ChangeKey"] = key
        subject = string(original.get("Subject"))
        reply_subject = subject if subject.casefold().startswith("re:") else f"Re: {subject}"
        reply_item = {
            "__type": ("ReplyAllToItem:#Exchange" if reply_all else "ReplyToItem:#Exchange"),
            "ReferenceItemId": reference,
            "ShouldIgnoreChangeKey": True,
            # HTML, not Text: a Text body makes Exchange convert the whole message
            # ("converted from text") and flatten every quoted level into <br> runs,
            # which Outlook mobile cannot fold. With HTML the server appends the
            # quote as native blockquote structure instead.
            "NewBodyContent": {
                "__type": "BodyContentType:#Exchange",
                "BodyType": "HTML",
                "Value": _text_to_html(body),
            },
            "Subject": reply_subject,
        }
        data = await self.api.call(
            "CreateItem",
            {
                "__type": "CreateItemRequest:#Exchange",
                "MessageDisposition": "SaveOnly",
                "SavedItemFolderId": _target_folder("drafts", distinguished=True),
                "Items": [reply_item],
            },
        )
        created = response_items(data)
        draft_id = item_id(created[0]) if created else ""
        if not draft_id:
            raise ApiError(
                "Outlook Web saved no identifiable reply draft.",
                "draft_id_missing",
            )
        try:
            draft = parse_mail_detail(
                await self._get_item_raw(draft_id),
                max_body_characters=1,
            )
            to = draft.to
            cc = draft.cc
            saved_subject = draft.subject or reply_subject
        except ApiError:
            to = ""
            cc = ""
            saved_subject = reply_subject
        return {
            "draftId": draft_id,
            "storeId": STORE_ID,
            "subject": saved_subject,
            "to": to,
            "cc": cc,
            "isReplyAll": reply_all,
            "status": "saved_to_drafts",
        }

    async def list_calendar_events(
        self,
        *,
        starts_after: str,
        ends_before: str,
        max_results: int = 50,
    ) -> dict[str, object]:
        start = _parse_iso(starts_after, "startsAfter")
        end = _parse_iso(ends_before, "endsBefore")
        if end <= start:
            raise ValueError("endsBefore must be later than startsAfter")
        if end - start > timedelta(days=31):
            raise ValueError("Calendar range cannot exceed 31 days")
        if not 1 <= max_results <= 100:
            raise ValueError("maxResults must be between 1 and 100")
        calendar = await self._get_folder(
            "calendar",
            distinguished=True,
            fallback_name="Calendar",
        )
        data = await self.api.call(
            "GetCalendarView",
            {
                "__type": "GetCalendarViewRequest:#Exchange",
                "CalendarId": _target_folder(calendar.folder_id),
                "RangeStart": _calendar_string(start),
                "RangeEnd": _calendar_string(end),
                "ClientSupportsIrm": True,
                "OptimizeExtendedPropertyLoading": True,
            },
            app="Calendar",
        )
        response_body = record(record(data).get("Body"))
        body_items = records(response_body.get("Items"))
        raw_events = body_items or response_items(data) or root_items(data)
        events: list[CalendarEvent] = [parse_calendar_event(raw) for raw in raw_events]
        events.sort(key=lambda item: item.starts_at)
        return {"result": [item.to_dict(STORE_ID) for item in events[:max_results]]}
