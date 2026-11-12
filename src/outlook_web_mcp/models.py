from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any


@dataclass(slots=True)
class CredentialBundle:
    refresh_token: str
    client_id: str
    tenant_id: str
    puid: str = ""
    display_name: str = "Outlook User"

    def to_secret_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_secret_dict(cls, value: dict[str, Any]) -> CredentialBundle:
        return cls(
            refresh_token=str(value["refresh_token"]),
            client_id=str(value["client_id"]),
            tenant_id=str(value["tenant_id"]),
            puid=str(value.get("puid") or ""),
            display_name=str(value.get("display_name") or "Outlook User"),
        )


@dataclass(slots=True)
class MailSummary:
    email_id: str
    subject: str
    sender_name: str
    sender_address: str
    received_at: str
    is_unread: bool
    has_attachments: bool
    body_preview: str | None = None
    message_class: str = ""
    internet_message_id: str = ""
    references: str = ""

    def to_dict(self, store_id: str) -> dict[str, object]:
        return {
            "emailId": self.email_id,
            "storeId": store_id,
            "subject": self.subject,
            "senderName": self.sender_name,
            "senderAddress": self.sender_address,
            "receivedAt": self.received_at,
            "isUnread": self.is_unread,
            "hasAttachments": self.has_attachments,
            "messageClass": self.message_class,
            "internetMessageId": self.internet_message_id,
            "references": self.references,
            "bodyPreview": self.body_preview,
        }


@dataclass(slots=True)
class MailDetail:
    email_id: str
    subject: str
    sender_name: str
    sender_address: str
    to: str
    cc: str
    received_at: str
    is_unread: bool
    has_attachments: bool
    body: str
    body_truncated: bool
    message_class: str = ""
    internet_message_id: str = ""
    references: str = ""

    def to_dict(self, store_id: str) -> dict[str, object]:
        return {
            "emailId": self.email_id,
            "storeId": store_id,
            "subject": self.subject,
            "senderName": self.sender_name,
            "senderAddress": self.sender_address,
            "to": self.to,
            "cc": self.cc,
            "receivedAt": self.received_at,
            "isUnread": self.is_unread,
            "hasAttachments": self.has_attachments,
            "messageClass": self.message_class,
            "internetMessageId": self.internet_message_id,
            "references": self.references,
            "body": self.body,
            "bodyTruncated": self.body_truncated,
        }


@dataclass(slots=True)
class MailHtmlDetail:
    email_id: str
    subject: str
    received_at: str
    html: str
    text_content: str
    links: list[dict[str, str]]
    body_truncated: bool

    def to_dict(self, store_id: str) -> dict[str, object]:
        return {
            "emailId": self.email_id,
            "storeId": store_id,
            "subject": self.subject,
            "receivedAt": self.received_at,
            "html": self.html,
            "textContent": self.text_content,
            "links": self.links,
            "bodyTruncated": self.body_truncated,
        }


@dataclass(slots=True)
class FolderInfo:
    folder_id: str
    name: str
    folder_path: str
    parent_folder_id: str | None
    unread_item_count: int
    total_item_count: int
    has_children: bool

    def to_dict(self, store_id: str) -> dict[str, object]:
        return {
            "folderId": self.folder_id,
            "storeId": store_id,
            "name": self.name,
            "folderPath": self.folder_path,
            "parentFolderId": self.parent_folder_id,
            "unreadItemCount": self.unread_item_count,
            "totalItemCount": self.total_item_count,
            "hasChildren": self.has_children,
        }


@dataclass(slots=True)
class CalendarEvent:
    event_id: str
    subject: str
    starts_at: str
    ends_at: str
    location: str
    organizer: str
    is_all_day: bool
    is_recurring: bool
    busy_status: str

    def to_dict(self, store_id: str) -> dict[str, object]:
        return {
            "eventId": self.event_id,
            "storeId": store_id,
            "subject": self.subject,
            "startsAt": self.starts_at,
            "endsAt": self.ends_at,
            "location": self.location,
            "organizer": self.organizer,
            "isAllDay": self.is_all_day,
            "isRecurring": self.is_recurring,
            "busyStatus": self.busy_status,
        }
