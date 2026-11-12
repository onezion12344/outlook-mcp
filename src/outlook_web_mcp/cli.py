from __future__ import annotations

import argparse
import asyncio
import json
import sys
import time
from pathlib import Path

from . import doctor, pixel_sync
from .api import OutlookWebApi
from .approval import approve
from .auth import AuthManager, DeviceCodePrompt, PendingDeviceLogin
from .config import Settings
from .credentials import KeyringPendingLoginStore
from .errors import OutlookWebMcpError
from .service import OutlookService


def _json(value: object) -> None:
    print(json.dumps(value, ensure_ascii=False, indent=2))


def _show_code(prompt: DeviceCodePrompt) -> None:
    print(f"Open: {prompt.verification_uri}")
    print(f"Code: {prompt.user_code}")


async def _run(args: argparse.Namespace) -> int:
    settings = Settings()
    pending_store = KeyringPendingLoginStore(settings.credential_profile)
    async with AuthManager(settings) as auth:
        if args.command == "login":
            bundle = await auth.login(
                open_browser=not args.no_browser,
                on_code=_show_code,
            )
            _json(
                {
                    "authenticated": True,
                    "profile": settings.credential_profile,
                    "tenant_id": bundle.tenant_id,
                    "display_name": bundle.display_name,
                }
            )
            return 0
        if args.command == "login-start":
            pending = await auth.begin_login(open_browser=not args.no_browser)
            pending_store.save(pending.to_secret_dict())
            _json(
                {
                    "verification_uri": pending.prompt.verification_uri,
                    "verification_uri_complete": pending.prompt.verification_uri_complete,
                    "user_code": pending.prompt.user_code,
                    "expires_in": pending.prompt.expires_in,
                }
            )
            return 0
        if args.command == "login-complete":
            raw = pending_store.load()
            if raw is None:
                raise OutlookWebMcpError(
                    "No pending login. Run `outlook-web-mcp login-start` first.",
                    "pending_missing",
                )
            pending = PendingDeviceLogin.from_secret_dict(raw)
            if pending.expires_at <= time.time():
                # An interrupted login leaves a redeemable code in the Keychain;
                # drop it as soon as we look at it rather than keeping it forever.
                pending_store.delete()
                raise OutlookWebMcpError(
                    "The stored device code expired. Run `outlook-web-mcp login-start` again.",
                    "device_code_expired",
                )
            try:
                bundle = await auth.complete_login(pending)
            finally:
                pending_store.delete()
            _json(
                {
                    "authenticated": True,
                    "profile": settings.credential_profile,
                    "tenant_id": bundle.tenant_id,
                    "display_name": bundle.display_name,
                }
            )
            return 0
        if args.command == "approve":
            async with OutlookWebApi(auth, settings) as api:
                service = OutlookService(auth, api, settings)
                info = await service.approval_digest(args.draft_id)
            path = approve(settings.approval_dir, str(info["digest"]))
            _json(
                {
                    "approved": True,
                    "draftId": info["draftId"],
                    "digest": info["digest"],
                    "subject": info["subject"],
                    "to": info["to"],
                    "cc": info["cc"],
                    "bcc": info.get("bcc", ""),
                    "attachments": info.get("attachments", []),
                    "nonce": str(path),
                    "expiresInMinutes": 30,
                    "covers": "exactly this content; editing the draft invalidates it",
                }
            )
            return 0
        if args.command == "ingest-ocr":
            # Local-only write-back: the screenshot+OCR runbook produces a text file,
            # this makes that mail searchable. No network, no renderer in the MCP.
            from . import index as index_store

            text = Path(args.text_file).read_text(encoding="utf-8")
            if not text.strip():
                raise OutlookWebMcpError("The OCR text file is empty.", "ocr_text_empty")
            connection = index_store.open_index(settings.index_path)
            try:
                exists = connection.execute(
                    "SELECT 1 FROM items WHERE email_id = ?", (args.email_id,)
                ).fetchone()
                if exists is None:
                    raise OutlookWebMcpError(
                        "That emailId is not in the index; run index_sync first.",
                        "email_not_indexed",
                    )
                index_store.store_body(connection, args.email_id, body_text=text, source="ocr")
            finally:
                connection.close()
            _json({"emailId": args.email_id, "source": "ocr", "characters": len(text)})
            return 0
        if args.command == "doctor":
            report = doctor.summarise(doctor.run_checks(settings))
            _json(report)
            return 0 if report["ok"] else 1
        if args.command == "pixels-sync":
            _json(await pixel_sync.sync_events(settings, since=args.since or ""))
            return 0
        if args.command == "status":
            _json(await auth.ensure_status())
            return 0
        if args.command == "logout":
            auth.logout()
            pending_store.delete()
            _json({"authenticated": False, "profile": settings.credential_profile})
            return 0
        if args.command == "probe":
            async with OutlookWebApi(auth, settings) as api:
                data = await api.call(
                    "GetFolder",
                    {
                        "__type": "GetFolderRequest:#Exchange",
                        "FolderShape": {
                            "__type": "FolderResponseShape:#Exchange",
                            "BaseShape": "IdOnly",
                        },
                        "FolderIds": [
                            {
                                "__type": "DistinguishedFolderId:#Exchange",
                                "Id": "inbox",
                            }
                        ],
                    },
                )
            _json(
                {
                    "success": True,
                    "response_has_body": isinstance(data.get("Body"), dict),
                }
            )
            return 0
    return 2


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(
        prog="outlook-web-mcp",
        description="Unofficial Outlook Web MCP server without Microsoft Graph or COM.",
    )
    sub = root.add_subparsers(dest="command", required=True)
    login = sub.add_parser("login", help="Sign in with Microsoft's device-code flow")
    login.add_argument("--no-browser", action="store_true")
    login_start = sub.add_parser(
        "login-start",
        help="Start sign-in and securely save the short-lived device code",
    )
    login_start.add_argument("--no-browser", action="store_true")
    sub.add_parser("login-complete", help="Complete a previously started sign-in")
    approve_cmd = sub.add_parser(
        "approve",
        help="Authorise one exact draft for sending (one-time token)",
    )
    approve_cmd.add_argument("draft_id")
    ingest = sub.add_parser(
        "ingest-ocr",
        help="Store OCR'd text for a JS-rendered mail so it becomes searchable (local only)",
    )
    ingest.add_argument("email_id")
    ingest.add_argument("text_file")
    sub.add_parser("doctor", help="Local-only health check (no network, no secrets)")
    pixels = sub.add_parser(
        "pixels-sync",
        help="Pull read-tracking events from the deployed collector",
    )
    pixels.add_argument(
        "--since", default="", help="ISO timestamp; defaults to the newest stored event"
    )
    sub.add_parser("status", help="Check and refresh authentication")
    sub.add_parser("logout", help="Remove the saved refresh token")
    sub.add_parser("probe", help="Run a content-free Inbox API health check")
    sub.add_parser("serve", help="Run the stdio MCP server")
    return root


def main() -> None:
    args = parser().parse_args()
    if args.command == "serve":
        from .server import main as server_main

        server_main()
        return
    try:
        raise SystemExit(asyncio.run(_run(args)))
    except OutlookWebMcpError as exc:
        print(f"{exc.code}: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc
    except ValueError as exc:
        print(f"invalid_input: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
