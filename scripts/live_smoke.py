"""Read-only live check against the owner's real mailbox. Prints counts only.

Never prints subjects, senders or bodies: this script runs against a live University
mailbox and its stdout ends up in agent context and logs. Exit codes:
0 = all probes answered, 2 = not signed in, 3 = local health check failed, 1 = a probe raised.
"""

from __future__ import annotations

import asyncio
import sys

from outlook_web_mcp import doctor
from outlook_web_mcp.api import OutlookWebApi
from outlook_web_mcp.auth import AuthManager
from outlook_web_mcp.config import Settings
from outlook_web_mcp.service import OutlookService


async def main() -> int:
    settings = Settings()
    auth = AuthManager(settings)
    api = OutlookWebApi(auth, settings)
    service = OutlookService(auth, api, settings)

    health = doctor.summarise(doctor.run_checks(settings))
    counts = health["counts"]
    print(
        "doctor:",
        f"ok={counts['ok']} warn={counts['warn']} fail={counts['fail']}",
        "|",
        ", ".join(f"{item['name']}={item['status']}" for item in health["checks"]),
    )
    if not health["ok"]:
        for item in health["checks"]:
            if item["status"] == "fail":
                print(f"  FAIL {item['name']}: {item['detail']}", file=sys.stderr)
        await api.aclose()
        await auth.aclose()
        return 3

    status = await service.auth_status()
    print("auth:", status.get("authenticated"), status.get("usable"))
    if not status.get("usable"):
        print("NOT SIGNED IN — run: uv run outlook-web-mcp login", file=sys.stderr)
        return 2

    failures = 0

    try:
        folders = await service.list_mail_folders(folder="root", recursive=True, max_results=200)
        names = sorted({str(row["name"]) for row in folders["result"]})
        print("folders:", len(names), names[:12])
    except Exception as exc:
        failures += 1
        print("folders: FAILED", type(exc).__name__, getattr(exc, "code", ""))

    for label, kwargs in [
        ("inbox30d", {"folder": "inbox", "days_back": 30, "max_results": 200}),
        ("archive30d", {"folder": "archive", "days_back": 30, "max_results": 50}),
        (
            "query-university-365d",
            {"folder": "inbox", "query": "university", "days_back": 365, "max_results": 200},
        ),
    ]:
        try:
            out = await service.search_emails(**kwargs)
            meta = out["meta"]
            print(
                label,
                "returned",
                meta["returnedCount"],
                "scanned",
                meta["scannedItems"],
                "folders",
                meta["scannedFolders"],
                "truncated",
                meta["truncated"],
                meta["truncatedReason"],
                "hitCap",
                meta["hitCapReached"],
                "continuation",
                meta.get("continuation") is not None,
            )
        except Exception as exc:
            failures += 1
            print(label, "FAILED", type(exc).__name__, str(exc)[:120])

    try:
        events = await service.list_calendar_events(
            starts_after="2026-11-06T00:00:00+08:00",
            ends_before="2026-12-06T00:00:00+08:00",
            max_results=100,
        )
        print("calendar31d:", len(events["result"]))
    except Exception as exc:
        failures += 1
        print("calendar31d FAILED", type(exc).__name__, str(exc)[:120])

    await api.aclose()
    await auth.aclose()
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
