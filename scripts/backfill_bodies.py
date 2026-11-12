"""Backfill full body text into the local index, newest first.

Resumable: every batch commits, so Ctrl-C (or a crash) loses at most one batch.
Run: ./.venv/bin/python scripts/backfill_bodies.py [--batch 200] [--sleep 0.5]
"""

from __future__ import annotations

import argparse
import asyncio
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

from outlook_web_mcp.api import OutlookWebApi  # noqa: E402
from outlook_web_mcp.auth import AuthManager  # noqa: E402
from outlook_web_mcp.config import Settings  # noqa: E402
from outlook_web_mcp.service import OutlookService  # noqa: E402


async def main(batch: int, sleep: float, max_batches: int, concurrency: int) -> int:
    settings = Settings()
    auth = AuthManager(settings)
    service = OutlookService(auth, OutlookWebApi(auth, settings), settings)
    done = 0
    for number in range(1, max_batches + 1):
        result = await service.index_bodies(max_fetches=batch, concurrency=concurrency)
        done += int(result["fetched"])
        print(
            f"batch {number}: fetched={result['fetched']} failed={result['failed']} "
            f"remaining={result['remaining']} chars={result['bodies']['chars']}",
            flush=True,
        )
        if int(result["remaining"]) <= 0:
            print(f"complete: {done} messages fetched in total", flush=True)
            return 0
        await asyncio.sleep(sleep)
    print(f"stopped after {max_batches} batches; {done} fetched this run", flush=True)
    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--batch", type=int, default=400)
    parser.add_argument("--concurrency", type=int, default=8)
    parser.add_argument("--sleep", type=float, default=0.5)
    parser.add_argument("--max-batches", type=int, default=10_000)
    args = parser.parse_args()
    raise SystemExit(asyncio.run(main(args.batch, args.sleep, args.max_batches, args.concurrency)))
