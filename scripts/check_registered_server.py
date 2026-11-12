"""Check the MCP binary registered in ~/.qoder-cn/settings.json runs current code.

Verifies that the registered server exposes exactly the shipped tool surface.
A stale long-running process (or a stale install path) is the failure mode that
made the recipient-loss bug survive a fixed codebase.
"""

from __future__ import annotations

import asyncio
import json
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "src"))

from mcp import ClientSession, StdioServerParameters  # noqa: E402
from mcp.client.stdio import stdio_client  # noqa: E402

SETTINGS = Path.home() / ".qoder-cn" / "settings.json"
EXPECTED_TOOLS = {
    "outlook_auth_status",
    "search_emails",
    "list_mail_folders",
    "get_email",
    "set_email_read_state",
    "list_calendar_events",
    "create_draft",
    "create_reply_draft",
    "send_draft",
    "track_status",
    "index_sync",
    "search_index",
    "index_status",
    "index_verify",
    "pixels_sync",
}


def registered_command() -> tuple[str, dict[str, str]]:
    config = json.loads(SETTINGS.read_text())
    entry = config["mcpServers"]["outlook-web"]
    return entry["command"], entry.get("env", {})


async def main() -> int:
    command, env = registered_command()
    binary = Path(command)
    if not binary.exists():
        print(f"FAIL: registered binary missing: {command}")
        return 1
    params = StdioServerParameters(
        command=command,
        env={**env, "PATH": "/usr/bin:/bin:/opt/homebrew/bin", "HOME": str(Path.home())},
    )
    async with stdio_client(params) as (read, write), ClientSession(read, write) as session:
        await session.initialize()
        tools = {t.name for t in (await session.list_tools()).tools}
    missing = EXPECTED_TOOLS - tools
    extra = tools - EXPECTED_TOOLS
    if missing or extra:
        print(f"FAIL: tool surface mismatch. missing={sorted(missing)} extra={sorted(extra)}")
        return 1
    print(f"OK: registered binary serves all {len(EXPECTED_TOOLS)} expected tools")
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
