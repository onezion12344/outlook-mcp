#!/bin/zsh
# Call the CURRENT code through mcporter without restarting Qoder's MCP process.
# Every invocation spawns a fresh server from the repo binary, so it always runs
# the latest committed code — the stale-process problem, sidestepped.
#
# usage: ./scripts/mcporter_call.sh <tool> [--args '<json>']
#   ./scripts/mcporter_call.sh outlook_auth_status
#   ./scripts/mcporter_call.sh search_index --args '{"query":"round trip","maxResults":3}'
#   MCPORTER_CALL_TIMEOUT=240000 ./scripts/mcporter_call.sh index_sync --args '{"folders":["inbox"],"daysBack":3650}'
set -euo pipefail
REPO="$(cd "$(dirname "$0")/.." && pwd)"
exec mcporter call --stdio "$REPO/.venv/bin/outlook-web-mcp-server" "$@"
