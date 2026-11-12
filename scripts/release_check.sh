#!/usr/bin/env bash
# Release/regression gate for the outlook-web fork.
#
# Runs every check in the order CI does, then the local-only doctor and the live
# smoke test (which needs a signed-in mailbox). Exits non-zero on the first failure.
# Usage: scripts/release_check.sh [--offline]
set -uo pipefail
cd "$(dirname "$0")/.." || exit 1

FAILED=0
step() {
  local name="$1"
  shift
  printf '%-28s' "$name"
  mkdir -p logs
  if "$@" >logs/release-check.log 2>&1; then
    echo "PASS"
  else
    echo "FAIL"
    tail -15 logs/release-check.log
    FAILED=1
  fi
}

step "ruff check" uv run --no-sync ruff check .
step "ruff format" uv run --no-sync ruff format --check .
step "mypy" uv run --no-sync mypy
step "pytest" uv run --no-sync pytest -q
step "doctor" .venv/bin/outlook-web-mcp doctor

if [ "${1:-}" != "--offline" ]; then
  step "live smoke" env OUTLOOK_WEB_TIME_ZONE_ID="China Standard Time" \
    uv run --no-sync python scripts/live_smoke.py
else
  printf '%-28s%s\n' "live smoke" "SKIPPED (--offline)"
fi

rm -f logs/release-check.log
if [ "$FAILED" -ne 0 ]; then
  echo "release check: FAILED"
  exit 1
fi
echo "release check: PASSED"
