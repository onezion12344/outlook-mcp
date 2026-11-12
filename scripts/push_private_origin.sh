#!/bin/zsh
# Create origin for the private fork and push. Refuses to touch `upstream`.
#
# OWNER MUST SET the remote URL before running:
#   OWNER_REPO=git@github.com:<you>/outlook-web-mcp.git ./scripts/push_private_origin.sh
set -euo pipefail

OWNER_REPO="${OWNER_REPO:-}"
if [[ -z "$OWNER_REPO" ]]; then
  print "Set OWNER_REPO=<private repo url> first. Refusing to guess." >&2
  exit 1
fi
if [[ "$OWNER_REPO" == *"ma-nakaya/"* ]]; then
  print "That URL is the upstream repo. Refusing." >&2
  exit 1
fi

cd "$(dirname "$0")/.."

if git remote get-url origin >/dev/null 2>&1; then
  current=$(git remote get-url origin)
  if [[ "$current" != "$OWNER_REPO" ]]; then
    print "origin already exists and differs: $current" >&2
    exit 1
  fi
else
  git remote add origin "$OWNER_REPO"
fi

# Prove we are pushing to origin, and that upstream is untouched.
print "remotes before push:"; git remote -v
git push -u origin "$(git branch --show-current)"
print "pushed to origin only. upstream was never an argument."
