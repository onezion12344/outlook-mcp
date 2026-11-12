# open-pixel collector — **NOT DEPLOYED**

Read-tracking pixel for the outlook-web fork. This directory is a *deployable artifact*; nothing
here has been published. Deploying it is an outward-facing action and remains the owner's call.

**Locally verified (2026-11-06)**: `tests/test_pixel_worker_contract.py` loads this exact
`worker.js`, backs `env.DB` with `node:sqlite` running the schema from the deploy section below,
and drives it with real HTTP: the GIF + `no-store` header, 404 on a malformed token, 401 without
the sync token, `since` filtering, and a full pull through `pixel_sync.sync_events` into
`pixels/events.json` with attribution back to the draft. A mutated `cache-control` turns that
suite red. What is *not* verified is Cloudflare's own D1 and routing — that needs the deploy.

## What it records

Per fetch of `GET /o/<32 hex>.gif`:

| Field | Content |
|---|---|
| `token` | the per-draft random token (no content, no recipient) |
| `at` | UTC timestamp of the fetch |
| `ua` | first 120 chars of the User-Agent |
| `ip4` | **4-byte truncated SHA-256 of the client IP** — never the address |

Nothing else. No email body, no recipient address, no mail id, no geolocation.

## What it can and cannot tell you

- **Can**: an image tied to this token was fetched, and when.
- **Cannot**: that a human read the message. Gmail proxies images (open time is still visible, but
  the IP is Google's); Apple Mail Privacy Protection **prefetches** → false positives; clients that
  block remote images → false negatives. Treat every event as *low confidence*.
- This is why `track_status` reports it as `opened (image fetched)` with
  `confidence: "low - an image fetch can be a proxy or a prefetch"` and never as "read".

## Deploy (owner's decision)

```bash
cd collector/pixel-worker
npx wrangler d1 create open-pixel            # put the returned id into wrangler.toml
npx wrangler d1 execute open-pixel --remote --command \
  "CREATE TABLE IF NOT EXISTS events (id INTEGER PRIMARY KEY AUTOINCREMENT, token TEXT NOT NULL, at TEXT NOT NULL, ua TEXT NOT NULL DEFAULT '', ip4 TEXT NOT NULL DEFAULT '')"
npx wrangler d1 execute open-pixel --remote --command \
  "CREATE INDEX IF NOT EXISTS events_token_at ON events(token, at)"
npx wrangler secret put SYNC_TOKEN           # any long random string
npx wrangler deploy --route pixel.example.com/*
```

## Wire the fork to it

Add to the MCP server env in `~/.qoder-cn/settings.json`:

```json
"OUTLOOK_WEB_PIXEL_BASE": "https://pixel.example.com"
```

`OUTLOOK_WEB_TRACK_PIXEL=0` disables tracking globally; `create_draft(track=false)` disables it per
message. Plain-text drafts never carry a pixel (they cannot), and say so in `trackingWarning`.

## Pulling events back

`GET /events?since=<iso>` returns `[{token, at, ua}]` and requires `Authorization: Bearer <SYNC_TOKEN>`.
The local side pulls them with `outlook-web-mcp pixels-sync` (or the `pixels_sync` MCP tool),
which stores them at `~/Library/Caches/outlook-web-mcp/pixels/events.json` (0600) and only
asks for events newer than the newest stored one. `track_status` then attributes events by token.

Both `OUTLOOK_WEB_PIXEL_BASE` and `OUTLOOK_WEB_PIXEL_SYNC_TOKEN` must be set; the sync client uses
the same posture as the mail client (pinned certifi trust root, `trust_env=False`, no redirects).

## Ethics

Every recipient's mail client can see that the message contains a remote image, and institutional
administrators can see the URL. For official correspondence (University offices, scholarship committees,
visa authorities) the owner should normally keep `track=false`; the tool does not decide that.
