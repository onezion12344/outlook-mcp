# Outlook MCP: Connecting to Outlook without 3rd-party auth

**Unofficial** local MCP server that drives Outlook the same way Outlook Web itself
does — the first-party device-code flow plus Outlook Web's internal JSON API. No
Microsoft Graph, no Azure app registration, no admin consent, no browser cookies,
no desktop COM. Run it locally, connect any MCP client, and your agent can read,
search, index and draft mail — sending only happens with a human-issued one-time
approval for the exact content.

> [!WARNING]
> `https://outlook.cloud.microsoft/owa/service.svc` is **not** a public, supported
> Microsoft API. It can change or be blocked by tenant policy at any time. For
> production integrations prefer Microsoft Graph. This server is stdio-only.

## What it does

| Area | Tools |
|---|---|
| Read & search | `search_emails` (live), `search_index` (local full-text, 2–3 ms), `get_email`, `get_email_html` (links + text), `list_mail_folders`, `list_calendar_events` |
| Local index | `index_sync` (all mail folders), `index_bodies` (full body text on disk), `index_attachment_text` (PDF/Office → text, file stays on the server), `index_status`, `index_verify` |
| Drafts | `create_draft`, `create_reply_draft` (HTML body so mobile clients fold quotes), `send_draft` (nonce-gated), `track_status` (sent / replied / bounced, never "read") |
| Trust & health | `outlook_auth_status`, `pixels_sync`, `outlook-web-mcp doctor` (six local checks, no network, no secrets) |

17 tools, stdio-only. Sending requires a one-time nonce issued by the human in a
terminal (`outlook-web-mcp approve <draftId>`); the digest binds the exact HTML
body, recipients, Bcc and attachments, and expires in 30 minutes.

## Why no third-party auth

Microsoft Graph and most Outlook integrations require an Azure app registration
and tenant admin consent — a wall if you are on a managed account. This fork
instead signs in through **Outlook Web's own first-party client ID** with the
device-code flow, then speaks the same private JSON dialect the web app speaks.
Nothing asks an admin for anything.

## Quickstart

```bash
# clone, then:
uv sync --locked
uv run outlook-web-mcp doctor    # local health check, no network, no secrets
uv run outlook-web-mcp login     # device-code sign-in with Outlook's own client id
```

Register in your MCP client (e.g. `claude mcp add` / `settings.json`):

```json
{
  "mcpServers": {
    "outlook-web": {
      "command": "/absolute/path/to/outlook-web-mcp/.venv/bin/outlook-web-mcp-server",
      "env": { "OUTLOOK_WEB_TIME_ZONE_ID": "China Standard Time" }
    }
  }
}
```

## The send gate

1. Agent calls `create_draft` / `create_reply_draft` — saved to Drafts, never sent.
2. You review it in Outlook, then run `outlook-web-mcp approve <draftId>` in a
   terminal. It prints exactly what you are approving and writes a one-time nonce.
3. Agent calls `send_draft` — it sends only if the nonce matches the **exact**
   content (digest covers HTML body, To/Cc/Bcc, attachment name:size) and is within
   30 minutes. Editing anything invalidates the approval; the nonce is consumed on
   send. **Never run `approve` for an agent, and never paste a nonce into a chat.**

## Security posture

- CA bundle pinned to `certifi`; `trust_env=False` so environment proxies cannot
  hijack token-bearing requests (explicit proxy opt-in only).
- Upstream error text is never echoed to tool results; opt-in logs are 0600.
- Response size cap enforced mid-stream; every response must come from the pinned
  OWA host.
- Stored credentials are refused if the client id changes; verification URLs must
  be a Microsoft sign-in host.
- Full findings ledger: `docs/audit/2026-11-06-security-audit.md` (P2 7/7 closed).

## Docs

- `PITFALLS.md` — 23 hard-won pitfalls of this private dialect (flat recipient
  entries, SendItem 500, ItemClass vs MessageClass, …). Read before hacking.
- `docs/specs/` `docs/specs/` `docs/acceptance-*.md` — dated design/acceptance trail.
- `docs/audit/2026-11-06-security-audit.md` — third-party-style audit + closure ledger.
- `skills/use-outlook-web-mcp/` — an agent-facing skill (MCP copy; runbook for
  JS-rendered mail is screenshot + OCR *outside* the MCP).
- `docs/internal-api.md` — upstream's dialect notes.

## Credits & license

Fork of [ma-nakaya/outlook-web-mcp-server](https://github.com/ma-nakaya/outlook-web-mcp-server)
(see `NOTICE`, upstream pin `aa41b13`). **Upstream code stays MIT; all fork
modifications are licensed under AGPL-3.0** — any hosted modification of this
server must publish its source. See `LICENSE` for the exact boundary.
Maintainer: <owner>. Not affiliated with Microsoft or the upstream author.

## Status

Tested against one real Exchange Online tenant (University) with ~12k messages indexed.
CI matrix configured (ubuntu/macOS/Windows); the collector under
`collector/pixel-worker/` is deployable but **not deployed**, and read-tracking
pixels are therefore off unless you deploy it yourself.
