# Security audit — outlook-web-mcp-server v0.1.0 (ma-nakaya)

Audited tree: `<local checkout>` (unpacked tarball, **no `.git`** — cannot be diffed against GitHub `main`).
Date: 2026-11-06. Method: read-only source audit + offline test run + non-mutating runtime probes. No auth command was run; no keyring secret value was read.

---

## 1. Verdict

**Safe with required changes.** Nothing in this codebase is malicious: the only network destinations in `src/` are Microsoft (Section 2), there is no telemetry/update-check/Sentry, no `eval/exec/subprocess/pickle/yaml`, no attachment download and no payload-bearing file writes, and the tool surface is honestly the 7 advertised names. Three P1s must be dealt with before pointing it at an University mailbox, worst first:

1. **The bearer token and refresh token travel through whatever proxy the environment/system tells httpx to use — and on this machine that is happening right now.** Both HTTP clients are built without `trust_env=False` (`api.py:113-117`, `auth.py:131-135`); httpx defaults `trust_env=True` (`httpx/_client.py:648`, applied at `:685-686` → `get_environment_proxies()`), and the installed `_utils.py:30-42` uses `urllib.request.getproxies()`, which on macOS also reads **System Configuration proxies**. Live probe in the repo venv returned `{'http://': 'http://127.0.0.1:7897', 'https://': 'http://127.0.0.1:7897'}`. `SSL_CERT_FILE`/`SSL_CERT_DIR` are likewise honored (`httpx/_config.py:34-37`). One env var (or one macOS proxy setting) plus a CA override turns "fixed host" into full disclosure of the access token, the refresh-token POST body, and every request body (search queries, item IDs, reply-draft text). The README's claim of never forwarding the bearer off the fixed host is therefore only true for the *URL*, not the *transport*.
2. **The credential's real capability is far larger than the tool surface.** Scope is `https://outlook.office.com/.default openid profile offline_access` (`config.py:10`) and the code *demands* `OWA.AccessAsUser.All` or `Mail.ReadWrite` (`auth.py:314-319`). On a university Exchange Online tenant that is full mailbox **read-write** + calendar read-write, plus an indefinitely-refreshing `offline_access` token. "No send/delete/move/archive tool" (`server.py:52-159`, `test_server.py:12-21`) is a client-side restraint, not a permission boundary — anyone holding the keychain item or the token can send, delete and read everything. There is no local way to narrow a first-party `.default` scope.
3. **Uncontrolled upstream text is returned into the agent's context via tool errors.** `_first_error()` copies Exchange `MessageText`/`FaultMessage` verbatim (truncated to 500) into `ApiError` (`api.py:60-80`, raised at `api.py:179-184` and `api.py:186-192`), and `_aad_error()` copies Microsoft's `error_description` (`auth.py:93-98`). FastMCP renders the exception message as tool-result text (`mcp/server/lowlevel/server.py:472-479`). This directly contradicts docs/internal-api.md's "API エラーにトークン、リクエスト本文、メール内容を含めない". It is not token-bearing (`test_api.py:110` proves `"Bearer"` is absent), but it is an extra, unmodelled channel for attacker-influenced strings to reach the model.

Everything else is P2/P3 and listed in Section 3.

---

## 2. Attack surface map

### 2.1 Outbound destinations reachable from `src/outlook_web_mcp/`

| Destination | Constructed at | Carries | Env-overridable? |
|---|---|---|---|
| `https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode` | `auth.py:20` + `auth.py:139-140`, called `auth.py:147-150` | `client_id`, `scope` | No (`AUTHORITY` is a module constant) |
| `https://login.microsoftonline.com/<tenant>/oauth2/v2.0/token` | `auth.py:143-144`, called `auth.py:250-257` (device) and `auth.py:280-288` (refresh) | **device_code**, **refresh_token**, client_id | No. `tenant_id` is interpolated but cannot change the host (path position) |
| `https://outlook.cloud.microsoft/owa/service.svc?action=<Action>&app=<Mail\|Calendar>` | `config.py:11-12` → `api.py:214` | **Bearer token** + full JSON request body (as body, or as `X-OWA-UrlPostData` header `api.py:206-211`) | **No.** `OWA_ORIGIN`/`OWA_ENDPOINT` are constants, not env — the "swap the origin via env" attack does not exist |
| `https://login.microsoft.com/device` | `auth.py:159` fallback string only | never fetched by code; only printed (`cli.py:20-21`) and opened (`auth.py:195`) | n/a |

Non-Microsoft destinations found in `src/`: **none**. Establishing commands in Section 7 (`grep -rnoE "https?://..." src/`, repo-wide telemetry keyword grep, and the `import` inventory — the only network module is `httpx`).

**Indirect egress** not visible in the table above: the proxy configured by `HTTP(S)_PROXY`/`ALL_PROXY` *or macOS system proxy settings*, plus a CA override from `SSL_CERT_FILE`/`SSL_CERT_DIR` — see P1-1. Additionally `mcp.json` launches `uv run --locked`, so first launch can hit `pypi.org` to populate `${PLUGIN_DATA}/venv` (integrity is pinned by `uv.lock`).

### 2.2 Secrets at rest / in transit

| What | Where | Encrypted? | Evidence |
|---|---|---|---|
| Refresh token | macOS Keychain, split into ≤900-char chunks, service `ma-nakaya/outlook-web-mcp-server`, account `<profile>:<generation>:refresh:<i>` | Yes (Keychain; ACL-gated, not access-prompt-gated) | `credentials.py:13-15,85-92,148-174`; item observed live: `"acct"="default:1df325c6…:refresh:0"`, `"svce"="ma-nakaya/outlook-web-mcp-server"` |
| Credential metadata (`tenant_id`, `client_id`, `puid`, `display_name`) | same Keychain service, account `<profile>:metadata` (refresh token explicitly popped before write) | Yes | `credentials.py:32-34,161-174`; `models.py:15-16` |
| **device_code** (short-lived secret) | Keychain service `ma-nakaya/outlook-web-mcp-server/pending-login`, account `<profile>` — deleted only by `login-complete`/`logout` | Yes | `credentials.py:195-221`, `auth.py:41-50`, `cli.py:44`, `cli.py:65`, `cli.py:80` |
| **Access token** | process memory only (`AuthManager._cached`) | n/a — never persisted. **Confirmed**: the only writes are `auth.py:236` and `auth.py:357`; there is no serialization path | `auth.py:76-79,114,236,357` |
| Cache dir contents | `~/Library/Caches/outlook-web-mcp-server/*.lock` — **two 0-byte files, no payloads** (`auth-default.lock`, `auth-refresh-test.lock`; the second is created by the test suite) | n/a | `config.py:61-64`; verified with `ls`/`find` |
| stdout / MCP responses | no token ever printed. `cli.py` prints only `verification_uri`, `user_code`, `tenant_id`, `display_name`; server logs go to **stderr** only | — | `cli.py:15-21,33-40,45-52,101-106`, `server.py:163-167` |
| MCP responses (leak surface) | `outlook_auth_status` returns `tenant_id` + `display_name`; tool errors return upstream `MessageText`/`error_description` | — | `auth.py:369-381`, `api.py:65-70` |

### 2.3 Tool surface and every `action=` the code can send

| MCP tool | Effect | Annotations | Underlying `action=` |
|---|---|---|---|
| `outlook_auth_status` | read; **can trigger a token refresh** | readOnly | — (auth only) |
| `search_emails` | read | readOnly | `GetFolder`, `FindFolder`, `FindItem` |
| `list_mail_folders` | read | readOnly | `GetFolder`, `FindFolder` |
| `get_email` | read | readOnly | `GetItem` |
| `list_calendar_events` | read | readOnly | `GetFolder`, `GetCalendarView` (`app=Calendar`) |
| `set_email_read_state` | **write** (one flag) | stateChange | `GetItem`, `UpdateItem` |
| `create_reply_draft` | **write** (adds draft) | additiveWrite | `GetItem`, `CreateItem`, `GetItem` |

Complete action set (7): `CreateItem`, `FindFolder`, `FindItem`, `GetCalendarView`, `GetFolder`, `GetItem`, `UpdateItem` — enumerated at `service.py:112,137,245,403,452,520,578` and `cli.py:85`. There is **no generic passthrough tool**, no raw-payload/`graph-batch`-style surface, and `action` is never caller-supplied (it is a literal at every call site; `api.py:214` and the `Action` header at `api.py:231` only ever receive those literals). `MessageDisposition` is hard-coded `"SaveOnly"` in exactly two places (`service.py:473`, `service.py:524`) and no tool parameter maps to it — `create_reply_draft` cannot be steered to send. `grep -rn "Send"` over `src/*.py` returns **no send action** at all.

---

## 3. Findings

### P1-1 — Proxy/CA from the environment hijacks token-bearing requests (config + code)
**Evidence:** `api.py:113-117`, `auth.py:131-135` construct `httpx.AsyncClient(timeout=…, follow_redirects=False, headers=…)` — no `trust_env=False`. httpx `_client.py:648` (`trust_env: bool = True`), `_client.py:685-686` (`allow_env_proxies = trust_env and transport is None`), `_utils.py:30-42` (`urllib.request.getproxies()`, which on macOS reads System Preferences proxies), `_config.py:34-37` (`SSL_CERT_FILE`, `SSL_CERT_DIR`).
**Runtime proof:** inside the repo's own venv, `httpx._utils.get_environment_proxies()` → `{'http://': 'http://127.0.0.1:7897', 'https://': 'http://127.0.0.1:7897'}`. So today the OWA POST and the AAD refresh POST (which contains the **refresh token**) are already being CONNECTed through a local proxy a rule engine can move anywhere.
**Why it matters for a student mailbox on a university tenant:** campus VPN/proxy tooling, Clash-class rule engines, and capture tools (mitmproxy/Charles/Proxyman) are routinely present on student machines, and MCP servers inherit whatever env the launching agent gives them. A single `HTTPS_PROXY` + `SSL_CERT_FILE` pair yields the refresh token, which on an un-protected tenant means persistent read-write mail and calendar, surviving password changes until admin revocation.
**Fix (code, 2 lines):** add `trust_env=False` to both client constructors. **Fix (config, zero-code):** launch the server with the proxy/CA vars scrubbed, e.g. `env -u HTTP_PROXY -u HTTPS_PROXY -u ALL_PROXY -u SSL_CERT_FILE -u SSL_CERT_DIR uv run …`, and confirm macOS system proxy is off for the account running the server.

### P1-2 — Token capability ≫ tool surface: `.default` grants read-write mailbox + `offline_access`
**Evidence:** `config.py:9-10` (scope), `auth.py:314-319` (code *rejects* tokens lacking `OWA.AccessAsUser.All`/`Mail.ReadWrite`), `auth.py:280-288` (refresh keeps working), `credentials.py:148-174` (refresh token durable in Keychain).
**Why it matters:** University `connect.example.edu` addresses are used for scholarships, registry, visas, supervisor and job correspondence. A leaked/abused refresh token gives an attacker the ability to *send as you* and *delete your evidence*, not merely read. Any prompt-injection that reaches an agent holding this token is therefore a mailbox-takeover primitive, even though the server itself exposes no send tool. Also note the client ID is a **first-party Microsoft app** (`config.py:8`) — sign-ins from it will look legitimate in tenant logs, so this is hard to spot in AAD reports.
**Fix:** cannot be narrowed locally (`.default` ignores narrower scope requests). Mitigate operationally: dedicated macOS user (separate Keychain) or at minimum treat this Keychain item as a crown-jewel secret; keep the MCP server off any HTTP transport (it is stdio-only: `server.py:169`, `mcp.json` `"type": "stdio"` — keep it that way); monitor University/AAD sign-in logs; revoke via password change + consent revocation if abused.

### P1-3 — Upstream error text is echoed into MCP tool results (code)
**Evidence:** `api.py:60-80` returns `MessageText`/`FaultMessage`/`error_description` verbatim (500-char truncation); `api.py:176-192` interpolates it into `ApiError(...)`; `auth.py:93-98` does the same for AAD; `server.py:52-159` never catches, and FastMCP converts the exception message into `TextContent` with `isError=True` (`mcp/server/lowlevel/server.py:472-479`).
**Why it matters:** Exchange error strings frequently embed the offending identifier or recipient text; a mailbox-content-derived string reaching the model through an *error* path bypasses the "treat mailbox content as untrusted data" framing in `server.py:40-48` because the agent sees it as a failure message, not as data.
**Fix (code):** return only the machine code (`response_code`, `http_status`) to callers; log the upstream text to stderr at DEBUG. One change at `api.py:179-192` / `auth.py:154-156`.

### P2-1 — Response-size cap is applied after the whole body is buffered (code)
`api.py:155` checks `len(response.content) > max_response_bytes`, but `api.py:215` uses `client.post()` which has already read the body into RAM. `docs/internal-api.md` sells this as a size guardrail. A hostile/compromised upstream can still force an arbitrarily large allocation. **Fix:** `client.stream("POST", …)` + abort past the cap. Low practical impact (the response is TLS-authenticated to Microsoft), hence P2.

### P2-2 — Keyring gate tests `priority > 0`, not "is this a secure store" (code)
`credentials.py:37-50`. Measured on this machine with the repo's own venv (keyring **25.7.0**): `keyring.get_keyring()` → `keyring.backends.macOS.Keyring`, `priority = 5`, and `is_secure()` **does not exist** (`AttributeError`) — so the precise API the author would need is gone in keyring 25 and priority is the only available proxy. Registered backends and priorities: `fail.Keyring = 0` (correctly refused), `chainer.ChainerBackend = -1` (would be refused even though it can route to the real Keychain — a false negative that breaks legitimate use if chainer is ever selected), `macOS.Keyring = 5` (correctly accepted). `keyring` also honors `~/.config/python_keyring/keyringrc.cfg` `default-keyring`, so a local config file selects the backend and only its own priority gates it. **UNVERIFIED:** what priorities third-party plaintext backends declare — `keyrings.alt`, `keyrings.cryptfile`, `keyring_realm` are all **absent** from this venv (probe), and the repo declares no credential backend beyond `keyring`, so no plaintext store with `priority > 0` is reachable *as installed*. **Fix (code):** replace the priority test with an explicit allowlist of `type(backend).__module__` (`keyring.backends.macOS`, `keyring.backends.Windows`) and refuse otherwise.

### P2-3 — The credential store has no tests, and CI never runs macOS (tests/CI)
`pyproject.toml` testpaths `tests/`; the suite is `test_api.py`, `test_auth.py`, `test_server.py`, `test_service.py` — **no `test_credentials.py`** exists. `ci.yml` runs only `jobs: test-windows` on `windows-latest`. Tests inject `MemoryStore` (`test_auth.py:24-38`, `test_api.py:21-33`) and `MockTransport`, so `_ensure_secure_backend`, chunked save/load, rotation cleanup (`credentials.py:180-181`) and the macOS Keychain path are never executed. Also untested: redirect refusal (`api.py:148-154`), size cap, `follow_redirects=False`, `trust_env`. **Fix:** add a credentials round-trip test behind a fake backend + a macOS CI runner.

### P2-4 — `OUTLOOK_WEB_CLIENT_ID` re-binds the stored credential (config, low likelihood)
`config.py:33-37` makes the client ID env-overridable; it is sent on the device-code POST (`auth.py:149`), persisted into the Keychain metadata (`credentials.py:161`, `models.py:15-16`) and reused for every refresh (`auth.py:282-283`). The authority is *not* overridable, so this cannot redirect the token to a non-Microsoft host; but anyone with env write access can rotate the stored refresh token to a different app registration. `OWA_ORIGIN`/`OWA_ENDPOINT` (`config.py:11-12`) are **not** env-readable — that specific swap is impossible. **Fix (code):** drop the env override, or refuse to reuse a stored bundle whose `client_id` differs from the active one.

### P2-5 — `tenant_id` and `display_name` are returned to the agent (code)
`auth.py:369-381` → exposed via `outlook_auth_status` (`server.py:52-55`). PUID is *not* returned (good). Tenant GUID + display name in every transcript is a small, needless identifier disclosure into third-party model context. **Fix:** return only `authenticated`/`usable`/expiry.

### P2-6 — Abandoned `device_code` lingers in the Keychain (code)
`cli.py:44` writes the pending login (`auth.py:41-50`, includes `device_code`); deletion happens only in `login-complete` (`cli.py:65`) or `logout` (`cli.py:80`). An interrupted `login-start` leaves a redeemable device code stored indefinitely (Microsoft-side expiry is not enforced locally before re-save). **Fix:** delete-on-load and enforce `expires_at` before returning.

### P2-7 — Microsoft-supplied verification URL is opened with no allowlist (code)
`auth.py:159-163` takes `verification_uri`/`verification_uri_complete` straight from the token-endpoint response body; `auth.py:195` does `webbrowser.open(...)` and `cli.py:20-21` prints it. Only reachable if `login.microsoftonline.com` (constant host) returns something hostile, but the browser will happily follow. **Fix:** require the URL to start with `https://login.microsoft.com/` or `https://login.microsoftonline.com/`.

### P3-1 — Audience/scope/PUID/tid are read from an **unsigned** JWT (defense-in-depth)
`auth.py:82-90` base64-decodes without signature verification; `auth.py:306-320` gates on `aud`, `api.py:224-226` derives `X-AnchorMailbox`'s PUID/tenant from it. Microsoft validates the token server-side, so this is a weak check rather than exploitable — but "we validate the token" overstates it, and `exp` is never checked locally. **Fix:** note it in docs; if a real gate is wanted, verify against Microsoft's JWKS.

### P3-2 — Request bodies are echoed into a header for small requests (design note)
`api.py:206-211`: everything up to 4096 ASCII bytes of percent-encoded JSON — which includes `emailId`, folder IDs and the **reply-draft body text** — is placed in `X-OWA-UrlPostData`. It is percent-encoded with `safe=""`, so **no CRLF/header injection is possible**. But request bodies in headers are more often logged by intermediate appliances than bodies are; with P1-1 present, the proxy is in the path.

### P3-3 — `search_emails` filtering is client-side and silently truncated (correctness)
`service.py:348` + `service.py:383-387`: `query` is a substring match over subject/sender-name/sender-address performed locally after paging through up to `max_scanned_items` (`service.py:356-357`, `config.py:51-56`). Positive: the caller's `query` never reaches the server, so there is no server-side injection path. Negative: results can be incomplete while looking complete; the skill file does tell the agent to state limits (`skills/.../SKILL.md`, "Do not claim complete coverage").

### P3-4 — No injection/execution surface found (positive result)
`grep -rnE "eval\(|exec\(|subprocess|os\.system|shell=True|pickle|yaml\.load|marshal|__import__|importlib"` over `src/` → **no matches** (only the `webbrowser.open` at `auth.py:195` and `os.getenv` reads in `config.py`). No dynamic dispatch from remote data: response JSON is walked by typed extractors (`parsers.py:11-24`) and re-emitted as fixed dict shapes (`models.py:40-132`). **No file writes at all** in the request path: `grep -rnE "open\(|write_text|write_bytes|makedirs|Attachment|download"` returned only `HasAttachments` boolean reads (`parsers.py:133,151`, `service.py:240`) — attachments are reported but never fetched, so there is no path-handling risk. Mailbody handling is deliberately de-HTMLified to plain text (`parsers.py:25-53`, `body_text` at `parsers.py:99-105`), so **inline links are rendered as inert text, not markdown links** — good prompt-injection posture; there is no "fetch this URL" tool anywhere.

### P3-5 — Supply chain is clean (positive result)
Runtime deps (`pyproject.toml:12-18`) with installed versions: `httpx==0.28.1`, `keyring==25.7.0`, `mcp==1.29.0`, `filelock==3.32.2`, `platformdirs==4.11.0`; transitive `anyio==4.14.2`, `pydantic==2.13.4`, `certifi==2026.7.22`, `cryptography==50.0.0`. All names are canonical PyPI packages — **no typosquat-shaped names**. All 55 external `source` entries in `uv.lock` are `registry = "https://pypi.org/simple"`; the only non-PyPI source is the project itself (`source = { editable = "." }`, `uv.lock:554`). No `git`, no URL, no vendored wheels. The only build backend is `hatchling` for this project itself; **no dependency ships an install-time build script of concern**. `plugin.json` is pure metadata (name/author/keywords, no lifecycle or install hooks). `mcp.json` is a stdio launch of `uv run --locked --no-dev --project ${PLUGIN_ROOT} outlook-web-mcp-server` with only `UV_CACHE_DIR`/`UV_PROJECT_ENVIRONMENT` env — it does not widen the tool surface and does not auto-install anything at the agent layer. `skills/use-outlook-web-mcp/` is two Markdown files with **no scripts**; their content is defensive (untrusted-data rules, "never claim a draft was sent", "re-resolve before state changes").

---

## 4. Claims: verified vs falsified

| # | Doc claim | Status | Evidence |
|---|---|---|---|
| 1 | Device-code flow with Outlook Web's first-party client ID `9199bf20-a13f-4107-85dc-02114787ef48`; no app registration, no admin consent, no secret; scope `…/.default openid profile offline_access` | **VERIFIED** (with caveat) | `config.py:8-10`; `auth.py:20-21,139-144,147-150,250-257,280-288`; no `client_secret` anywhere. Caveat: `.default` ⇒ read-write mailbox (P1-2), and the README itself warns CA/consent policy may block it |
| 2 | `POST …/owa/service.svc?action=…&app=…` with `Bearer` + `X-AnchorMailbox: PUID:<puid>@<tid>`, `X-OWA-UrlPostData` for small; no cookies, no CANARY, no browser profile | **VERIFIED** | `api.py:214`, `api.py:227-243`, `api.py:206-211` (`URL_POST_DATA_MAX = 4096`, `api.py:15`). Cookie/CANARY/browser grep over `src/*.py` → zero matches |
| 3 | Refresh→OS keyring only; access→memory only; nothing sensitive on disk; nothing sensitive in stdout/MCP responses | **VERIFIED for storage; FALSIFIED in part for "nothing sensitive in responses"** | Access token never persisted: `auth.py:114,236,357` only. Refresh chunked into Keychain: `credentials.py:148-174` (item observed by name only). Cache dir holds only 0-byte `.lock` files. Stdout is token-free (`cli.py:15-21,33-40`). **But** `tenant_id`/`display_name` (`auth.py:374-381`) and upstream error text (`api.py:65-70,179-192`) do reach MCP responses — P2-5, P1-3 |
| 4 | Tool surface excludes send/delete/move/archive; exactly the 7 named tools; `create_reply_draft` is SaveOnly | **VERIFIED as code, NOT a security boundary** | `server.py:52-159`; `test_server.py:12-21` asserts the exact 7-name set; action enumeration yields 7 values; `MessageDisposition` hard-coded twice (`service.py:473,524`); `"Send"` never appears as an action. Capability caveat: P1-2 |
| 5 | Guardrails: size cap, scan caps, calendar-range cap, no cross-host redirect following, no bearer forwarding to non-fixed hosts, error messages carry no tokens/bodies/mail content | **MOSTLY FALSIFIED** | ✅ redirects: `follow_redirects=False` (`api.py:115`, `auth.py:133`) + refuse all 3xx (`api.py:148-154`) — stronger than claimed. ✅ caps: `service.py:185-188,328-331,430-431,567-572`, `config.py:66-78`. ⚠️ size cap post-buffer (P2-1). ❌ "no bearer to non-fixed hosts": URL is pinned but transport follows `HTTPS_PROXY`/macOS system proxy (P1-1). ❌ error-message hygiene (P1-3) |

---

## 5. Hardening checklist (cheap-first, all local)

1. **[config, ~1 min — do this first]** Scrub proxy/CA env for the server process: `env -u HTTP_PROXY -u HTTPS_PROXY -u ALL_PROXY -u SSL_CERT_FILE -u SSL_CERT_DIR uv run --locked outlook-web-mcp-server`, and check System Settings → Network → Proxies is off for that account. (Today it resolves to `127.0.0.1:7897`.)
2. **[code, 2 lines]** Add `trust_env=False` to `api.py:113` and `auth.py:131`. This is the durable version of #1.
3. **[code, ~5 lines]** Return only `code`/`http_status` in tool errors; keep `MessageText` on stderr (P1-3).
4. **[config, 1 min]** Keep the server **stdio-only, never HTTP**: current `server.py:169` / `mcp.json` are correct — do not put it behind any network transport or shared tunnel.
5. **[config]** Lower blast radius of reads: `OUTLOOK_WEB_MAX_RESPONSE_BYTES=2000000`, and have the agent always pass small `maxBodyCharacters` (server allows up to 50,000 — `service.py:430-431`).
6. **[config, strongest available mitigation for P1-2]** Run the server as a **dedicated macOS user** (separate Keychain, separate login) used only for mailbox automation; never reuse the founder's primary Keychain for this item.
7. **[code, ~4 lines]** Replace `priority > 0` with an explicit backend-module allowlist (P2-2).
8. **[code, 1 line]** In `_send`, assert `response.url.host == "outlook.cloud.microsoft"` before parsing — cheap tripwire against future redirect/proxy regressions.
9. **[code]** Stream the response and abort past the cap (P2-1).
10. **[config, needs sudo — optional egress fence]** pf anchor allowing only `login.microsoftonline.com` and `outlook.cloud.microsoft` :443 for the dedicated uid (`pfctl` anchor + `anchor "out"`), or an application firewall in proxy-aware mode. Note: a *proxy*-based allowlist (Little Snitch/Clash) is exactly the P1-1 vector, so if you fence with a local proxy you must also pin the CA (#1/#2).
11. **[code]** Delete pending-login records on load and enforce `expires_at` (P2-6); require the verification URL to be on a Microsoft host (P2-7).
12. **[config]** Pin the code: there is no `.git` here, so nothing ties this tree to the commit `c78d4fd9…` that `skills/.../tool-contract.md:3` claims to be verified against. Check out that commit (or a tag) and diff before use.
13. **[config, agent-side]** Restrict the exposed tool list where the client supports per-tool allowlists — read-only (`search_emails`, `get_email`, `list_mail_folders`, `list_calendar_events`, `outlook_auth_status`) and leave `set_email_read_state`/`create_reply_draft` disabled unless actually needed. Do **not** co-register a URL-fetching or shell tool in the same agent that holds mailbox content (P3-4 shows the server itself is clean on this).
14. **[monitor]** Watch for anomalies that the server cannot see: University/AAD sign-in logs showing the first-party Outlook Web client ID from unexpected IPs or non-browser clients; unexpected Keychain access prompts for `ma-nakaya/outlook-web-mcp-server`; a new `<profile>:metadata` item you did not create.

---

## 6. What could not be verified from source alone

- **University Conditional Access / tenant policy.** Whether device-code flow is blocked or MFA-constrained for `connect.example.edu`, and whether the tenant's policy even honors a first-party Outlook Web client ID. The README itself concedes this (`README.md`, the `OUTLOOK_WEB_CLIENT_ID` paragraph).
- **Effective delegated permissions of `9199bf20-…` in this tenant.** Whether `.default` really lands `OWA.AccessAsUser.All` (read-write) vs a narrower set is a server-side grant; `auth.py:314-319` only proves the code *requires* one of the two.
- **What Microsoft logs server-side**, whether `service.svc` JSON-dialect use is rate-limited/flagged, and whether Microsoft may break or block it without notice (their own EWS-retirement posture is cited in `docs/internal-api.md`).
- **macOS Keychain ACL behavior** for these items (which binaries may read them without a prompt) and whether the login keychain is unlocked by an unattended launch.
- **Whether the on-disk tree equals GitHub `main`.** No `.git`, so provenance is asserted, not proven — the highest-value remaining risk, since this audit certifies *this tarball*.
- **Priorities of hypothetical third-party keyring backends** (nothing plaintext-capable is installed — P2-2).
- **Whether the live in-flight device-code login ends up scoped to the intended account**, and whether the tenant is the expected University one.

---

## 7. Method (reproducible commands)

Reads: `config.py`, `auth.py`, `credentials.py`, `api.py`, `server.py`, `service.py`, `errors.py`, `models.py`, `parsers.py`, `cli.py`, `__init__.py`, all 4 test files, `README.md`, `docs/internal-api.md`, `pyproject.toml`, `mcp.json`, `plugin.json`, `.github/workflows/ci.yml`, `skills/**`.

```bash
# A — every URL literal in shipped source
grep -rnoE "https?://[a-zA-Z0-9._/?=&%-]+" src/            # -> 4 hits, all Microsoft
# A — clean host enumeration, project files only (excludes .venv and uv.lock)
grep -rhoE "https?://[a-zA-Z0-9.-]+" README.md docs/ src/ tests/ plugin.json mcp.json \
  pyproject.toml .github/ skills/ | sed 's|https\?://||' | sort | uniq -c | sort -rn
# A — telemetry / phoning-home keyword sweep
grep -rniE "(telemetry|sentry|posthog|analytics|amplitude|segment|tracking|update.?check|version.?check)" \
  --include='*.py' --include='*.json' --include='*.toml' --include='*.md' --include='*.yml' . | grep -v '^./.venv'   # -> NONE FOUND
# A — module inventory (only httpx touches the network; no requests/aiohttp/socket/sentry)
grep -rhoE "^(import|from) [a-zA-Z0-9_.]+" src/ | sort -u
# B — secrets: stdout/logging/print surfaces, env reads, cache dir, keychain item NAMES only (never -w)
grep -rn "print(\|logging\.\|logger\." src/outlook_web_mcp/*.py
grep -rnE "getenv\(|environ" src/outlook_web_mcp/*.py
ls -la ~/Library/Caches/outlook-web-mcp-server/ ; find ~/Library/Caches/outlook-web-mcp-server -type f
security find-generic-password -s "ma-nakaya/outlook-web-mcp-server" | grep -E '"acct"|"svce"'   # no -w
# B — keyring backend reality check in the repo venv
uv run --no-sync python -c "import keyring; b=keyring.get_keyring(); print(type(b), b.priority, hasattr(b,'is_secure'))"
uv run --no-sync python -c "from keyring.backend import get_all_keyring; [print(type(k).__module__,k.priority) for k in get_all_keyring()]"
uv run --no-sync python -c "import importlib.metadata as m; print(m.version('keyring'))"
# C — client construction, redirect refusal, host pinning, proxy/CA handling
grep -rn "httpx\.\|AsyncClient\|follow_redirects\|trust_env\|verify=\|proxies\|base_url" src/
grep -n "trust_env\|allow_env_proxies\|_get_proxy_map" .venv/lib/python3.13/site-packages/httpx/_client.py
sed -n '30,45p' .venv/lib/python3.13/site-packages/httpx/_utils.py
grep -n "SSL_CERT_FILE\|SSL_CERT_DIR" .venv/lib/python3.13/site-packages/httpx/_config.py
uv run --no-sync python -c "import httpx; print(httpx._utils.get_environment_proxies())"   # -> 127.0.0.1:7897
grep -n "netrc" .venv/lib/python3.13/site-packages/httpx/_client.py                          # -> no matches (httpx 0.28 does not read .netrc)
# D — action / disposition enumeration and the absence of any send path
grep -rn "\.call(" src/outlook_web_mcp/service.py src/outlook_web_mcp/cli.py
grep -rhoE '"[A-Z][A-Za-z]+"' src/outlook_web_mcp/service.py | sort -u
grep -rn "MessageDisposition" src/ ; grep -rn "Send" src/outlook_web_mcp/*.py | grep -viE "DateTimeSent|Sender|sender"
# E — injection/execution and file writes
grep -rnE "eval\(|exec\(|subprocess|os\.system|shell=True|pickle|yaml\.load|marshal|__import__|importlib" src/   # -> none
grep -rnE "open\(|write_text|write_bytes|Attachment|GetAttachment|download|makedirs|mkdir" src/outlook_web_mcp/*.py
grep -rniE "cookie|canary|set-cookie|browser_profile|playwright|selenium|sqlite|localStorage" src/ --include='*.py'  # -> NONE
# G — supply chain
uv run --no-sync python -c "import importlib.metadata as m; [print(p,m.version(p)) for p in [...]]"
grep -c 'registry = "https://pypi.org/simple"' uv.lock ; grep -n "source = { git" uv.lock    # -> 55 pypi, no git sources
# H/I — tests, offline
cd <local checkout> && uv run --no-sync pytest -q
# -> 11 passed, 0 failed, no network required (MockTransport + StubApi/StubAuth only)
```

**Test-honesty verdict.** The suite does assert real security properties, not just parsing: fixed host/path/`action` + `Bearer` + `X-AnchorMailbox` + `Origin` + envelope shape (`test_api.py:70-80`), no `"Bearer"` in error text (`test_api.py:110`), refresh token absent from `ensure_status()` output (`test_auth.py:161-165`), rotation actually saved (`test_auth.py:158-159`), `UpdateItem` touching only `IsRead` (`test_service.py:163-167`), `CreateItem` with `MessageDisposition == "SaveOnly"` and `"Send" not in repr(payload)` (`test_service.py:190-194`), the exact 7-tool surface with correct annotations (`test_server.py:12-31`). `test_service.py:40-47` even documents a time-bomb avoidance for `DateTimeReceived`.

**But every test is fixture-driven, so the whole suite would pass identically against an empty or entirely absent mailbox** — and indeed it passes with zero network access here. Nothing in it observes a real message, so a regression that made the tools silently return nothing (e.g. `service.py:372` skipping items with an empty id, or the scan caps at `service.py:356`/`service.py:294` short-circuiting) would go undetected; the only live-data check is the `probe` CLI (`cli.py:83-107`), which is excluded from CI and asserts only `response_has_body`. `credentials.py` — the most security-critical module — has **no tests at all**, and CI is Windows-only (`ci.yml:8`).

---

## 2026-11-07 复核：哪些已关、哪些还开着

这份审计是对 **fork 之前** 的树（当时 7 个工具）做的。逐条回看当前代码：

| 项 | 状态 | 现在的证据 |
|---|---|---|
| P1-1 环境代理/CA 劫持 | **已关** | `api.py` 与 `auth.py` 两个 client 都 `trust_env=False`、`verify=certifi` 钉根；显式代理只来自 `OUTLOOK_WEB_PROXY`/系统设置，不是 `HTTPS_PROXY`。`tests/test_proxy_resolution.py`、`tests/test_tls_trust.py` |
| P1-2 `.default` 权限过大 | **不可在本地收窄**（审计结论未变） | 运维缓解：stdio-only、独立 macOS 用户/Keychain、盯 AAD 登录日志 |
| P1-3 上游错误文本回显 | **已关** | `_first_error` 只返回 `(code, code)`；原文只进 opt-in 的 0600 日志。`tests/test_error_hygiene.py` |
| P2-1 大小上限在整包读入后才判 | **已关** | 改成 `client.stream()` 逐块累加，超限立刻 `response_too_large`；测试断言在**第 3 块**就停（不是读完再判）。注意重建响应时剔除 `content-encoding`，否则重复解压（PITFALLS #22） |
| P2-2 keyring 只测 `priority > 0` | **仍开着** | 换成了文档化的理由 + 集成测试，但没有 `type(backend).__module__` 允许清单 |
| P2-3 凭据无测试 / CI 不跑 macOS | **已关** | `tests/test_credentials.py` 存在；CI 已是 ubuntu+macOS+Windows 矩阵（但**从未跑过**，仓库无 origin） |
| P2-4 `OUTLOOK_WEB_CLIENT_ID` 可重绑 | **已关** | 刷新前比对 bundle 的 `client_id`，不一致直接 `client_id_mismatch` 拒绝（`tests/test_auth.py`）；不给任何回退 |
| P2-5 `tenant_id`/`display_name` 进模型上下文 | **已关** | 只在工具出口过滤（`server.py`），CLI 仍给本人看；删掉过滤则 `tests/test_server.py` 变红 |
| P2-6 中断的 device code 留在 Keychain | **已关（读时清理，无主动 GC）** | `login-complete` 一读到过期 pending 就删除并报 `device_code_expired`；`complete_login` 自带过期检查。若再也不跑 `login-complete`，这条记录会一直留着——但没有本地代码会去兑现它 |
| P2-7 验证 URL 无允许清单 | **已关** | `safe_verification_url()` 只放行 https + Microsoft 主机；`tests/test_auth.py` |
| P3-1 未验签 JWT 取 aud/scope/PUID | **仍在**（防御纵深） | 文档已说明"校验"指的是调用方服务端校验 |
| P3-3 搜索是客户端过滤 | **已关（按设计）** | `meta` 里明写 `truncated/hitCapReached/coverage`，skill 要求据此报覆盖范围 |
| 加固清单 #8 断言响应 host | **已做** | `api.call()` 每次比对 `response.request.url.host` 与钉死的 OWA 主机，不符即 `host_mismatch` |
| 加固清单 #9 流式截断 | **未做** | 同 P2-1 |

结论：P1 全部有结论（两条已修、一条只能运维缓解）；P2 **7 条全部关闭**；
这些"仍开"项都是本地可修但**不影响当前威胁模型**的低危项，已在上面列明，不假装已修。
