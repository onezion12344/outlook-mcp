# outlook-web fork 设计（outlook-web-mcp）

日期：2026-11-06　状态：待 owner 复核
上游：`ma-nakaya/outlook-web-mcp-server` main = `aa41b13`（2026-11-01），MIT
本地已审副本：`<local checkout>`（含 `AUDIT-2026-11-06.md`）

## 0. 目标与非目标

目标：把上游这个"能读 University 邮箱的 OWA 私有协议 MCP"变成我们自己长期维护的版本，解决三件实测缺陷（搜索语义、静默截断、文件夹覆盖），加上发信 + 已读追踪，并把所有坑沉淀成可复用的文档与 skill。

非目标：不做 delete/move/archive 类破坏性操作；不做多账号；不做 GUI；不替换 Superhuman 的网页发信路径（发信改为走我们的 nonce 闸门，Superhuman 那条保留为兜底）。

## 1. 已核实事实（这些同时是 PITFALLS 的种子）

| # | 事实 | 证据 |
|---|---|---|
| F1 | "搜索"不是服务端搜索：分页拉条目后在本地对 `subject + sender_name + sender_address` 做 casefold 子串匹配。**正文永远搜不到，KQL 完全无效** | `src/outlook_web_mcp/service.py:383-387`；实测 `from:example.edu` → 0 条，而收件箱确有大量 example.edu 邮件 |
| F2 | 扫描上限 5000 条 / 100 文件夹 / 时间窗 ≤365 天 / `maxResults` 硬顶 50；**响应里没有任何 truncated 或 coverage 字段**，"没命中"与"被截断"不可区分 | `service.py:356-400`、`config.py` `max_scanned_items=5000` |
| F3 | 只认 `inbox / sent / drafts`；`archive`、`junkemail`、`deleteditems` 直接报错 | 实测 `Supported folders: inbox, sent, drafts` |
| F4 | `includeSubfolders` **其实已实现**（递归 `FindFolder` 并按 `FolderClass` 前缀 `IPF.Note` 过滤，`service.py:284-313`）。我先前实测"无差别"是探针无效：owner 的 Inbox 没有子文件夹，正负对照都返回同样结果。缺的是测试，不是功能 | `service.py:284-313`；实测不可判定 |
| F4b | `_resolve_folder` 的报错文案写死 "Supported folders: inbox, sent, drafts"，但 `DISTINGUISHED_FOLDERS` 实际有 6 个键（含 `calendar`）——文案与代码不一致，误导调用方 | `service.py:25-32`、`service.py:168` |
| F5 | 日历窗口 >31 天会 0 秒响亮报错（校验有效）；但每个合法窗口都返回 0 事件，**无法证伪是空邮箱还是读取路径不通** | 实测 |
| F6 | 性能：列最近 0.4–0.8s、单封正文 0.30s、关键词搜 30d 0.77s / 90d 1.85s / 180d 4.09s / 365d 6.3–6.9s（线性重扫，无缓存） | 实测 |
| F7 | 出网只有两个微软主机；`follow_redirects=False` 且拒绝所有 3xx；`OWA_ORIGIN`/`OWA_ENDPOINT` 是模块常量，**不能**用环境变量改写；无 eval/exec/subprocess/pickle；无附件落盘 | `AUDIT-2026-11-06.md`，`api.py:148-154` |
| F8 | 但两个 httpx client 都没设 `trust_env=False`，故跟随系统代理与 `SSL_CERT_FILE`。本机实况：mihomo **未配 mitm**、`tun.enable=false`，但 System keychain 里有一张 **mitmproxy CA（2026-05-08 → 2036）**。**2026-11-06 实测追加**：本机 `NO_PROXY` 含 `[::1]`，httpx 在 `trust_env=True` 时**构造 AsyncClient 就抛** `InvalidURL: Invalid port: ':1]'` —— 单变量复现：`trust_env=True`→崩、`trust_env=False`→正常、`trust_env=False`+显式 `proxy=`→正常。所以"保留 trust_env"在这台机器上是坏的，且崩不崩取决于调用方 shell 有没有该变量 | `api.py:113-117`、`auth.py:131-135`；本机 repro 脚本实测 |
| F9 | token 实际是 `Mail.ReadWrite` 级（`.default` + `offline_access`），"没有 send 工具"只是客户端自律，不是权限边界 | `config.py:10`、`auth.py:314-319` |
| F10 | `_first_error()` 原样回显 Exchange `MessageText` 并作为 tool result 文本进入 agent 上下文，与文档的错误卫生承诺矛盾 | `api.py:60-80` → `api.py:179-192` |
| F11 | 11 个测试全 fixture 驱动，**空邮箱/无邮箱也全绿**；`credentials.py` 零测试；CI 只有 Windows | 审计 + `uv run --no-sync pytest -q` → 11 passed |
| F12 | keyring 在本机是 `macOS.Keyring` priority=5；`is_secure()` 在 keyring 25.7.0 已不存在，`priority>0` 是唯一可用代理判据；access token 确认从不落盘 | `credentials.py:37-50`、`auth.py:114,236,357` |
| F13 | 上游自己的 tool-contract 文档引用 `c78d4fd`（8-04 initial release），已过时 | `skills/.../tool-contract.md:3` vs `aa41b13` |
| F14 | Superhuman 的 read status = 追踪像素；2019 年舆论反弹后**移除位置追踪、删除定位数据、把已读回执改为默认关闭** | [blog.superhuman.com/read-statuses](https://blog.superhuman.com/read-statuses/)、geekpark/36kr/engadget 2019-07 |
| F15 | Gmail 官方：MDN 已读回执**仅 Workspace 账号**可用、个人 @gmail.com 不支持，且收件人需批准；Outlook 侧 MDN 走协议而非图片 | support.google.com/mail/answer/9413651 |
| F16 | 像素双向不准：Gmail 对陌生发件人不自动加载远程图 → **漏报**；Apple Mail Privacy Protection 预取 → **误报**；Gmail 代理只隐藏 IP，仍暴露打开时刻 | filippo.io 2013、mailbird 2026 |

## 2. 仓库形态

- 新目录 `~/Projects/outlook-web-mcp`，`git init`，首个 commit 原样落上游 `aa41b13` 树并在 commit message + `NOTICE` 里写明 pin。
- remote：`origin` = GitHub 私有仓 `<owner>/outlook-web-mcp`（**推送前单独征求同意**）；`upstream` = ma-nakaya。
- `FORK.md`：rebase 政策 `git fetch upstream && git rebase upstream/main`；我们的 patch 保持小叠、每 patch 一个行为；不整文件重写上游。
- 保留上游 `LICENSE`（MIT）+ 加 `NOTICE`（fork 署名、上游 commit、修改清单）。
- **keyring 服务名沿用 `ma-nakaya/outlook-web-mcp-server`**，这样不重新登录；作为有意的凭据继承写进文档。

## 3. 安全修复（v0.1，先于任何新功能）

- **S1（2026-11-06 由 A 修正为 A′）**：两个 client 都 `verify=certifi.where()` **且** `trust_env=False`；
  代理不再靠环境变量继承，而是我们自己解析：`OUTLOOK_WEB_PROXY` → macOS `scutil --proxy` → 直连。
  理由：owner 有意保留自有代理出口（A 的初衷），但实测 `trust_env=True` 在本机 `NO_PROXY=[::1]`
  下构造客户端即崩（F8），且崩不崩取决于调用方 shell —— 不可接受。A′ 同时拿到三件事：
  信任根钉死、走不走代理变成显式可控、消除环境依赖型崩溃。票：`T02b`。
- **S2**：`_first_error()` 只返回 `errorCode` + Exchange 错误**名**，绝不回显 `MessageText`；完整文本仅写入 `0600` 的本地日志，且需 `OUTLOOK_WEB_VERBOSE_ERRORS=1`。
- **S3**：README/docs 明写 F9（token 是 ReadWrite），把"发信安全"从"缺工具"改为"服务端闸门"（§5）。
- **S4**：`credentials.py` 补测试；keyring 判定加注释说明 `is_secure()` 已不存在这一事实。
- **S5**：文档过期 commit 引用（F13）改为动态 pin。

## 4. 检索修复（v0.1）

- **R1 诚实信号**：所有搜索响应加 `{scanned_items, coverage:{from,to}, truncated: bool, hit_cap_reached: bool}`。`truncated` 在 scanned 触顶或命中触顶时为 true。这是本版最重要的一条：让"没搜到"变成可区分的两种情况。
- **R2 真分页**：`search_emails` 加 `offset`（沿用内部 `next_offset` 游标语义），`maxResults` 上限提到 200；`max_scanned_items` 默认提到 20000 并允许 env 覆盖，触顶必须报 `truncated=true`。
- **R3 文件夹覆盖**：`list_mail_folders` 根改为 `MsgFolderRoot`，支持 `archive / deleteditems / junkemail / outbox` 与自定义文件夹；`search_emails` 接受 `folderId` 直查任意文件夹。
- **R4 参数可证**：`includeSubfolders` 保留（F4：它已实现），补一个"父文件夹有子文件夹"的 fixture 测试证明递归真的发生；并在 `meta.scanned_folders` 里回报实际扫描了几个文件夹，让调用方能判断它是否起作用。同时修正 F4b 的过期报错文案，改为从 `DISTINGUISHED_FOLDERS` 动态生成。
- **R5 服务端搜索 spike**：OWA 的 `FindItem` 支持 AQS `IndexedItemView`/`QueryString`，可能能把关键词交给 Exchange。做成开关：`OUTLOOK_WEB_SERVER_SEARCH=1` 时试服务端（顺带解决正文搜索与 KQL），失败自动回落本地匹配并在响应里标 `engine: "local"|"server"`。**UNVERIFIED，需先 spike。**

## 5. 发信：默认只起草，**owner 授权时可按草稿发送**（2026-11-06 二次修订）

**默认流程**：agent 走 `create_draft` / `create_reply_draft` 把内容放进 Drafts，owner 在
Superhuman 或 OWA 里 review 后自己按发送；已读追踪交给 Superhuman 自带机制。

**授权发送**（owner 追加要求："不是永远不碰，我允许时还是可以"）：保留一个发送能力，但它是
**对一封已存在草稿**的动作，且必须绑定 owner 对**那一份确切内容**的一次性授权：
- `send_draft(draftId)` 先重读草稿，按 `draft_id ‖ subject ‖ 排序去重后的小写收件人 ‖ 排序去重后的小写密送收件人 ‖ 附件指纹(name:size) ‖ sha256(**HTML 正文**，即真正发出去的那份；无 HTML 时退回文本)`
  算 16 位 digest；
- 在 `~/.config/outlook-web-mcp-server/approvals/<digest>.nonce` 找不到 token 就**拒绝**，错误码
  `send_not_approved`，消息里带上要执行的确切命令；
- token 由 owner 在终端亲手签发：`outlook-web-mcp approve <draftId>`（CLI 会重算 digest 并打印
  收件人/主题，确认他批的是真实内容）；发送时**消费即删除**，且 nonce **30 分钟过期**（过期的会被清掉并拒绝，避免"今天批的、一个月后才发"）；
- 因此**草稿在批准后被改一个字，digest 就变，发送被拒**。

设计取舍：
- 不做通用的 `send_email`（凭空构造并发送）：内容必须先作为草稿存在，owner 能在真实邮件客户端
  里看到它，发送对象是一个可指认的邮件实体，而不是 agent 现编的一段文本。
- 不做像素采集端 / Cloudflare Worker / D1：追踪交回 Superhuman（§6）。
- `SendOnly` 全仓出现次数必须为 **0** —— 发送只经 `SendItem`，不给 `CreateItem` 任何"写完即发"
  的路径（这也是 v0.2 的硬门禁之一）。
- delete / move / archive 依旧不做。

（历史：初版方案是通用 `send_email` + digest nonce + 每日上限；随后 owner 提出"只起草、由我在
Superhuman 发"，遂删去整套发送/像素子系统；最终 owner 补一句"我允许时还是可以"，于是以
`send_draft` + CLI 签发 nonce 的最小形态复活发送能力，不带任何自动发送路径。）

v0.2 因此只含：
- `create_draft`：新建邮件草稿（上游只有 `create_reply_draft`）。HTML body、收件人/抄送/主题/正文，
  `MessageDisposition=SaveOnly` 落 Drafts，永不发送。
- `track_status`：对已登记草稿回报三档**能从本人邮箱读出**的事实 ——
  ① 是否已从 Drafts 进入 Sent Items（= owner 已发出）；② 是否收到回复（`In-Reply-To`/`References`）；
  ③ 是否收到退信（`MessageClass == REPORT.IPM.Note.NDR`）。
- 登记写本地 ledger（`draft_tracking` 表，并入 v0.3 索引），语义是"草稿追踪"，不是"发信审计"。

**明确不可见**：Superhuman 的像素事件落在 Superhuman 自己的服务器，不进本 fork，
所以 `track_status` 不提供 `opened`；UI 与文档不得暗示能读到。

**待实测的厂商假设**（owner 本人，本次无源码级证据；反证见 F14：Superhuman 2019 后默认关闭）：
"用 Superhuman 发送自带已读回执"。验收实验：agent 起草一封收件人为 owner 自有 Gmail 的草稿 →
owner 在 Superhuman 发送 → `get_email` 读那封已发送邮件的 HTML，检查是否含第三方像素 URL。
含 → 假设成立；不含 → 需在 Superhuman 设置里开启，且追踪只剩 ①②③ 三档。

（原 nonce 设计仅作历史记录保留：digest = 收件人排序 ‖ 主题  含像素的最终正文哈希，
`approve <digest>` 生成一次性 nonce，发送时消费。owner 决定不做。）

（原 §5 的 `send_email` / 每日上限 / `prepare_send` 预览 / outbox.jsonl 全部作废。）
- delete/move 仍然不做。

## 6. 追踪（v0.2 缩减版，从属 §5）

原四档里只有三档在本 fork 可见 —— 发信与像素都不归我们做：

| 档 | 信号 | 实现 | 状态 |
|---|---|---|---|
| T1 | 退信 NDR | 按 `MessageClass == REPORT.IPM.Note.NDR` 过滤，用 `internetMessageId` / 主题关联 | 保留 |
| T2 | 协议层已读回执（MDN） | 由 **Superhuman 发送时**决定是否请求；本 fork 不写 `Disposition-Notification-To` | **移出范围** |
| T3 | 追踪像素 | **本地半边已实现**（`pixel.py` + `pixel_store.py`）：`create_draft(track=true)` 注入并登记 token，`send_draft` 拒发外来像素，`track_status` 从本地事件文件报 `opened (image fetched)` + 低置信、无 IP。采集端 `collector/pixel-worker/`（Worker + D1）**已写好，未部署** | 本地已实现／**部署待 owner**（`pixel.example.com`） |
| T4 | 收到回复 | `In-Reply-To` / `References` 链到被回复项 | 保留 |
| T5 | 草稿是否已被 owner 发出 | 同一 `internetMessageId`/主题从 Drafts 消失且出现在 Sent Items | **新增**（§5 的必然产物） |

`track_status(draftId)` 返回 `{sent, replied, bounced}`，每个布尔都带证据（哪封邮件的 id 与时间），
不返回任何"已读"字样。

**保留的未验证项**（挪到 §11）：`MessageClass` 是否能从 `FindItem` 的响应形状里稳定取到
（上游 `_mail_page` 当前不请求该属性，见 v0.3 前置）；NDR 在 University 侧的实际 `MessageClass` 取值。

## 7. 本地索引（v0.3）

- SQLite + FTS5：`~/Library/Caches/outlook-web-mcp/index.db`（0700 目录 / 0600 文件）。
- 列：`email_id, store_id, folder, message_class, subject, sender_name, sender_address, received_at, is_read, has_attachments, body_preview, internet_message_id, in_reply_to`。**默认不存正文**（审计确认当前磁盘零敏感数据，整箱正文明文落盘是真实的姿态变化）；`--index-bodies` 为 opt-in，届时用 Keychain 密钥加密。
- 查询能力：FTS5 全文（含 preview）、字段查询（`from:` `since:` `messageclass:`）、按 `internet_message_id` 关联回执/回复/退信 —— 这三类都是服务端搜索做不到的（F1、F5）。
- 同步：每文件夹 `received_at` 水位 + R2 分页增量。删除/移动抓不到 → 先 spike `SyncFolderItems`；拿不到就每晚重建近 90 天窗口 + 全量重建命令，并把索引定性为**最终一致、只当线索**：`search_index` 命中后必须用 `get_email` 落实，id 失效返回明确错误而非假答案。
- 新工具：`index_sync`、`index_status`、`search_index`；`track_status`（§6 四档汇总）读索引。
- 文档口径：索引=广度与字段过滤，实时=权威。

## 8. 测试策略

- 非空 fixture：断言条数、`truncated` 触发、水位增量正确 —— 消灭 F11 那种"空箱也绿"。
- nonce 三态：无 nonce 拒、digest 不符拒、发送后 nonce 被消费。
- 追踪：像素注入位置正确、回复/转发被剥离、`track=false` 生效、no-store 头存在。
- 错误卫生：`MessageText` 不出现在 tool result（用含敏感串的 fixture 断言）。
- 网络：断言 `verify` 参数为 certifi、redirect 一律拒绝。
- CI：加 macOS runner。
- **live 红测**：一条对真实邮箱的最小断言（能列出 ≥1 封），只在 owner 手动触发时跑，避免 CI 持有凭据。

## 9. 沉淀（与代码同批交付）

- `PITFALLS.md`：§1 全部事实 + 每条四段（症状 / 根因 / 现在怎么防 / 复现命令）。
- 根目录 `AGENTS.md`：按 owner 的分层规范，指向 `PITFALLS.md`、`FORK.md`、`NOTICE`，带日期化 edit。
- skill `outlook-web-mcp`：引用本 fork（不复制逻辑），路由读/搜/草稿/发信审批/追踪，发信路径硬停在 nonce。
- 注册进 `~/.qoder-cn/settings.json` 的 `mcpServers.outlook-web`：**仅在 v0.1 的 S1+S2 落地且测试通过后**。

## 10. 版本切分与验收

- **v0.1**（安全 + 检索诚实）：S1–S5、R1–R4。验收 = 非空 fixture 测试全绿；一次 365 天搜索能明确告诉我它是否被截断；`archive` 文件夹可搜；`grep -n "trust_env\|verify" api.py auth.py` 人眼可核。
- **v0.2**（草稿 + 追踪，owner 决定后重写）：§5、§6。验收 = ① `create_draft` 能把一封带收件人/主题/HTML 正文的邮件存进 Drafts，且**没有任何路径能把它发出去**（grep 全仓 `SendOnly` 必须为零命中）；② `track_status` 对同一草稿能报 `sent=false`，owner 在 Superhuman 发出后报 `sent=true`；③ **厂商假设实验**：草稿发给 owner 自有 Gmail → Superhuman 发送 → 读已发送 HTML 判断是否含像素 URL，结论写进 PITFALLS。除 owner 自有地址外不发任何收件人。
- **v0.3**（索引）：§7。验收 = 一年正文关键词查询在本地 <200ms 返回且带覆盖区间；索引与实时结果对同一查询的差异能量化报告。

## 11. 未验证 / 风险清单

1. `SyncFolderItems` 在 OWA JSON dialect 是否可用（决定索引能否感知删除/移动）。
2. `Disposition-Notification-To` / `InternetMessageHeaders` 的属性名与 University 内部 MDN 实际行为。
3. 服务端 AQS 搜索是否可用（R5）。
4. Outlook 日历返回 0 事件到底是空还是读不通（F5）。
5. University Conditional Access 对 device-code 流程长期放行与否（今天放行，已实测）。
6. 像素采集域名与 Worker 部署需要的 owner 决定（用哪个域、是否复用 example.com）。
7. 隐私/合规：对收件人做打开追踪在部分机构 AUP 下可能不被接受；官方申请类信件的风险由 owner 承担，本设计只做知情与可关。
