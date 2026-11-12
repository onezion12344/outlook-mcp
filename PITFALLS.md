# PITFALLS — outlook-web fork

每条四段：**症状 → 根因 → 现在怎么防 → 复现命令**。
数字与行号来自 2026-11-06 在本机对 owner 的 University 邮箱（`connect.example.edu`，租户 `00000000-…`）的实测，
不是从 README 抄的。行号以 `5b24404` 为准，改动代码后要重新核对。

---

### 1. 搜索不是服务端搜索，正文永远搜不到
- **症状**：`search_emails(query="from:example.edu")` 返回 0 条，而收件箱里明明有大量 example.edu 邮件；换成 `query="university"` 才有结果。
- **根因**：`service.py:420-423` 把每页条目的 `subject + sender_name + sender_address` 拼成一个字符串做 **casefold 子串匹配**。KQL / AQS 字段语法（`from:` `subject:` `messageclass:` `kind:`）完全不参与请求，**正文与 Preview 也不在匹配范围内**。
- **现在怎么防**：`meta.engine` 恒为 `"local"`，工具 docstring 与 `server.py` instructions 明写"body text is never matched"。需要正文检索或字段过滤 → 等 v0.3 本地索引（FTS5 + `message_class`）。
- **复现**：`uv run --no-sync python scripts/live_smoke.py`（看 `query-university-365d` 那行的 scanned 与 returned）。

### 2. 满 200 条 ≠ 只有 200 条
- **症状**：上游返回正好 50 条时，调用方无法区分"就这些"和"被截断了"。实测 `query="a"` 搜一年返回 50 条、零提示。
- **根因**：上游响应只有 `{"result": [...]}`，没有任何覆盖或截断字段。
- **现在怎么防**：v0.1 起每次搜索带 `meta.scannedItems / matchCount / returnedCount / truncated / truncatedReason / hitCapReached / coverage / nextPage`。**只有 `truncated=false` 且 `hitCapReached=false` 的结果集才可以说"找完了"**。
- **曾经的假信号（2026-11-06 修）**：不带 `query` 且 `found == returned == maxResults` 时，`hitCapReached` 曾报 false，尽管服务端 `IncludesLastItemInRange=false`。`index_sync` 因此只灌 200 封就"完成"。现在"停在结果上限"会单独置 `more_available`，并入 `hitCapReached`。
- **复现**：live smoke 的 `query-university-365d` 行；或 `tests/test_search_meta.py::test_exactly_a_full_page_is_not_reported_as_complete`。

### 3. 扫描上限会静默停止（`daysBack` 上限已放开到 3650 天）
- **症状**：窗口越老、邮箱越大，越容易"搜不到"，且没有任何信号。
- **根因**：扫描循环受 `max_scanned_items` 与 `max_scanned_folders` 约束（`config.py:53-54`、`service.py:370` 每页 200）。
- **现在怎么防**：默认上限从 5000 提到 **20000**（`OUTLOOK_WEB_MAX_SCANNED_ITEMS` 可覆盖，合法区间 1–50000）；触顶时 `truncated=true` + `truncatedReason="scan_cap"`，并给出 `meta.continuation = {"folderId","offset"}` 供续扫。
- **复现**：`uv run --no-sync pytest tests/test_search_pagination.py -q`（`test_continuation_is_returned_when_scan_cap_hits`）。

### 4. 文件夹覆盖：上游只认三个名字
- **症状**：`folder="archive"` / `"junkemail"` / `"deleteditems"` 直接 `ValueError: Supported folders: inbox, sent, drafts`，归档里的历史完全查不到。
- **根因**：`DISTINGUISHED_FOLDERS`（`service.py:25`）只映射 6 个键，而报错文案还写死成三个（F4b，文案与代码不一致）。
- **现在怎么防**：映射补齐 `archive / deleteditems / deleted_mail / junk / junkemail / outbox / root`，`root` → `msgfolderroot` 可列整棵树；报错文案改为从映射表动态生成。实测 owner 邮箱 `folders: 32`，`archive30d returned 25`。
- **复现**：`uv run --no-sync pytest tests/test_search_folders.py -q`。

### 5. `includeSubfolders` 曾被误判为"假参数"
- **症状**：开与关结果一模一样（6.66s vs 6.89s，条数都是 50），当时结论写成"参数无效"。
- **根因**：**探针没有正对照** —— owner 的 Inbox 没有子文件夹，所以两种取值必然同结果。实现其实存在于 `service.py:296`（递归 `FindFolder`，按 `FolderClass` 前缀 `IPF.Note` 过滤）。
- **现在怎么防**：加了带子文件夹的 fixture 测试（`tests/test_search_subfolders.py`），并让 `meta.scannedFolders` 成为可观测证据。**教训通用：看不出差别 ≠ 没实现，先造一个能看出差别的样本。**
- **复现**：`uv run --no-sync pytest tests/test_search_subfolders.py -q`。

### 6. 日历：31 天硬上限是真的；"永远 0 条"是我探针的错
- **症状**：窗口 >31 天 0 秒报错（`Calendar range cannot exceed 31 days`，校验有效）。早先我又报告"每个合法窗口都 0 条"，据此怀疑日历读取坏了。
- **根因（两次实测后）**：① 上限是真的；② **"0 条"是假象** —— 我拿 `2026-10-06T00:00 → 2026-11-06T23:59` 当 31 天窗口，实际 32 天，被拒；换成本月（11-01 → 11-30）也确实 0，因为他 11 月 Outlook 里没有日程。
- **现在的结论（2026-11-06 实测）**：`list_mail_folders(folder="calendar")` 报 `totalItemCount=43`；合法 31 天窗口 `10-06T00:00 → 11-06T00:00` 返回 **1 条事件**。**日历读取是通的**，只是他的日程主要在 Google Calendar + SIS，Outlook 日历很稀。
- **教训**：用"几个空窗口"判定功能坏掉，等于探针没有正对照 —— 先查文件夹里到底有没有条目。
- **复现**：`list_mail_folders(folder="calendar", recursive=False)` 看 `totalItemCount`，再用一个严格 ≤31 天的窗口调 `list_calendar_events`。

### 7. 性能：线性重扫，没有缓存
- **症状**：搜一年比搜一个月慢一个数量级。
- **实测**：列最近邮件 `0.4–0.8s`（窗口无关，按时间倒序取 top N）；读单封正文 `0.30s`；关键词搜 `30d 0.77s / 90d 1.85s / 180d 4.09s / 365d 6.3–6.9s`；发件箱 365d `1.37s`。
- **根因**：每次调用都重新分页拉元数据再本地匹配，无索引无缓存。
- **现在怎么防**：v0.3 本地 SQLite/FTS5 索引（元数据 + preview，正文默认不落盘）。在此之前，查历史请分段扫并读 `meta.truncated`。
- **量窗口成本要用"匹配不到的词"**：无查询词时搜索在 `found >= maxResults` 就停（第一页 200 条），所以拿它比 1 年 vs 10 年会得到"一样快"的假结论。用 `query="zzzznomatch"` 才会走到底：实测 365d 与 3650d **都扫 4846 封**（这就是他 Inbox 的全部，`truncated=false`）。
- **波动不是回归**（2026-11-06 实测）：同一查询连跑两次 **5.8s / 7.3s**；曾观测到一次 **20.2s**，发生在 `index_sync` 与三次查询连发之后。逐属性二分显示单页（200 条）恒为 **0.23–0.25s**，与是否请求 `InternetMessageId` / `References` 无关。所以：**不要把单次慢归因到代码或新属性**，先复测两次再判断。

### 8. `trust_env=True` 在本机必崩（曾让 MCP 一启动就死）
- **症状**：`httpx.InvalidURL: Invalid port: ':1]'`，抛在**构造 AsyncClient 时**；崩不崩取决于调用方 shell 有没有 `NO_PROXY`。
- **根因**：本机 `NO_PROXY='localhost,.localhost,127.0.0.1,127.0.0.0/8,[::1],::1'`，httpx 解析 `[::1]` 时炸。单变量实测：`trust_env=True`→崩、`trust_env=False`→正常、`trust_env=False`+显式 `proxy=`→正常。
- **现在怎么防**：`trust_env=False`，代理改由 `proxy.py` 显式解析：`OUTLOOK_WEB_PROXY` → `scutil --proxy` → 直连；解析失败一律降级为 `None`，绝不抛。
- **复现**：`NO_PROXY='localhost,[::1],::1' uv run --no-sync python scripts/live_smoke.py`

### 9. TLS 信任根：本机有一张 2036 年才过期的 mitmproxy CA
- **症状**：若客户端跟随环境，`SSL_CERT_FILE` 能把信任根换掉。
- **根因**：System keychain 里存在 `CN=mitmproxy`（2026-05-08 → 2036）；mihomo 未配 mitm、TUN 关闭，所以当前没有实际解密发生。代理是 owner 自有的 Oracle SG 出口，不构成第三方泄露。
- **现在怎么防**：`tls.py` 用 `ssl.create_default_context(cafile=certifi.where())` 显式构造（顺带避开 httpx 0.28 对 `verify=<str>` 的废弃警告）。
- **复现**：`uv run --no-sync pytest tests/test_tls_trust.py -q`（含"设了 evil SSL_CERT_FILE 也不换根"那条）。

### 10. token 实际是 `Mail.ReadWrite` 级
- **症状**：看起来"没有 send 工具所以很安全"。
- **根因**：`config.py:12` 请求 `https://outlook.office.com/.default openid profile offline_access`，`auth.py:319` 要求 `OWA.AccessAsUser.All` 或 `Mail.ReadWrite`。缺 send 是**应用层自律，不是权限边界**。
- **现在怎么防**：README / AGENTS.md 都明写这一点；发送只保留 `send_draft`（只发**已存在**的草稿）+ 终端一次性 nonce（`approval.py`，digest 覆盖收件人/主题/正文），`SendOnly` 全仓 0 命中，没有"现编一封直接发"的路径。
- **复现**：`grep -rn "Mail.ReadWrite" src/`

### 11. 错误文本曾是邮箱内容的泄露通道
- **症状**：Exchange 的 `MessageText` 会引用邮件正文，被原样塞进 `ApiError` 并作为 MCP tool result 进入 agent 上下文。上游测试甚至把这个泄露当成期望断言。
- **根因**：`_first_error` 直接返回 `MessageText[:500]`。
- **现在怎么防**：`api.py:64` 只返回错误码；原文由 `api.py:89 _raw_error_detail` 提取，仅在 `OUTLOOK_WEB_VERBOSE_ERRORS=1` 时写入 `api.py:251` 的本地日志（`0600`）。`tests/test_api.py` 的旧断言已反转。
- **复现**：`uv run --no-sync pytest tests/test_error_hygiene.py -q`
- **CLI 也要干净失败**（2026-11-06 实测补）：`approve` 传坏 id → `ews_error: Outlook Web GetItem failed [ErrorInvalidIdMalformed]`、exit 1；`pixels-sync` 未配置 → `pixel_not_configured: …`、exit 1。两条都是单行、无 traceback（此前 ValueError 会吐栈）。

### 12. 上游测试全为 fixture，空邮箱也全绿
- **症状**：11 个测试全过，但真实链路上 MCP 一启动就崩（见 8）。
- **根因**：测试用 `StubApi` 喂假响应，对真实网络/环境零覆盖；`credentials.py` 无测试；CI 只有 Windows。
- **现在怎么防**：新增 `scripts/live_smoke.py`（只读、只打印计数、不打印正文）作为验收门；票面测试绿 ≠ 完成，**必须跑 live**。
- **复现**：`uv run --no-sync pytest -q` 绿之后仍要 `uv run --no-sync python scripts/live_smoke.py`

### 13. keyring 安全判据只剩 `priority > 0`
- **症状**：`credentials.py` 想验证"后端是否安全"。
- **根因**：`is_secure()` 在 keyring 25.7.0 已不存在（AttributeError）；本机 `get_keyring()` 是 `macOS.Keyring`，`priority=5`。该判据能拒掉 `fail.Keyring`(0)，但也会误拒 `ChainerBackend`(-1)。
- **现在怎么防**：保持 `priority <= 0` 即拒绝明文存储；本机未安装明文后端（`keyrings.alt` / `cryptfile` 探测均不存在）。访问令牌确认从不落盘。
- **复现**：`uv run --no-sync python -c "import keyring; b=keyring.get_keyring(); print(type(b).__name__, b.priority)"`

### 14. 上游文档引用过期 commit
- **症状**：`skills/.../tool-contract.md` 说"verified against `c78d4fd9`"，那是 2026-08-04 的 initial release，而 main 已到 `aa41b13`（2026-11-01）。
- **现在怎么防**：本 fork 以 `NOTICE` 里的 pin 为准；tool-contract 在 T08 改写。
- **复现**：`gh api repos/ma-nakaya/outlook-web-mcp-server/commits/main --jq .sha`

### 15. 已读追踪：能做的三档 + 像素那一档的真实上限
- **症状/背景**：owner 想知道"对方有没有看我的信"。
- **根因（协议层）**：邮件没有可靠的已读信号。可拿到的只有：① 退信 NDR（硬）；② 同租户 MDN（University→University 较可靠，落回收件箱为 `REPORT.IPM.Note.IPNRN`/`IPNNRN`）；③ 追踪像素（Gmail 对陌生发件人不自动加载远程图 → **漏报**；Apple Mail Privacy Protection 预取 → **误报**）；④ 是否收到回复（硬）。
- **厂商事实**：Gmail 官方 —— MDN **仅 Workspace 账号**，个人 @gmail.com 不支持，且收件人需批准。Superhuman 2019 年像素风波后**移除位置追踪、删除定位数据、把已读回执改为默认关闭**。
- **现在怎么防**：`track_status` 报 `sent / replied / bounced` 三档（扫 Sent+Inbox+Archive，带证据行），**不把 `opened` 说成"已读"**。
- **退信富化只认相关主题**：`MessageClass` 不在 `FindItem` 里，只能对候选再 `GetItem`。早先对任意"看起来像退信"的前 5 条都去取详情，无关退信会占满名额、把真正相关的那条挤掉。现在只富化**主题提及本草稿**的退信。
- **窗口本身就是一种漏**：默认只扫 30 天，60 天前发出的信，其回复不在窗口里，而 `conclusive` 仍为 true——读起来像"没回"。现在若记录了发送时间，扫描窗口自动延伸"发出日 + 7 天"，并在 `windowDays` 里如实回报。
- **发出去的自己记得**：`send_draft` 在 `SendItem` 之后会去 Sent Items 找那封（同主题 + 15 分钟内），把它记进 `sent-links.json`（draftId → sentId + internetMessageId，保留 180 天）。`track_status` 有这条链接时**不再靠主题猜**，`linked=true`、`ambiguous=false`。查不到就退回主题匹配 + 歧义保护；查失败绝不影响已发出的那封。
- **"默认追踪"要真的默认**：纯文本草稿装不了 `<img>`，早先只回一个 warning——等于默认没有追踪。现在需要追踪时把纯文本安全转成 HTML（转义 + `<br>`）再注入；没配 `OUTLOOK_WEB_PIXEL_BASE` 时才保持纯文本并警告。
- **同主题歧义**：`track_status` 原先把"已发"按主题匹配，两封同主题就会把回复归错。现在多封匹配时返回 `ambiguous: true` + `sentCandidates`，并把 `conclusive` 置 false、`replied` 置 false —— **宁可说"分不清"，也不给错答案**。
- **像素那一档（2026-11-06 落地本地半边）**：`create_draft(track=true)` 会把 `<img data-owt=…>` 注入 HTML 草稿并登记 token；`send_draft` 拒绝发送"不是本服务创建"的像素草稿（`pixel_unknown`），发送成功才把 token 标为已发；`track_status` 只从本地 `pixels/events.json` 读事件，报 `opened (image fetched)` + `confidence: low`，且永不带 IP。
  采集端已写成可部署产物（`collector/pixel-worker/`，Worker + D1，IP 只存 4 字节截断哈希，`cache-control: no-store`），**未部署** —— 域名 `pixel.example.com` 与部署动作等 owner。拉取端已实现：`outlook-web-mcp pixels-sync`（或 `pixels_sync` 工具）按 `since` 增量拉 `/events`，写入 `pixels/events.json`（0600），同一套 TLS/代理姿态；缺 base 或 token 时明确报错，不静默跳过。
- **复现**：`uv run --no-sync pytest tests/test_pixel.py tests/test_send_approval.py tests/test_track_status.py -q`

### 16. 索引抓不到"未扫窗口"里的删除与移动（窗口内已能自愈）
- **症状**：`search_index` 会命中一封早已删掉的邮件。
- **根因**：删除/移动需要 `SyncFolderItems`，但上游和两个参照实现（nhype、jlentink）里 **0 命中**，本 fork 也没有采用，所以索引只能按收件时间倒序拉窗口。
- **现在怎么防**：① 索引被定性为 eventually consistent 的**线索** —— 命中后必须 `get_email` 落实；索引里有、服务器没有的 id 会抛 `item_not_found`（`tests/test_index_authority.py` 钉住，绝不用缓存内容兜底）。② **窗口替换**：只要某文件夹这次把窗口完整走完（`truncated=false`），`index_sync` 就 prune 掉该窗口内这次没见到的行 —— 实测 resync 一年 inbox `pruned=0`、总行数稳定 4871，幂等。③ **截断时绝不 prune**（否则会删掉"只是没取到"的行），有单测钉住。跨文件夹移动若只同步了源文件夹，会表现为"源已删、目标缺失"，直到目标文件夹也被同步。
- **回填性能（2026-11-06 修）**：早先 `index_sync` 用公共 `offset` 翻页，而那是"结果切片"，每页都会从服务端 Offset 0 重扫（O(n²)）。现在改用 `meta.nextPage`（服务端游标）续页，一年 inbox 回填 **4846 封 / 13–15s**。公共 `offset` 参数语义未变。
- **复现**：`uv run --no-sync pytest tests/test_index_sync.py tests/test_index_authority.py -q`

### 17. 这个 dialect 里字段叫 `ItemClass`，不叫 `MessageClass`
- **症状**：按 `MessageClass` 取退信类型，搜索行永远是空；往 `FindItem` 里请求它会直接 400。`GetItem` 里看着像"服务器没给这个值"。
- **根因**：EWS 的属性名是 **`ItemClass`**。`MessageClass` 在这个 JSON dialect 里根本不存在——`GetItem` 对**不存在的属性静默返回 None**（所以看起来像空值而不是错误），`FindItem` 则直接拒绝未知属性。
- **实证（2026-11-06 真邮箱）**：真退信上 `ItemClass = REPORT.IPM.Note.NDR`；`FindItem` 请求 `ItemClass` 正常返回，搜索行直接带回 `IPM.Note` / `REPORT.IPM.Note.NDR` / `REPORT.IPM.Schedule.Meeting.Request.NDR`。
- **现在怎么防**：搜索形状请求 `ItemClass`；解析用 `ItemClass or MessageClass`；`_is_bounce` 在**有 class 时以 class 为准**（`REPORT.IPM.Note.NDR` 前缀），无 class 才退回主题正则——否则一封人类讨论 "undeliverable" 的邮件会被误判成退信。会议邀请的 NDR（`REPORT.IPM.Schedule.*`）不算邮件退信。
- **教训**：字段名猜错时，服务端给的是"空值"而不是"报错"；只有正对照（真退信 / 真回复）才能区分"没有值"和"名字错"。
- **复现**：`uv run --no-sync pytest tests/test_track_status.py -q`；真机 `search_emails(query="undeliverable", daysBack=365)` 看 `messageClass`。

### 18. "批准这封"必须绑定**全部**可发送面，而不只是正文
- **症状**：第一版 digest 只覆盖 `draft_id + subject + To/Cc + 正文`。于是批准后：换掉正文里的链接 `href`（可见文本不变）、偷偷加一个 **Bcc**、或挂上一个附件 —— 三者都不会改变 digest，旧批准照样能发出。
- **根因**：把"内容"理解成"可见文本"，而实际投递物是 HTML + 完整收件人集合 + 附件。
- **现在怎么防**：digest = `draft_id ‖ subject ‖ To/Cc ‖ **Bcc** ‖ 附件指纹(name:size) ‖ sha256(**HTML 正文**)`。`approve` 的终端输出同时打印 bcc 与附件清单，让"我批的是什么"看得见。三条测试分别钉住：换链接、加 Bcc、加附件 → digest 必变。
- **教训**：凡是"人批准后机器执行"的闸门，先列出**执行侧的全部输入**，再决定 digest 覆盖哪些；漏一个面就等于没有闸门。
- **还要加上时间维度**：批准过、没发送的 nonce 是一把上膛的枪。`consume` 校验 30 分钟 TTL，过期即删除并拒绝（`purge_expired` 在每次 `approve` 时顺手清），corrupt/无时间戳的 nonce 一律当无效并删掉——**失败方向永远是不发**。
- **复现**：`uv run --no-sync pytest tests/test_send_approval.py -q`

### 19. 时间戳带租户偏移，字符串比较会算错窗口
- **症状**：索引 `since` 过滤与 prune 窗口都拿 `received_at` 做**字符串**比较。服务端给的是 `2026-11-06T09:00:00+08:00` 这类带偏移的值，`2026-11-05T23:30:00-05:00`（= 11-06T04:30Z）会被判成"11-05、在窗口外"而删掉/漏掉。
- **根因**：把"看起来像时间"的字符串当成可比较的序；跨偏移时字典序 ≠ 时间序。
- **现在怎么防**：`index.normalize_timestamp()` 在**写入时**统一成 UTC `…Z`；`open_index()` 额外做一次性迁移，把历史行就地改写（实测 4871 行、迁移后非 UTC 行数 0）；`search_emails` 的排序改成按解析后的 UTC 时刻排，解析失败的排最后。
- **教训**：任何"用字符串比时间"的地方都是潜在 bug；只在写入时归一化一次，比在每个查询点都记得转换更可靠。
- **复现**：`uv run --no-sync pytest tests/test_index.py tests/test_time_ordering.py -q`

### 20. "索引能搜正文"只到每封前 255 字
- **症状**：`search_index` 能搜到正文里的词，但只在开头附近命中；深一点的段落搜不到。
- **根因**：服务端 `Preview` 属性是**截断**的 —— 实测开 `includeBodyPreview=true` 后 5/5 条都有值，但**长度恒为 255 字符**。索引存的是这个 preview，不是全文。
- **实证**：索引 7905 行 / 7904 行有 preview；从某条 preview 取一个 11 字母词去 `search_index`，源 id 命中。所以"能搜正文"是真的，但**只覆盖前 255 字**。
- **现在怎么防**：文档与回答里不说"全文检索"，说"正文前 255 字 + 主题/发件人"。真要全文只能 `get_email` 逐封读（有 5 万字符上限），或做 opt-in 的正文落盘（需加密 + 新依赖，尚未批准）。
- **复现**：`search_emails(daysBack=7, maxResults=5, includeBodyPreview=True)` 看 `bodyPreview` 长度恒 255。

### 21. `index_verify` 的 gap 是**新鲜度**，不是 bug
- **症状**：同一分钟内跑两次 `index_verify`，一次 `complete: true`，一次 `inbox gap=2 / complete: false`；第一反应是"注册的那条命令连到另一个库了"。
- **根因**：`gap = 服务器 TotalItemCount − 索引行数`。两封新邮件到了，索引就落后 2 —— 这是**正确的信号**，说明索引不是实时的。两次调用之间没有任何代码差异（库路径全局唯一，`server.py` 与 CLI 用同一个 `settings.index_path`）。
- **实证（2026-11-06）**：`index_verify` 报 inbox 4846/4848；`index_sync(folders=["inbox"], days_back=3650)` 回 `indexed=4848 pruned=0`；再 verify 四个文件夹 gap 全 0。
- **现在怎么防**：看到 `gap>0` 先当成"有新邮件"，跑一次 `index_sync` 再看；只有 sync 之后 gap 仍在，才是真问题。**不要**用"我刚刚还是 complete"去反推命令装错了。
- **教训**：把"库是快照"写进契约，就不会把时钟的走动误判成连接错误。凡是"两边计数对不上"的信号，先问"两边读的是同一个时刻吗"。
- **复现**：`index_verify` → 等一封新邮件 → 再 `index_verify`（gap 变大）→ `index_sync(folders=["inbox"])` → gap 归零。

### 22. 流式重打包响应时，别把 `content-encoding` 一起搬过去
- **症状**：只改 `_send` 用 `client.stream()` 之后，单测全绿，但真机 smoke 全部 `DecodingError: Error -3 while decompressing data: incorrect header check`——文件夹、搜索、日历一次全废。
- **根因**：`response.aiter_bytes()` 交给你的**已经是解压后的字节**；我重建 `httpx.Response(..., content=decoded)` 时把原始 headers 原样带上，于是消费者看到 `content-encoding: gzip` 又解压一遍。MockTransport 的响应没有 gzip，所以测试看不出来。
- **现在怎么防**：重建 Response 时剔除 `content-encoding` / `content-length` / `transfer-encoding`；并补一条**带 gzip 的** MockTransport 测试（`test_gzipped_response_is_decoded_exactly_once`），让单测也能碰到这条路径。
- **教训**：凡是"拆开再拼回去"的响应，帧头和载荷的契约必须一起维护；只有真机 smoke 会暴露"测试里没有压缩"这类盲区。

### 23. 这个 dialect 的收件人是**平铺条目**，不是 EWS 的 `{Mailbox: {...}}` 包裹
- **症状**：`create_draft` 报 `saved_to_drafts`、`ResponseCode: NoError`，但草稿 `DisplayTo` 为空——收件人被**静默丢弃**。改用 `SendAndSaveCopy` 则直接 `ErrorInvalidRecipients`。`UpdateItem SetItemField ToRecipients` 六种形状全部无效（要么 BadRequest，要么同样被丢）。
- **根因**：OWA JSON dialect 的收件人条目是**平的** `{Name, EmailAddress, RoutingType, MailboxType, ...}`；EWS 风格的 `{"Mailbox": {...}}` 包裹在 SaveOnly 下被服务端"接受然后丢弃"（比报错更危险），在 SendAndSaveCopy 下才报 `ErrorInvalidRecipients`。`ReplyToItem` 路径之所以一直正常，是因为收件人由服务端从被引用邮件推导，从不走这个字段。
- **实证（2026-11-07 真邮箱）**：wrapper 形状 × 5（含 typed `Recipient:#Exchange`、`Email` vs `EmailAddress`、header/body 两种承载）全部丢失；从真邮件 `GetItem` 抄回的**平铺条目**注入后，SaveOnly 与 SendAndSaveCopy 都保留收件人；`SendAndSaveCopy` + 平铺条目真实发送并落进收件箱。
- **现在怎么防**：`_mailbox_entry()` 统一产出平铺条目；`create_draft` 保存后**回读 `DisplayTo`**，为空即抛 `recipients_dropped`——宁可失败也不假装保存成功。
- **教训**：服务端对"合法但不对"的结构给 `NoError` 而不是报错时，唯一防线是**回读验证**；测试里 stub 不带 `DisplayTo` 的返回会让 fail-loud 闸门在 CI 里正确变红。
- **复现**：`uv run --no-sync pytest tests/test_create_draft.py -q`；真机对照 `create_draft` 后 `GetItem` 看 `DisplayTo`。

## 把坑变成命令

`outlook-web-mcp doctor` 把上面这些本地可检的项做成一条命令：keyring 后端优先级、凭据是否可读、索引行数与新鲜度（>7 天 warn）、待用/过期批准计数、像素是否配置、TLS 根与出口代理。**不联网、不打印令牌**；任一项 fail 时 exit 1。

## 上游 vs 本 fork：一句话对照

| 能力 | 上游 `aa41b13` | 本 fork v0.1 |
|---|---|---|
| 出网主机 | 仅微软两主机、拒绝所有 3xx | 同 |
| TLS 信任根 | 跟随环境（可被 `SSL_CERT_FILE` 换） | 钉死 certifi，`trust_env=False` |
| 环境代理 | 跟随（本机必崩） | 显式解析，失败降级直连 |
| 搜索覆盖信号 | 无 | `meta.*` 全套 |
| 分页/续扫 | 无 | `offset` + `continuation` |
| 文件夹 | inbox/sent/drafts | + archive/deleted/junk/outbox/root |
| 错误文本 | 回显 `MessageText` | 只回错误码，原文进 opt-in 日志 |
