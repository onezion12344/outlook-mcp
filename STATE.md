# STATE — outlook-web fork（2026-11-07）

一句话：**只用 Outlook Web 自己的第一方登录**（无 Azure 应用、无第三方授权）的本地 MCP，
能读/搜/建索引/起草，发信需要你在终端签一次性授权；所有坑都在 `PITFALLS.md`。

## 现在能做什么

```bash
cd ~/Projects/outlook-web-mcp
outlook-web-mcp doctor          # 本地健康自检：6 项，不联网、不打印密钥
outlook-web-mcp status          # 登录态
outlook-web-mcp login           # 首次 device code 登录（Outlook Web 第一方 client id）
outlook-web-mcp approve <id>    # 给一封草稿签发 30 分钟有效的发送授权
outlook-web-mcp pixels-sync     # 从采集端拉已读事件（需先配置 + 部署）
outlook-web-mcp serve           # stdio MCP（Qoder 已注册为 outlook-web）
scripts/release_check.sh        # 回归门：lint/类型/测试/doctor/烟雾，任一失败即红（--offline 跳过联网那段）
```

MCP 工具 17 个：`outlook_auth_status` `search_emails` `list_mail_folders` `get_email`
`get_email_html`（HTML body + http(s) links + text，2026-11-08 加）`set_email_read_state`
`list_calendar_events` `create_draft` `create_reply_draft` `send_draft` `track_status`
`index_sync` `index_bodies` `index_attachment_text` `search_index` `index_status` `index_verify` `pixels_sync`。

## 已验证（真机，非单测）

| 项 | 证据 |
|---|---|
| 读：搜索/文件夹/正文/日历 | 文件夹 32 个；`archive` 25 条；日历 43 条目、合法 31 天窗口 1 条 |
| 索引 | 四文件夹全量、`index_verify` **gap 全为 0**；全量回填 98s；查询 2–3ms。**行数每天在变，不要抄这里的数字**——要看就跑 `index_verify`，它按文件夹报 indexed/mailbox/gap |
| 发送闸门 | 真草稿上 `send_draft` **被拒**（无 nonce），草稿未动；digest 覆盖 To/Cc/**Bcc**/附件/**HTML** + 30 分钟 TTL；收件人缺失的草稿**发送前**即拒（`draft_has_no_recipients`）；create/save 回读 `DisplayTo` 防静默丢收件人（PITFALLS #23） |
| 追踪 | `sent/replied/bounced` + `linked`（自发邮件有 id 级关联）+ `ambiguous`（同主题多封时拒绝归因）；像素采集端**未部署**，但其 `worker.js` 已在本地用真 SQLite + 真 HTTP 跑通契约测试（GIF/`no-store`/401/`since`/sync→归因），改坏 `cache-control` 会红 |
| 状态变更 | `set_email_read_state` 幂等写验证通过 |
| 安全 | 121 CA 钉到 certifi；`trust_env=False`；输出无 token；`doctor` 无密钥 |
| 质量门 | `pytest` + ruff / format / mypy 全绿；CI 矩阵已配置（ubuntu/macOS/Windows + Node 24，`PIXEL_WORKER_REQUIRE_NODE=1` 让缺 Node 变红），但**仓库从未推送，CI 一次都没跑过** |
| 交付 | 私有 fork 63 commits；`NOTICE` 钉上游 `aa41b13`；skill 已装入两个 hub |

## 没做（需要你）

| 事 | 为什么停在这 |
|---|---|
| 像素采集端部署 | 对外部署（`pixel.example.com` + Worker + D1），四步命令在 `collector/pixel-worker/README.md` |
| 推 GitHub | 至今**零远端**，创建私有仓是对外动作 |
| live 草稿 / live 发送 | 前者会给你 Drafts 留残留；后者要你签 nonce，会真发一封 |
| Superhuman 像素实测 | 需要你用 Superhuman 发一封我起草的信，才能读已发送 HTML 判断有无像素 |
| 正文全文索引 | 现在只索引服务端 `Preview`（**每封前 255 字**）；要更深必须正文加密落盘（新增 `cryptography`） |

## 边界（别忘）

- `opened` = "图片被取了一次"，**不是"已读"**；无事件时返回 null。
- 索引是**线索**不是真相：命中后 `get_email` 落实；删除/移动只在窗口完整扫描时自愈。
- 搜索（实时）只匹配主题+发件人，**不匹配正文**；正文只能走索引的 255 字 preview。
- 没有 `send_email`；发信只能发**已存在**的草稿 + 你的 nonce。

## 这份文档里的数字怎么刷新

- 测试数：`uv run --no-sync pytest | tail -1`
- 工具数：MCP `tools/list` 返回的条数（或 `grep -c "^async def " src/outlook_web_mcp/server.py`）
- 索引行数与缺口：`index_verify`（`gap` 必须为 0 才叫"全量"）
- 全量门禁：`scripts/release_check.sh`

改了代码就重跑上面四条再改这里的数字；**这个文件里的数字不来自手算**。

## 文档

`PITFALLS.md`（四段式，每条带复现命令；条数以 `grep -c '^### ' PITFALLS.md` 为准）· （人类决策日志，私有维护）（谁决定了什么，含机器默认单列）· `FORK.md`（rebase 政策）· `AGENTS.md` ·
`docs/acceptance-2026-11-06.md`（逐条验收）· `docs/specs/2026-11-06-fork-design.md`（设计）·
`docs/audit/2026-11-06-security-audit.md`（上游审计）
