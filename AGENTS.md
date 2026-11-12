# AGENTS.md — outlook-web fork

面向在此仓库工作的 agent。改动前先读，不要凭记忆。

- **必读顺序**：本文件 → `STATE.md`（一页交接）→ `FORK.md` → `docs/specs/2026-11-06-fork-design.md` → `PITFALLS.md` → `docs/audit/2026-11-06-security-audit.md`。
- **事实来源**：带日期的 spec 与 PITFALLS 是事实来源；README 是对外文案，可能落后于代码。
- **改动纪律**：一次一个行为；每个改动配一个能判红的测试；`uv run --no-sync pytest -q` 必须绿才提交。
- **禁止**：`git add -A` / `git add .`、`git push`（需 owner 单独批准）、改 keyring 服务名、把令牌写进任何文件或日志、新增出网主机。
- **凭据边界**：device-code token 实际是 `Mail.ReadWrite` 级。"没有 send 工具"是应用层自律，不是权限边界。
- **对外发信**：默认只起草（`create_draft` / `create_reply_draft`）。`send_draft` 只能发**已存在**的草稿，
  且必须由 owner 在终端 `outlook-web-mcp approve <draftId>` 为那份**确切内容**签发一次性 nonce；
  正文只用 owner 说过的内容，草稿改一个字授权即失效。没有通用 `send_email`。
- **搜索口径**：搜索是本地对 subject+sender 的子串匹配，正文永不匹配；结论只有在 `meta.truncated=false` 且 `meta.hitCapReached=false` 时才可信。
- 日期化编辑：改本文档时追加带日期的条目，不删除历史条目。

## 2026-11-06
- 建立 fork，v0.1 计划与 spec 落库；实现票经 free external subagent 链路逐任务执行，台账在 `（本地构建台账）/sdd/progress.md`。

## 2026-11-08
- 新增只读工具 `get_email_html`（第 16 个）：GetItem `BodyType="HTML"` + `html_links_and_text` 解析（连续锚点换行分隔、只留 http(s) href、链接永不跟随）。动机：hCaptcha accessibility magic link 事件的 `get_email` 纯文本把 href 全剥掉，链接即内容的邮件（验证信/OTP）取不到。owner 提出并把"JS 渲染邮件"指定为下一步走截图+OCR（ego 渲染 HTML → 截图 → local OCR skill，不进 MCP——禁新增出网主机）。TDD：`tests/test_get_email_html.py` 5 例先红后绿；`release_check.sh --offline` 全绿；真机验收在真实 hCaptcha 验证信上直接返回 verify_email 链接。
- 同次提交一并落库 c5d97de 的后续增量：`index_attachment_text` 工具 + `extract.py`（附件下载→文本提取→即弃，pdftotext/OCR/Office XML；读不了的留元数据）+ `index.py` 对应存储 + proxy 测试契约化（7897→1080 换端口不再红）。全套 109 测试与 offline release gate 绿。
