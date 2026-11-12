# FORK.md — 怎么维护这个 fork

## Pin 与同步
- 基线 commit 记在 `NOTICE`。同步上游：
  `git fetch upstream && git rebase upstream/main`
- 我们的改动保持为一小叠 patch，**一个 patch 一个行为**；不整文件重写上游代码。
- 上游改动与我们的 patch 冲突时，先读上游意图，再决定 re-apply 还是放弃我们那条。

## 不可破坏的约定（上游 + 本 fork）
- keyring 服务名固定 `ma-nakaya/outlook-web-mcp-server`：改名 = 踢掉已登录态。
- 出网主机只有 `outlook.cloud.microsoft` 与 `login.microsoftonline.com`。
- 不暴露 delete / move / archive 工具，且没有引入计划。
- 发送只有一条路：`send_draft`（只发已存在的草稿）+ owner 在终端签发的一次性 nonce；没有通用 `send_email`，`SendOnly` 全仓必须 0 命中。
- 访问令牌只在进程内存；刷新令牌只在 OS keyring。
- TLS 信任根钉死 `certifi`（不认 `SSL_CERT_FILE`），但保留跟随系统代理（owner 自有出口）。

## 版本切分
- v0.1 = 安全底线 + 诚实检索（`PITFALLS.md`）
- v0.2 = 草稿优先：`create_draft` + `track_status`（sent/replied/bounced）+ **授权后**的 `send_draft`
- v0.3 = 本地 SQLite/FTS5 索引

**owner 决定（2026-11-06，二次修订）**：默认只起草、由本人在 Superhuman 发送；
**他明确授权某一封时**，可通过 `send_draft` 发出，前提是他已在终端为那份确切内容签发 nonce。

## 已知上游缺陷
见 `docs/specs/2026-11-06-fork-design.md` 的 F1–F16，以及 `PITFALLS.md`。
