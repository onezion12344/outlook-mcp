# 2026-11-08 — Superhuman schedule / Mail MCP 可行性（一手证据）

来源：owner 给的官方帮助页（WebFetch 403，改用 ego lite 页面内读取成功）、
`https://mcp.mail.superhuman.com/mcp` 的未鉴权探测、本机已登录的 Superhuman Web。

## 官方 MCP Server

- 端点 `https://mcp.mail.superhuman.com/mcp`；未带 token 的 initialize 返回
  `401 {"code":401,"detail":"missing-mcp-token"}`，
  `WWW-Authenticate: Bearer ... resource_metadata=https://mcp.mail.superhuman.com/.well-known/oauth-protected-resource/mcp`。
- **资格（原文）**："A Business plan or higher. The MCP Server is not available on the Starter plan."
  另需 Ask AI 已启用、Superhuman Chrome 扩展做认证；企业/托管账号还可能需要管理员通过 CSM 打开。
- 认证由客户端生成登录链接后在浏览器完成（`add_account`），不能由 agent 代签。

## 与 schedule 直接相关的工具

- `send_draft`：**"Supports Smart Send, Scheduled Send, and Undo Send"**；Smart send 是可选参数，
  按收件人最可能打开的时间发送。官方明确建议客户端开启"每次发送需确认"。
- `get_draft_send_status`：查已发送或**已排期**草稿的投递状态。
- `create_or_update_reminder` / `discard_reminder` / `list_reminders`：线程级定时提醒（不是发信排期）。
- `undo_send`：Undo 窗口内撤回。
- 另有 `create_or_update_event` / `get_availability`（会议排期，与邮件定时发送是两件事）。

## 本机账号现状（可证的）

- Superhuman Web 已登录，Ask AI 可见（含 "New: Ask AI can now read email attachments"）。
- 应用正文里搜不到 plan / Business / Starter / billing 字样；`/settings` SPA 路由在 ego lite 里只回退到收件箱，
  **未能从 UI 读出档位**。所以"能不能用官方 MCP"目前是**未定**，不是已否定。
- 早期会话里 owner 说过没有买会员；那与 Starter 一致，但需要本人复核，不能拿旧话当现值。

## 本 fork 的缺口

- `send_draft` 是"批准后立刻发"（SendItem 500 → CreateItem SendAndSaveCopy 回退），**没有延迟/排期参数**。
- Agent 侧的 `AGENTS.md` 禁止新增出网主机；OCR/渲染器已按此规则放在 MCP 之外。
  接入 Superhuman 后端同理：若走逆向路径，调用 `mail.superhuman.com` 属于新主机，
  必须作为 MCP 之外的手工/旁路模块，不塞进 server。

## 待定（owner）

1. 档位：Business+ 吗？若是，官方 MCP 可用，但需要你在浏览器完成 OAuth 登录链接。
2. 若 Starter：走逆向 web。需要确认 "REA 仓库" 的具体路径/URL —— 本机
   `~/Projects` 下没有精确叫 REA 的仓库，最相近的是 `~/Projects/ego-re`
   （ego lite 逆向/审计/改装工程，含 SDK 重建、域名白名单、Linux CDP 桥）。
3. 排期语义要哪一种：① 定时刻发送；② Smart Send（挑最可能被打开的时刻）；
   ③ 只是 Undo/撤回窗口。三者实现完全不同，不能混成一个 "schedule"。
