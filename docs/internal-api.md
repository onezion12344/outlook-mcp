# Private Outlook Web API notes

## 選択した方式

この実装は Outlook Web のプライベート EWS-JSON エンドポイントを利用する。

```text
POST https://outlook.cloud.microsoft/owa/service.svc?action=<Action>&app=<Mail|Calendar>
Authorization: Bearer <Outlook audience access token>
X-AnchorMailbox: PUID:<puid>@<tenant-id>
X-OWA-UrlPostData: <percent-encoded JSON request, small requests only>
```

大きなリクエストは JSON を本文に置く。リダイレクトは追跡せず、固定ホスト以外へ Bearer
トークンを転送しない。Cookie、`X-OWA-CANARY`、ブラウザープロファイルの読み取りは不要。

認証は Microsoft identity platform のデバイスコード／更新トークンを使い、audience が
`https://outlook.office.com` であることと、メールボックス用 scope があることを毎回検証する。
更新トークンだけを OS keyring に分割保存し、アクセストークンはプロセス内キャッシュに限る。

## 公開 API との区別

- Microsoft Graph mail API は公式の長期運用向け選択肢だが、このプロジェクトでは利用しない。
- 旧 Outlook REST `/api/v2.0` は廃止済みであり、この実装では利用しない。
- Exchange Online の SOAP EWS も利用しない。`service.svc` の JSON dialect は EWS 型名を
  含むが、SOAP EWS エンドポイントとは別物である。

参考資料:

- Microsoft Graph mail overview: https://learn.microsoft.com/graph/api/resources/mail-api-overview
- Outlook REST retirement: https://learn.microsoft.com/outlook/rest/compare-graph
- Exchange Online EWS retirement: https://techcommunity.microsoft.com/blog/exchange/retirement-of-exchange-web-services-in-exchange-online/3924440
- MIT licensed cookie/canary protocol reference: https://github.com/nhype/owa-exchange-mcp
- Bearer-token protocol observation: https://github.com/jlentink/outlook-web-mcp

最後のリポジトリにはライセンス表記がないため、コードは転載せず、通信形式の確認材料として
のみ参照した。このプロジェクトの payload、認証、資格情報保存、安全境界は独自実装である。

## 実装する安全境界

- 読み取り操作で既読状態を変えない。
- `set_email_read_state` は指定された1件の `IsRead` だけを更新する。
- `create_reply_draft` は `MessageDisposition=SaveOnly` で Drafts に保存し、送信しない。
- 送信、削除、移動、アーカイブ用 MCP ツールを公開しない。
- メール本文・件名・会議情報を信頼済み命令として扱わない。
- API 応答サイズ、検索件数、フォルダー走査数、カレンダー期間を制限する。
- API エラーにトークン、リクエスト本文、メール内容を含めない。

## 互換性上の差分

COM 版の `emailId` / `folderId` は Outlook Web の immutable ID へ変わるため、以前保存した
COM Entry ID は再利用できない。`storeId` も `outlook-web` 固定になる。MCP ツール名と主要な
JSON フィールドは維持しているので、ID を永続保存していない呼び出し側は設定コマンドの差し替え
だけで移行できる。
