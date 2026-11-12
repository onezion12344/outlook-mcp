from __future__ import annotations

import json
import os
import urllib.parse
import uuid
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import httpx

from .auth import AuthManager, jwt_claims
from .config import OWA_ENDPOINT, OWA_ORIGIN, Settings
from .errors import ApiError
from .models import CredentialBundle
from .tls import default_context

URL_POST_DATA_MAX = 4_096


def _record(value: object) -> dict[str, Any]:
    return value if isinstance(value, dict) else {}


def _records(value: object) -> list[dict[str, Any]]:
    if not isinstance(value, list):
        return []
    return [item for item in value if isinstance(item, dict)]


def response_messages(data: object) -> list[dict[str, Any]]:
    root = _record(data)
    body = _record(root.get("Body"))
    messages = _record(body.get("ResponseMessages"))
    return _records(messages.get("Items"))


def response_items(data: object) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    for message in response_messages(data):
        items.extend(_records(message.get("Items")))
    return items


def root_items(data: object) -> list[dict[str, Any]]:
    for message in response_messages(data):
        root = _record(message.get("RootFolder"))
        items = _records(root.get("Items"))
        if items:
            return items
    return []


def root_folders(data: object) -> list[dict[str, Any]]:
    for message in response_messages(data):
        root = _record(message.get("RootFolder"))
        folders = _records(root.get("Folders"))
        if folders:
            return folders
    return []


def _first_error(value: object) -> tuple[str, str] | None:
    """Return (code, safe_message). safe_message is always the code: Exchange
    MessageText can quote mailbox content and must never reach an agent."""
    if isinstance(value, dict):
        response_class = str(value.get("ResponseClass") or "")
        response_code = str(value.get("ResponseCode") or "")
        if response_class.casefold() == "error" or (response_code and response_code != "NoError"):
            code = response_code or "owa_error"
            return code, code
        error_code = value.get("ErrorCode")
        if error_code not in (None, 0, "0", "NoError"):
            code = str(error_code)
            return code, code
        for child in value.values():
            error = _first_error(child)
            if error is not None:
                return error
    elif isinstance(value, list):
        for child in value:
            error = _first_error(child)
            if error is not None:
                return error
    return None


def _raw_error_detail(value: object) -> str:
    """Best-effort raw error text for the opt-in local log. Never returned upward."""
    if isinstance(value, dict):
        for key in ("MessageText", "FaultMessage", "Message"):
            text = value.get(key)
            if isinstance(text, str) and text.strip():
                return text[:2000]
        for child in value.values():
            found = _raw_error_detail(child)
            if found:
                return found
    elif isinstance(value, list):
        for child in value:
            found = _raw_error_detail(child)
            if found:
                return found
    return ""


class OutlookWebApi:
    """Thin client for the private EWS-JSON endpoint used by Outlook Web."""

    def __init__(
        self,
        auth: AuthManager,
        settings: Settings | None = None,
        http: httpx.AsyncClient | None = None,
    ) -> None:
        self.auth = auth
        self.settings = settings or auth.settings
        self._http = http
        self._owned_http: httpx.AsyncClient | None = None
        self._session_id = str(uuid.uuid4())

    async def __aenter__(self) -> OutlookWebApi:
        return self

    async def __aexit__(self, *_: object) -> None:
        await self.aclose()

    async def aclose(self) -> None:
        if self._owned_http is not None:
            await self._owned_http.aclose()
            self._owned_http = None

    def _client(self) -> httpx.AsyncClient:
        if self._http is not None:
            return self._http
        if self._owned_http is None:
            self._owned_http = httpx.AsyncClient(
                timeout=self.settings.request_timeout_seconds,
                follow_redirects=False,
                # Explicit trust root: SSL_CERT_FILE in the environment must not swap it.
                verify=default_context(),
                # Env proxy parsing is disabled on purpose: this machine's NO_PROXY
                # contains "[::1]", which makes httpx raise InvalidURL at construction.
                # The proxy is resolved explicitly instead (see proxy.resolve_proxy).
                trust_env=False,
                proxy=self.settings.proxy,
                headers={"User-Agent": "outlook-web-mcp-server/0.1"},
            )
        return self._owned_http

    def envelope(self, action: str, body: dict[str, Any]) -> dict[str, Any]:
        return {
            "__type": f"{action}JsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2018_01_08",
                "TimeZoneContext": {
                    "__type": "TimeZoneContext:#Exchange",
                    "TimeZoneDefinition": {
                        "__type": "TimeZoneDefinitionType:#Exchange",
                        "Id": self.settings.time_zone_id,
                    },
                },
            },
            "Body": body,
        }

    async def call(
        self,
        action: str,
        body: dict[str, Any],
        *,
        app: str = "Mail",
    ) -> dict[str, Any]:
        envelope = self.envelope(action, body)
        response = await self._send(action, envelope, app=app, force_refresh=False)
        if response.status_code == 401:
            response = await self._send(action, envelope, app=app, force_refresh=True)
        if response.request.url.host != urllib.parse.urlsplit(OWA_ORIGIN).hostname:
            # A fixed URL plus no-redirect should make this impossible; this is the
            # tripwire that says so if either of those assumptions ever breaks.
            raise ApiError(
                "Outlook Web answered from an unexpected host; refusing the response.",
                "host_mismatch",
                http_status=response.status_code,
            )
        if 300 <= response.status_code < 400:
            raise ApiError(
                "Outlook Web unexpectedly redirected an API request; "
                "refusing to forward credentials.",
                "unexpected_redirect",
                http_status=response.status_code,
            )
        if len(response.content) > self.settings.max_response_bytes:
            raise ApiError(
                "Outlook Web response exceeded the configured size limit.",
                "response_too_large",
                http_status=response.status_code,
            )
        try:
            parsed = response.json()
        except json.JSONDecodeError as exc:
            raise ApiError(
                f"Outlook Web returned a non-JSON response (HTTP {response.status_code}).",
                "invalid_json",
                http_status=response.status_code,
            ) from exc
        if not isinstance(parsed, dict):
            raise ApiError(
                "Outlook Web returned an unexpected JSON shape.",
                "invalid_shape",
                http_status=response.status_code,
            )
        if response.status_code != 200:
            error = _first_error(parsed)
            response_code = error[0] if error else None
            self._log_error(action, parsed)
            code = response_code or f"http_{response.status_code}"
            raise ApiError(
                f"Outlook Web {action} failed [{code}]",
                "http_error",
                http_status=response.status_code,
                response_code=response_code,
            )
        error = _first_error(parsed)
        if error is not None:
            self._log_error(action, parsed)
            raise ApiError(
                f"Outlook Web {action} failed [{error[0]}]",
                "ews_error",
                http_status=response.status_code,
                response_code=error[0],
            )
        return parsed

    async def _send(
        self,
        action: str,
        envelope: dict[str, Any],
        *,
        app: str,
        force_refresh: bool,
    ) -> httpx.Response:
        token = await self.auth.get_access_token(force=force_refresh)
        identity = self.auth.identity()
        raw = json.dumps(envelope, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
        encoded = urllib.parse.quote_from_bytes(raw, safe="")
        headers = self._headers(action, app, token, identity)
        content: bytes | None
        if len(encoded.encode("ascii")) <= URL_POST_DATA_MAX:
            headers["X-OWA-UrlPostData"] = encoded
            content = None
        else:
            content = raw
        url = f"{OWA_ENDPOINT}?action={urllib.parse.quote(action)}&app={urllib.parse.quote(app)}"
        # Stream so a hostile or broken upstream cannot force an allocation past the
        # cap: we stop reading the moment the accumulated bytes exceed it.
        client = self._client()
        async with client.stream("POST", url, headers=headers, content=content) as streamed:
            chunks: list[bytes] = []
            total = 0
            async for chunk in streamed.aiter_bytes():
                total += len(chunk)
                if total > self.settings.max_response_bytes:
                    raise ApiError(
                        "Outlook Web returned a response larger than the configured cap.",
                        "response_too_large",
                        http_status=streamed.status_code,
                    )
                chunks.append(chunk)
            # aiter_bytes() already decoded the body, so the framing headers would lie:
            # keeping them makes the rebuilt Response decompress a second time.
            headers_out = [
                (key, value)
                for key, value in streamed.headers.multi_items()
                if key.lower() not in {"content-encoding", "content-length", "transfer-encoding"}
            ]
            return httpx.Response(
                streamed.status_code,
                headers=headers_out,
                content=b"".join(chunks),
                request=streamed.request,
            )

    def _log_error(self, action: str, parsed: dict[str, Any]) -> None:
        if not self.settings.verbose_errors:
            return
        detail = _raw_error_detail(parsed)
        if not detail:
            return
        path = self.settings.error_log_path
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        line = f"{datetime.now(UTC).isoformat(timespec='seconds')} {action} {detail}\n"
        with open(path, "a", encoding="utf-8") as handle:
            os.chmod(path, 0o600)
            handle.write(line)

    def _headers(
        self,
        action: str,
        app: str,
        token: str,
        identity: CredentialBundle,
    ) -> dict[str, str]:
        claims = jwt_claims(token)
        puid = str(claims.get("puid") or identity.puid)
        tenant_id = str(claims.get("tid") or identity.tenant_id)
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json; charset=utf-8",
            "Accept": "application/json",
            "Action": action,
            "Prefer": (
                'IdType="ImmutableId", exchange.behavior="IncludeThirdPartyOnlineMeetingProviders"'
            ),
            "X-OWA-CorrelationId": str(uuid.uuid4()),
            "X-OWA-SessionId": self._session_id,
            "X-OWA-Hosted-UX": "false",
            "X-Req-Source": app,
            "Origin": OWA_ORIGIN,
        }
        if puid and tenant_id:
            headers["X-AnchorMailbox"] = f"PUID:{puid}@{tenant_id}"
        return headers
