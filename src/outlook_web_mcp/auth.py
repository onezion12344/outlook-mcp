from __future__ import annotations

import asyncio
import base64
import json
import time
import urllib.parse
import webbrowser
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

import httpx
from filelock import FileLock

from .config import OUTLOOK_AUDIENCE, OUTLOOK_SCOPE, Settings
from .credentials import CredentialStore, KeyringCredentialStore
from .errors import AuthenticationError
from .models import CredentialBundle
from .tls import default_context

AUTHORITY = "https://login.microsoftonline.com"
ORGANIZATIONS = "organizations"
REFRESH_EARLY_SECONDS = 300
FALLBACK_VERIFICATION_URI = "https://login.microsoft.com/device"
# The device-code response tells us which URL to open in a browser. Only trust hosts
# that Microsoft actually uses for this flow; anything else falls back to the constant.
VERIFICATION_HOSTS = frozenset(
    {"login.microsoft.com", "login.microsoftonline.com", "login.live.com"}
)


def safe_verification_url(value: str, fallback: str) -> str:
    """Return `value` only if it is an https URL on a Microsoft sign-in host."""
    try:
        parsed = urllib.parse.urlsplit(value)
    except ValueError:
        return fallback
    if parsed.scheme != "https" or not parsed.hostname:
        return fallback
    if parsed.hostname.lower() not in VERIFICATION_HOSTS:
        return fallback
    return value


@dataclass(frozen=True, slots=True)
class DeviceCodePrompt:
    user_code: str
    verification_uri: str
    verification_uri_complete: str
    expires_in: int
    interval: int


@dataclass(frozen=True, slots=True)
class PendingDeviceLogin:
    device_code: str
    prompt: DeviceCodePrompt
    client_id: str
    expires_at: float

    def to_secret_dict(self) -> dict[str, object]:
        return {
            "device_code": self.device_code,
            "user_code": self.prompt.user_code,
            "verification_uri": self.prompt.verification_uri,
            "verification_uri_complete": self.prompt.verification_uri_complete,
            "interval": self.prompt.interval,
            "client_id": self.client_id,
            "expires_at": self.expires_at,
        }

    @classmethod
    def from_secret_dict(cls, value: dict[str, object]) -> PendingDeviceLogin:
        try:
            expires_at = float(str(value["expires_at"]))
            remaining = max(int(expires_at - time.time()), 0)
            prompt = DeviceCodePrompt(
                user_code=str(value["user_code"]),
                verification_uri=str(value["verification_uri"]),
                verification_uri_complete=str(value["verification_uri_complete"]),
                expires_in=remaining,
                interval=max(int(str(value["interval"])), 1),
            )
            return cls(
                device_code=str(value["device_code"]),
                prompt=prompt,
                client_id=str(value["client_id"]),
                expires_at=expires_at,
            )
        except (KeyError, TypeError, ValueError) as exc:
            raise AuthenticationError(
                "Stored pending login is invalid.", "pending_invalid"
            ) from exc


@dataclass(frozen=True, slots=True)
class CachedAccessToken:
    token: str
    expires_at: float


def jwt_claims(token: str) -> dict[str, Any]:
    """Decode non-security JWT metadata; Microsoft still validates the token."""
    try:
        part = token.split(".")[1]
        part += "=" * (-len(part) % 4)
        value = json.loads(base64.urlsafe_b64decode(part.encode("ascii")))
    except (IndexError, ValueError, UnicodeDecodeError, json.JSONDecodeError):
        return {}
    return value if isinstance(value, dict) else {}


def _aad_error(body: object, status: int) -> str:
    if isinstance(body, dict):
        raw = body.get("error_description") or body.get("error")
        if isinstance(raw, str):
            return raw.splitlines()[0][:500]
    return f"HTTP {status}"


class AuthManager:
    def __init__(
        self,
        settings: Settings | None = None,
        store: CredentialStore | None = None,
        http: httpx.AsyncClient | None = None,
    ) -> None:
        self.settings = settings or Settings()
        self.settings.validate()
        self.store = store or KeyringCredentialStore(self.settings.credential_profile)
        self._http = http
        self._owned_http: httpx.AsyncClient | None = None
        self._process_lock = asyncio.Lock()
        self._cached: CachedAccessToken | None = None

    async def __aenter__(self) -> AuthManager:
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
                verify=default_context(),
                trust_env=False,
                proxy=self.settings.proxy,
                headers={"User-Agent": "outlook-web-mcp-server/0.1"},
            )
        return self._owned_http

    @staticmethod
    def _device_url() -> str:
        return f"{AUTHORITY}/{ORGANIZATIONS}/oauth2/v2.0/devicecode"

    @staticmethod
    def _token_url(tenant_id: str) -> str:
        return f"{AUTHORITY}/{tenant_id}/oauth2/v2.0/token"

    async def request_device_code(self) -> tuple[str, DeviceCodePrompt]:
        response = await self._client().post(
            self._device_url(),
            data={"client_id": self.settings.client_id, "scope": OUTLOOK_SCOPE},
        )
        body = self._json(response)
        if response.status_code != 200 or not isinstance(body, dict):
            raise AuthenticationError(
                f"Microsoft device-code request failed: {_aad_error(body, response.status_code)}",
                "device_code_failed",
            )
        device_code = str(body.get("device_code") or "")
        user_code = str(body.get("user_code") or "")
        verification_uri = safe_verification_url(
            str(body.get("verification_uri") or ""), FALLBACK_VERIFICATION_URI
        )
        complete = safe_verification_url(
            str(
                body.get("verification_uri_complete")
                or f"{verification_uri}?{httpx.QueryParams({'otc': user_code})}"
            ),
            f"{verification_uri}?{httpx.QueryParams({'otc': user_code})}",
        )
        if not device_code or not user_code:
            raise AuthenticationError(
                "Microsoft returned an incomplete device code.", "device_code_invalid"
            )
        return device_code, DeviceCodePrompt(
            user_code=user_code,
            verification_uri=verification_uri,
            verification_uri_complete=complete,
            expires_in=int(body.get("expires_in") or 900),
            interval=max(int(body.get("interval") or 5), 1),
        )

    async def login(
        self,
        *,
        open_browser: bool = True,
        on_code: Callable[[DeviceCodePrompt], None] | None = None,
    ) -> CredentialBundle:
        pending = await self.begin_login(open_browser=open_browser, on_code=on_code)
        return await self.complete_login(pending)

    async def begin_login(
        self,
        *,
        open_browser: bool = True,
        on_code: Callable[[DeviceCodePrompt], None] | None = None,
    ) -> PendingDeviceLogin:
        device_code, prompt = await self.request_device_code()
        if on_code is not None:
            on_code(prompt)
        if open_browser:
            webbrowser.open(prompt.verification_uri_complete, new=2)
        return PendingDeviceLogin(
            device_code=device_code,
            prompt=prompt,
            client_id=self.settings.client_id,
            expires_at=time.time() + prompt.expires_in,
        )

    async def complete_login(self, pending: PendingDeviceLogin) -> CredentialBundle:
        if pending.expires_at <= time.time():
            raise AuthenticationError("The Microsoft device code expired.", "device_code_expired")
        response = await self._poll_device_token(
            pending.device_code,
            pending.prompt,
            client_id=pending.client_id,
        )
        refresh_token = str(response.get("refresh_token") or "")
        access_token = str(response.get("access_token") or "")
        if not refresh_token or not access_token:
            raise AuthenticationError(
                "Sign-in did not return the refresh and access tokens needed for Outlook.",
                "token_missing",
            )
        claims = self._validate_access_token(access_token)
        tenant_id = str(claims.get("tid") or "")
        if not tenant_id:
            raise AuthenticationError(
                "The Outlook token did not identify an organization tenant.",
                "tenant_missing",
            )
        bundle = CredentialBundle(
            refresh_token=refresh_token,
            client_id=pending.client_id,
            tenant_id=tenant_id,
            puid=str(claims.get("puid") or ""),
            display_name=str(claims.get("name") or "Outlook User"),
        )
        self.store.save(bundle)
        expires_at = float(
            claims.get("exp") or time.time() + int(response.get("expires_in") or 3600)
        )
        self._cached = CachedAccessToken(access_token, expires_at)
        return bundle

    async def _poll_device_token(
        self,
        device_code: str,
        prompt: DeviceCodePrompt,
        *,
        client_id: str,
    ) -> dict[str, Any]:
        deadline = time.monotonic() + prompt.expires_in
        interval = prompt.interval
        while time.monotonic() < deadline:
            await asyncio.sleep(interval)
            response = await self._client().post(
                self._token_url(ORGANIZATIONS),
                data={
                    "client_id": client_id,
                    "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                    "device_code": device_code,
                },
            )
            body = self._json(response)
            if response.status_code == 200 and isinstance(body, dict):
                return body
            error = body.get("error") if isinstance(body, dict) else None
            if error == "authorization_pending":
                continue
            if error == "slow_down":
                interval += 5
                continue
            if error == "authorization_declined":
                raise AuthenticationError("Microsoft sign-in was declined.", "device_code_declined")
            if error == "expired_token":
                break
            raise AuthenticationError(
                f"Microsoft device authorization failed: {_aad_error(body, response.status_code)}",
                "device_code_failed",
            )
        raise AuthenticationError("The Microsoft device code expired.", "device_code_expired")

    async def _exchange_refresh_token(
        self, bundle: CredentialBundle
    ) -> tuple[str, str | None, dict[str, Any], int]:
        response = await self._client().post(
            self._token_url(bundle.tenant_id),
            data={
                "client_id": bundle.client_id,
                "grant_type": "refresh_token",
                "refresh_token": bundle.refresh_token,
                "scope": OUTLOOK_SCOPE,
            },
        )
        body = self._json(response)
        if (
            response.status_code != 200
            or not isinstance(body, dict)
            or not body.get("access_token")
        ):
            raise AuthenticationError(
                f"Microsoft token refresh failed: {_aad_error(body, response.status_code)}",
                "token_refresh_failed",
            )
        token = str(body["access_token"])
        claims = self._validate_access_token(token)
        rotated = str(body.get("refresh_token") or "") or None
        expires_in = int(body.get("expires_in") or 3600)
        return token, rotated, claims, expires_in

    @staticmethod
    def _validate_access_token(token: str) -> dict[str, Any]:
        claims = jwt_claims(token)
        audience = str(claims.get("aud") or "").rstrip("/")
        if audience != OUTLOOK_AUDIENCE:
            raise AuthenticationError(
                "Microsoft returned a token for an unexpected resource audience.",
                "token_audience_mismatch",
            )
        scopes = set(str(claims.get("scp") or "").split())
        if not ({"OWA.AccessAsUser.All", "Mail.ReadWrite"} & scopes):
            raise AuthenticationError(
                "The Outlook token does not include a usable mailbox scope.",
                "token_scope_missing",
            )
        return claims

    async def get_access_token(self, *, force: bool = False) -> str:
        cached = self._cached
        if (
            not force
            and cached is not None
            and cached.expires_at - time.time() > REFRESH_EARLY_SECONDS
        ):
            return cached.token
        async with self._process_lock:
            cached = self._cached
            if (
                not force
                and cached is not None
                and cached.expires_at - time.time() > REFRESH_EARLY_SECONDS
            ):
                return cached.token
            with FileLock(self.settings.lock_path, timeout=30):
                bundle = self.store.load()
                if bundle is None:
                    raise AuthenticationError(
                        "Outlook Web is not authenticated. Run `outlook-web-mcp login` first.",
                        "not_authenticated",
                    )
                if bundle.client_id != self.settings.client_id:
                    # A different app registration would silently re-bind the stored
                    # refresh token just because the environment changed.
                    raise AuthenticationError(
                        "Stored Outlook credentials belong to a different client id; "
                        "run `outlook-web-mcp logout` and sign in again.",
                        "client_id_mismatch",
                    )
                token, rotated, claims, expires_in = await self._exchange_refresh_token(bundle)
                changed = False
                if rotated and rotated != bundle.refresh_token:
                    bundle.refresh_token = rotated
                    changed = True
                updated_puid = str(claims.get("puid") or bundle.puid)
                if updated_puid != bundle.puid:
                    bundle.puid = updated_puid
                    changed = True
                if changed:
                    self.store.save(bundle)
                expires_at = float(claims.get("exp") or time.time() + expires_in)
                self._cached = CachedAccessToken(token, expires_at)
                return token

    def identity(self) -> CredentialBundle:
        bundle = self.store.load()
        if bundle is None:
            raise AuthenticationError(
                "Outlook Web is not authenticated. Run `outlook-web-mcp login` first.",
                "not_authenticated",
            )
        return bundle

    def status(self) -> dict[str, object]:
        bundle = self.store.load()
        if bundle is None:
            return {"authenticated": False, "profile": self.settings.credential_profile}
        expiry = self._cached.expires_at if self._cached is not None else None
        return {
            "authenticated": True,
            "profile": self.settings.credential_profile,
            "tenant_id": bundle.tenant_id,
            "display_name": bundle.display_name,
            "token_expires_at": expiry,
            "token_expired": expiry is not None and expiry <= time.time(),
        }

    async def ensure_status(self) -> dict[str, object]:
        before = self.status()
        if not before.get("authenticated"):
            return {
                **before,
                "usable": False,
                "token_refreshed": False,
                "reauthentication_required": True,
            }
        previous_expiry = before.get("token_expires_at")
        try:
            await self.get_access_token()
        except AuthenticationError as exc:
            return {
                **self.status(),
                "usable": False,
                "token_refreshed": False,
                "reauthentication_required": True,
                "refresh_error_code": exc.code,
            }
        current = self.status()
        return {
            **current,
            "usable": not bool(current.get("token_expired")),
            "token_refreshed": current.get("token_expires_at") != previous_expiry,
            "reauthentication_required": False,
        }

    def logout(self) -> None:
        self._cached = None
        self.store.delete()

    @staticmethod
    def _json(response: httpx.Response) -> object:
        try:
            return response.json()
        except json.JSONDecodeError:
            return {}
