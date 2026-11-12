from __future__ import annotations

import httpx

from outlook_web_mcp.api import OutlookWebApi
from outlook_web_mcp.auth import AuthManager
from outlook_web_mcp.config import Settings
from outlook_web_mcp.proxy import parse_scutil_proxy, resolve_proxy

# The owner's real value: it made httpx raise while constructing the client.
NASTY_NO_PROXY = "localhost,.localhost,127.0.0.1,127.0.0.0/8,[::1],::1"


def test_parse_scutil_enabled_https_proxy() -> None:
    text = """<dictionary> {
  HTTPEnable : 1
  HTTPPort : 7897
  HTTPProxy : 127.0.0.1
  HTTPSEnable : 1
  HTTPSPort : 7897
  HTTPSProxy : 127.0.0.1
}"""
    assert parse_scutil_proxy(text) == "http://127.0.0.1:7897"


def test_parse_scutil_disabled_returns_none() -> None:
    assert parse_scutil_proxy("<dictionary> {\n  HTTPEnable : 0\n}") is None


def test_parse_scutil_garbage_returns_none() -> None:
    assert parse_scutil_proxy("not a plist at all") is None
    assert parse_scutil_proxy("") is None


def test_explicit_value_wins() -> None:
    assert resolve_proxy("http://127.0.0.1:8888") == "http://127.0.0.1:8888"


def test_resolve_never_raises(monkeypatch) -> None:
    # Reads the live system proxy, so assert the contract, not a specific port:
    # the owner changes proxy ports (7897 -> 1080) and that must not break tests.
    monkeypatch.delenv("OUTLOOK_WEB_PROXY", raising=False)
    value = resolve_proxy(None)
    assert value is None or (isinstance(value, str) and value.startswith(("http://", "https://")))


def test_client_construction_survives_nasty_no_proxy(monkeypatch) -> None:
    monkeypatch.setenv("NO_PROXY", NASTY_NO_PROXY)
    monkeypatch.setenv("no_proxy", NASTY_NO_PROXY)
    settings = Settings(credential_profile="proxy-test")
    api = OutlookWebApi(AuthManager(settings), settings)
    assert api._client() is not None


def test_trust_env_disabled_and_proxy_explicit(monkeypatch) -> None:
    captured: dict[str, object] = {}

    class SpyClient:
        def __init__(self, **kwargs: object) -> None:
            captured.update(kwargs)

        async def aclose(self) -> None:
            return None

    monkeypatch.setattr(httpx, "AsyncClient", SpyClient)
    monkeypatch.setenv("OUTLOOK_WEB_PROXY", "http://127.0.0.1:8888")
    settings = Settings(credential_profile="proxy-spy")
    OutlookWebApi(AuthManager(settings), settings)._client()
    assert captured["trust_env"] is False
    assert captured["proxy"] == "http://127.0.0.1:8888"
    assert captured["verify"] is not False


def test_auth_client_also_disables_trust_env(monkeypatch) -> None:
    captured: dict[str, object] = {}

    class SpyClient:
        def __init__(self, **kwargs: object) -> None:
            captured.update(kwargs)

        async def aclose(self) -> None:
            return None

    monkeypatch.setattr(httpx, "AsyncClient", SpyClient)
    monkeypatch.setenv("OUTLOOK_WEB_PROXY", "http://127.0.0.1:8888")
    AuthManager(Settings(credential_profile="proxy-spy-auth"))._client()
    assert captured["trust_env"] is False
    assert captured["proxy"] == "http://127.0.0.1:8888"
