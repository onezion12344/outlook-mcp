from __future__ import annotations

import ssl

import httpx
import pytest

from outlook_web_mcp.api import OutlookWebApi
from outlook_web_mcp.auth import AuthManager
from outlook_web_mcp.config import Settings
from outlook_web_mcp.tls import default_context

SPY_CALLS: list[dict[str, object]] = []


class SpyClient:
    def __init__(self, **kwargs: object) -> None:
        SPY_CALLS.append(kwargs)

    async def aclose(self) -> None:
        return None


@pytest.fixture
def spy(monkeypatch):
    SPY_CALLS.clear()
    monkeypatch.setattr(httpx, "AsyncClient", SpyClient)
    return SPY_CALLS


def _assert_pinned(kwargs: dict[str, object]) -> None:
    verify = kwargs["verify"]
    assert isinstance(verify, ssl.SSLContext), "verify must be an explicit SSLContext"
    assert verify.get_ca_certs(), "the pinned context must actually carry CA certs"
    assert kwargs["follow_redirects"] is False
    assert kwargs["trust_env"] is False


def test_default_context_is_certifi_backed() -> None:
    from pathlib import Path

    import certifi

    assert Path(certifi.where()).is_file()
    context = default_context()
    assert isinstance(context, ssl.SSLContext)
    assert context.verify_mode is ssl.CERT_REQUIRED
    assert context.get_ca_certs(), "the pinned context must actually carry CA certs"


def test_api_client_pins_ca_bundle(spy) -> None:
    settings = Settings(credential_profile="tls-api-test")
    OutlookWebApi(AuthManager(settings), settings)._client()
    assert len(spy) == 1
    _assert_pinned(spy[0])


def test_auth_client_pins_ca_bundle(spy) -> None:
    AuthManager(Settings(credential_profile="tls-auth-test"))._client()
    assert len(spy) == 1
    _assert_pinned(spy[0])


def test_ssl_cert_file_env_cannot_swap_trust_root(spy, tmp_path, monkeypatch) -> None:
    evil = tmp_path / "evil-ca.pem"
    evil.write_text("-----BEGIN CERTIFICATE-----\nZmFrZQ==\n", encoding="utf-8")
    monkeypatch.setenv("SSL_CERT_FILE", str(evil))

    settings = Settings(credential_profile="tls-env-test")
    OutlookWebApi(AuthManager(settings), settings)._client()
    _assert_pinned(spy[0])
