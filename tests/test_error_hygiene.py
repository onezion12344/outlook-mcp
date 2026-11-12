from __future__ import annotations

import os
from pathlib import Path

import httpx
import pytest
from test_api import StubAuth  # tests/ is not a package; pytest prepends it to sys.path.

from outlook_web_mcp.api import OutlookWebApi
from outlook_web_mcp.config import Settings
from outlook_web_mcp.errors import ApiError


def _error_response(message: str) -> httpx.Response:
    return httpx.Response(
        200,
        json={
            "Body": {
                "ResponseMessages": {
                    "Items": [
                        {
                            "ResponseClass": "Error",
                            "ResponseCode": "ErrorItemNotFound",
                            "MessageText": message,
                        }
                    ]
                }
            }
        },
    )


@pytest.mark.asyncio
async def test_message_text_never_reaches_the_caller() -> None:
    settings = Settings(credential_profile="hygiene-test")
    auth = StubAuth(settings)
    secret = "SECRET-BODY-TEXT 请阅读第 12 封信"

    async with httpx.AsyncClient(
        transport=httpx.MockTransport(lambda _: _error_response(secret))
    ) as client:
        api = OutlookWebApi(auth, settings, client)  # type: ignore[arg-type]
        with pytest.raises(ApiError) as raised:
            await api.call("GetItem", {"__type": "GetItemRequest:#Exchange"})

    text = str(raised.value)
    assert secret not in text
    assert "ErrorItemNotFound" in text
    assert raised.value.response_code == "ErrorItemNotFound"
    assert "Bearer" not in text


@pytest.mark.asyncio
async def test_verbose_log_is_opt_in_and_restricted(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("OUTLOOK_WEB_VERBOSE_ERRORS", "1")
    settings = Settings(credential_profile="hygiene-verbose")
    auth = StubAuth(settings)
    secret = "SECRET-VERBOSE-TEXT"
    monkeypatch.setattr(
        Settings, "error_log_path", property(lambda self: str(tmp_path / "errors.log"))
    )

    async with httpx.AsyncClient(
        transport=httpx.MockTransport(lambda _: _error_response(secret))
    ) as client:
        api = OutlookWebApi(auth, settings, client)  # type: ignore[arg-type]
        with pytest.raises(ApiError):
            await api.call("GetItem", {"__type": "GetItemRequest:#Exchange"})

    log = Path(settings.error_log_path)
    assert log.exists()
    assert secret in log.read_text(encoding="utf-8")
    if os.name == "posix":
        assert oct(log.stat().st_mode & 0o777) == "0o600"
