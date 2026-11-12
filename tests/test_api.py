from __future__ import annotations

import base64
import json
import urllib.parse

import httpx
import pytest

from outlook_web_mcp.api import OutlookWebApi
from outlook_web_mcp.config import Settings
from outlook_web_mcp.models import CredentialBundle


def _jwt(claims: dict[str, object]) -> str:
    raw = json.dumps(claims, separators=(",", ":")).encode()
    payload = base64.urlsafe_b64encode(raw).decode().rstrip("=")
    return f"e30.{payload}.signature"


class StubAuth:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.forces: list[bool] = []
        self.token = _jwt({"puid": "12345", "tid": "tenant-guid"})

    async def get_access_token(self, *, force: bool = False) -> str:
        self.forces.append(force)
        return self.token

    def identity(self) -> CredentialBundle:
        return CredentialBundle("unused", "client", "tenant-guid", puid="12345")


@pytest.mark.asyncio
async def test_call_uses_fixed_owa_headers_and_retries_one_401() -> None:
    settings = Settings(credential_profile="api-test")
    auth = StubAuth(settings)
    requests: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        requests.append(request)
        if len(requests) == 1:
            return httpx.Response(401, json={"error": "expired"})
        return httpx.Response(
            200,
            json={
                "Body": {
                    "ResponseMessages": {
                        "Items": [{"ResponseClass": "Success", "ResponseCode": "NoError"}]
                    }
                }
            },
        )

    async with httpx.AsyncClient(transport=httpx.MockTransport(handler)) as client:
        api = OutlookWebApi(auth, settings, client)  # type: ignore[arg-type]
        result = await api.call(
            "GetFolder",
            {
                "__type": "GetFolderRequest:#Exchange",
                "FolderIds": [],
            },
        )

    assert "Body" in result
    assert auth.forces == [False, True]
    assert len(requests) == 2
    request = requests[-1]
    assert request.url.host == "outlook.cloud.microsoft"
    assert request.url.path == "/owa/service.svc"
    assert request.url.params["action"] == "GetFolder"
    assert request.headers["authorization"] == f"Bearer {auth.token}"
    assert request.headers["x-anchormailbox"] == "PUID:12345@tenant-guid"
    assert request.headers["origin"] == "https://outlook.cloud.microsoft"
    encoded = request.headers["x-owa-urlpostdata"]
    envelope = json.loads(urllib.parse.unquote(encoded))
    assert envelope["__type"] == "GetFolderJsonRequest:#Exchange"
    assert envelope["Header"]["RequestServerVersion"] == "V2018_01_08"


@pytest.mark.asyncio
async def test_ews_error_becomes_sanitized_api_error() -> None:
    settings = Settings(credential_profile="api-error-test")
    auth = StubAuth(settings)

    def handler(_: httpx.Request) -> httpx.Response:
        return httpx.Response(
            200,
            json={
                "Body": {
                    "ResponseMessages": {
                        "Items": [
                            {
                                "ResponseClass": "Error",
                                "ResponseCode": "ErrorInvalidIdMalformed",
                                "MessageText": "The id is malformed",
                            }
                        ]
                    }
                }
            },
        )

    async with httpx.AsyncClient(transport=httpx.MockTransport(handler)) as client:
        api = OutlookWebApi(auth, settings, client)  # type: ignore[arg-type]
        with pytest.raises(Exception, match="ErrorInvalidIdMalformed") as raised:
            await api.call("GetItem", {"__type": "GetItemRequest:#Exchange"})

    assert "Bearer" not in str(raised.value)
    assert "The id is malformed" not in str(raised.value)


@pytest.mark.asyncio
async def test_oversized_response_is_refused_mid_stream() -> None:
    from outlook_web_mcp.errors import ApiError

    settings = Settings(credential_profile="api-size-test", max_response_bytes=1024)
    auth = StubAuth(settings)
    delivered = {"chunks": 0}

    class BigStream(httpx.AsyncByteStream):
        async def __aiter__(self):  # type: ignore[override]
            for _ in range(8):
                delivered["chunks"] += 1
                yield b"x" * 512

    def handler(_: httpx.Request) -> httpx.Response:
        return httpx.Response(200, stream=BigStream())

    async with httpx.AsyncClient(transport=httpx.MockTransport(handler)) as client:
        api = OutlookWebApi(auth, settings, client)  # type: ignore[arg-type]
        with pytest.raises(ApiError) as raised:
            await api.call(
                "GetFolder",
                {"__type": "GetFolderRequest:#Exchange", "FolderIds": []},
            )

    assert raised.value.code == "response_too_large"
    # 1024-byte cap over 512-byte chunks: reading must stop at the third one,
    # not after buffering all eight.
    assert delivered["chunks"] == 3


@pytest.mark.asyncio
async def test_gzipped_response_is_decoded_exactly_once() -> None:
    import gzip

    settings = Settings(credential_profile="api-gzip-test")
    auth = StubAuth(settings)
    payload = {"Body": {"ResponseMessages": {"Items": [{"ResponseClass": "Success"}]}}}
    encoded = gzip.compress(json.dumps(payload).encode())

    def handler(_: httpx.Request) -> httpx.Response:
        return httpx.Response(
            200,
            content=encoded,
            headers={"content-encoding": "gzip", "content-type": "application/json"},
        )

    async with httpx.AsyncClient(transport=httpx.MockTransport(handler)) as client:
        api = OutlookWebApi(auth, settings, client)  # type: ignore[arg-type]
        result = await api.call(
            "GetFolder",
            {"__type": "GetFolderRequest:#Exchange", "FolderIds": []},
        )
    assert result == payload
