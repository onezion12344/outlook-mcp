from __future__ import annotations

import re
import secrets

TOKEN_BYTES = 16
_OUR_IMG = re.compile(
    r'<img\b[^>]*\bdata-owt="(?P<token>[0-9a-f]{32})"[^>]*>\s*',
    re.IGNORECASE,
)
_ANY_IMG_WITH_TOKEN = re.compile(
    r'<img\b[^>]*\bsrc="[^"]*?/o/(?P<token>[0-9a-f]{32})\.gif[^"]*"[^>]*>'
)


def new_token() -> str:
    return secrets.token_hex(TOKEN_BYTES)


def tracking_url(base: str, token: str) -> str:
    return f"{base.rstrip('/')}/o/{token}.gif"


def build_img(url: str, token: str) -> str:
    return f'<img src="{url}" width="1" height="1" alt="" style="display:none" data-owt="{token}">'


def inject(html: str, *, base: str, token: str) -> str:
    """Append our pixel once. Re-injecting replaces the previous one (idempotent)."""
    cleaned = strip(html)
    snippet = build_img(tracking_url(base, token), token)
    lowered = cleaned.lower()
    if "</body>" in lowered:
        index = lowered.rindex("</body>")
        return f"{cleaned[:index]}{snippet}{cleaned[index:]}"
    return f"{cleaned}{snippet}"


def strip(html: str) -> str:
    """Remove our pixel (and any orphan of ours). Never touches other images."""
    without_ours = _OUR_IMG.sub("", html)
    return _ANY_IMG_WITH_TOKEN.sub("", without_ours)


def find_token(html: str) -> str | None:
    match = _OUR_IMG.search(html)
    if match:
        return match.group("token")
    fallback = _ANY_IMG_WITH_TOKEN.search(html)
    return fallback.group("token") if fallback else None
