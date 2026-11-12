from __future__ import annotations

import json
import os

from outlook_web_mcp import pixel, pixel_store

BASE = "https://pixel.example.com"


def test_token_is_32_hex_and_unique() -> None:
    tokens = {pixel.new_token() for _ in range(50)}
    assert len(tokens) == 50
    assert all(len(token) == 32 and set(token) <= set("0123456789abcdef") for token in tokens)


def test_url_shape() -> None:
    token = pixel.new_token()
    assert pixel.tracking_url(BASE, token) == f"{BASE}/o/{token}.gif"
    assert pixel.tracking_url(BASE + "/", token) == f"{BASE}/o/{token}.gif"


def test_inject_before_body_close_and_strip_back() -> None:
    token = pixel.new_token()
    html = "<html><body><p>hi</p></body></html>"
    injected = pixel.inject(html, base=BASE, token=token)
    assert injected.endswith("</body></html>")
    assert f"/o/{token}.gif" in injected
    assert pixel.find_token(injected) == token
    assert pixel.strip(injected) == html


def test_inject_is_idempotent_and_replaces_previous_token() -> None:
    first = pixel.new_token()
    second = pixel.new_token()
    html = pixel.inject("<p>hi</p>", base=BASE, token=first)
    again = pixel.inject(html, base=BASE, token=first)
    assert again.count('data-owt="') == 1
    swapped = pixel.inject(again, base=BASE, token=second)
    assert swapped.count('data-owt="') == 1
    assert pixel.find_token(swapped) == second


def test_strip_leaves_foreign_images_alone() -> None:
    token = pixel.new_token()
    tracking = pixel.build_img(pixel.tracking_url(BASE, token), token)
    html = f'<p><img src="cid:logo" alt="logo"></p>{tracking}'
    stripped = pixel.strip(html)
    assert "cid:logo" in stripped
    assert token not in stripped


def test_find_token_returns_none_for_plain_html() -> None:
    assert pixel.find_token("<p>no images</p>") is None


def test_registry_round_trip(tmp_path) -> None:
    token = pixel.new_token()
    directory = str(tmp_path / "pixels")
    pixel_store.register(
        directory,
        token=token,
        draft_id="draft-1",
        url=pixel.tracking_url(BASE, token),
        at="2026-11-06T10:00:00Z",
    )
    assert pixel_store.known(directory, token) is True
    entry = pixel_store.for_draft(directory, "draft-1")
    assert entry is not None and entry["token"] == token and entry["sentAt"] is None

    assert pixel_store.mark_sent(directory, token, "2026-11-06T11:00:00Z") is True
    assert pixel_store.for_draft(directory, "draft-1")["sentAt"] == "2026-11-06T11:00:00Z"
    if os.name == "posix":
        assert oct(os.stat(tmp_path / "pixels" / "pixels.json").st_mode & 0o777) == "0o600"


def test_unknown_token_and_draft(tmp_path) -> None:
    directory = str(tmp_path / "pixels")
    assert pixel_store.known(directory, "0" * 32) is False
    assert pixel_store.mark_sent(directory, "0" * 32, "x") is False
    assert pixel_store.for_draft(directory, "nope") is None


def test_corrupt_registry_degrades_to_empty(tmp_path) -> None:
    directory = str(tmp_path / "pixels")
    path = tmp_path / "pixels" / "pixels.json"
    path.parent.mkdir(parents=True)
    path.write_text("{not json", encoding="utf-8")
    assert pixel_store.known(directory, "0" * 32) is False
    pixel_store.register(directory, token="a" * 32, draft_id="d", url="u", at="t")
    assert json.loads(path.read_text())["a" * 32]["draftId"] == "d"


def test_events_filter_keeps_only_matching_token() -> None:
    token = "b" * 32
    events = [
        {"token": token, "at": "2026-11-06T10:00:00Z", "ua": "Apple Mail", "ip": "1.2.3.4"},
        {"token": "c" * 32, "at": "2026-11-06T11:00:00Z", "ua": "Gmail"},
    ]
    picked = pixel_store.events_for("/tmp/does-not-matter", token, events)
    assert picked == [{"at": "2026-11-06T10:00:00Z", "ua": "Apple Mail"}]
    assert "ip" not in picked[0]
