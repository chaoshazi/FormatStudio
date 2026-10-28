from __future__ import annotations

import functools
import threading
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pytest

from app.desktop import DesktopApi


class FakeWindow:
    def __init__(self, target: str | None):
        self.target = target
        self.calls = []

    def create_file_dialog(self, *args, **kwargs):
        self.calls.append((args, kwargs))
        if self.target is None:
            return None
        return (self.target,)


@pytest.fixture()
def http_file(tmp_path: Path):
    root = tmp_path / "served"
    root.mkdir()
    payload = root / "payload.txt"
    payload.write_text("download payload", encoding="utf-8")
    handler = functools.partial(SimpleHTTPRequestHandler, directory=str(root))
    server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield f"http://127.0.0.1:{server.server_port}/payload.txt", payload
    finally:
        server.shutdown()
        server.server_close()
        thread.join(timeout=5)


def test_save_absolute_url(http_file, tmp_path: Path):
    url, payload = http_file
    target = tmp_path / "saved.txt"
    api = DesktopApi()
    api.window = FakeWindow(str(target))
    result = api.save_file(url, "saved.txt")
    assert result["ok"] is True
    assert target.read_text(encoding="utf-8") == payload.read_text(encoding="utf-8")


def test_save_relative_url_uses_base_url(http_file, tmp_path: Path):
    url, payload = http_file
    base = url.rsplit("/", 1)[0]
    target = tmp_path / "saved-relative.txt"
    api = DesktopApi()
    api.window = FakeWindow(str(target))
    api.base_url = base
    result = api.save_file("/payload.txt", "saved-relative.txt")
    assert result["ok"] is True
    assert target.read_text(encoding="utf-8") == payload.read_text(encoding="utf-8")


def test_save_cancel(tmp_path: Path):
    api = DesktopApi()
    api.window = FakeWindow(None)
    result = api.save_file("http://127.0.0.1:1/x", "x.txt")
    assert result["ok"] is False
    assert result["cancelled"] is True


def test_save_rejects_non_local_url(tmp_path: Path):
    api = DesktopApi()
    api.window = FakeWindow(str(tmp_path / "x.txt"))
    result = api.save_file("https://example.com/file", "x.txt")
    assert result["ok"] is False
    assert result["error"] == "invalid download source"
    assert not (tmp_path / "x.txt").exists()