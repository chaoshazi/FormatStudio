from __future__ import annotations

import logging
import multiprocessing
import os
import secrets
import shutil
import socket
import tempfile
import threading
import time
import traceback
import urllib.parse
import urllib.request
import webbrowser
from pathlib import Path

import uvicorn

LOGGER = logging.getLogger("format-studio.desktop")

class DesktopApi:
    def __init__(self) -> None:
        self.window = None
        self.base_url: str | None = None

    def save_file(self, url: str, filename: str = "download") -> dict:
        try:
            if url.startswith("/") and self.base_url:
                url = self.base_url + url
            parsed = urllib.parse.urlsplit(url)
            if parsed.scheme != "http" or parsed.hostname not in {"127.0.0.1", "localhost"}:
                LOGGER.warning("save_file rejected non-local URL: scheme=%s host=%s", parsed.scheme, parsed.hostname)
                return {"ok": False, "error": "invalid download source"}
            if self.window is None:
                return {"ok": False, "error": "window not ready"}
            import webview

            suggested = Path(filename).name or "download"
            selected = self.window.create_file_dialog(webview.FileDialog.SAVE, save_filename=suggested)
            if not selected:
                return {"ok": False, "cancelled": True}
            if isinstance(selected, (list, tuple)):
                target = Path(selected[0])
            else:
                target = Path(selected)
            target.parent.mkdir(parents=True, exist_ok=True)
            partial = target.with_name(target.name + ".part")
            with urllib.request.urlopen(url, timeout=120) as response, partial.open("wb") as handle:
                shutil.copyfileobj(response, handle)
            partial.replace(target)
            return {"ok": True, "path": str(target)}
        except Exception as exc:
            LOGGER.exception("save_file failed")
            return {"ok": False, "error": str(exc)}



def _log_dir() -> Path:
    custom = os.environ.get("FORMAT_STUDIO_DATA_DIR")
    if custom:
        path = Path(custom) / "logs"
    else:
        base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
        path = Path(base) / "FormatStudio" / "logs"
    path.mkdir(parents=True, exist_ok=True)
    return path


def _configure_logging() -> Path:
    log_file = _log_dir() / "startup.log"
    handler = logging.FileHandler(log_file, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s %(message)s"))
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    if not any(isinstance(item, logging.FileHandler) for item in root.handlers):
        root.addHandler(handler)
    return log_file


def _show_error(message: str, log_file: Path | None = None) -> None:
    text = message
    if log_file:
        text += f"\n\nLog: {log_file}"
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(None, text, "Format Studio", 0x10)
    except Exception:
        pass


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _wait_ready(url: str, timeout: float = 60.0) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=1) as response:
                if response.status == 200:
                    return True
        except Exception:
            time.sleep(0.25)
    return False


def _single_instance_mutex():
    if os.name != "nt":
        return object()
    import ctypes

    handle = ctypes.windll.kernel32.CreateMutexW(None, False, "Local\\FormatStudioSingleInstance")
    if not handle:
        return object()
    if ctypes.windll.kernel32.GetLastError() == 183:
        return None
    return handle


def _run() -> None:
    log_file = _configure_logging()
    mutex = _single_instance_mutex()
    if mutex is None:
        _show_error("\u683c\u5f0f\u5de5\u574a\u5df2\u5728\u8fd0\u884c\u3002", log_file)
        return
    token = secrets.token_urlsafe(24)
    os.environ["FORMAT_STUDIO_TOKEN"] = token
    from .config import APP_NAME

    port = _free_port()
    config = uvicorn.Config(
        "app.main:app",
        host="127.0.0.1",
        port=port,
        log_level="info",
        access_log=False,
        log_config=None,
    )
    server = uvicorn.Server(config)

    def run_server() -> None:
        try:
            server.run()
        except BaseException:
            LOGGER.exception("uvicorn server failed")

    thread = threading.Thread(target=run_server, daemon=True, name="uvicorn")
    thread.start()
    base = f"http://127.0.0.1:{port}"
    if not _wait_ready(f"{base}/api/v1/health?token={token}"):
        detail = "\u672c\u5730\u670d\u52a1\u542f\u52a8\u5931\u8d25\u3002"
        if not thread.is_alive():
            detail += "\u540e\u7aef\u8fdb\u7a0b\u5df2\u9000\u51fa\u3002"
        raise RuntimeError(detail)
    url = f"{base}/?token={token}"
    LOGGER.info("server ready at %s", base)
    try:
        import webview

        api = DesktopApi()
        api.base_url = base
        window = webview.create_window(
            APP_NAME,
            url,
            width=1320,
            height=860,
            min_size=(1024, 720),
            text_select=True,
            js_api=api,
        )
        api.window = window
        webview.start()
    except Exception:
        LOGGER.exception("pywebview failed, falling back to browser")
        webbrowser.open(url)
        try:
            while thread.is_alive():
                time.sleep(0.5)
        except KeyboardInterrupt:
            pass
    finally:
        server.should_exit = True
        thread.join(timeout=5)


def main() -> None:
    multiprocessing.freeze_support()
    try:
        _run()
    except BaseException:
        log_file = _log_dir() / "startup.log"
        LOGGER.exception("startup failed")
        _show_error("\u5e94\u7528\u542f\u52a8\u5931\u8d25\uff0c\u8be6\u60c5\u89c1\u65e5\u5fd7\u3002", log_file)


if __name__ == "__main__":
    main()
