from __future__ import annotations

import os
import shutil
import threading
import time
import uuid
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Callable

from ..config import DATA_DIR, JOB_TIMEOUT_SECONDS
from ..utils import kill_process_ids, list_process_ids, run_process


class EngineError(RuntimeError):
    def __init__(self, code: str, message: str, retryable: bool = True):
        super().__init__(message)
        self.code = code
        self.message = message
        self.retryable = retryable


@dataclass
class EngineInfo:
    id: str
    label: str
    available: bool
    path: str | None = None
    version: str | None = None
    busy: bool = False
    error: str | None = None

    def to_dict(self) -> dict:
        return asdict(self)


def _find_libreoffice() -> Path | None:
    candidates: list[Path] = []
    env_path = os.environ.get("FORMAT_STUDIO_LO_PATH")
    if env_path:
        candidates.append(Path(env_path))
    candidates.extend(
        [
            Path(__file__).resolve().parents[2] / "vendor" / "libreoffice" / "program" / "soffice.exe",
            Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
            Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
        ]
    )
    which = shutil.which("soffice")
    if which:
        candidates.append(Path(which))
    for path in candidates:
        if path.is_file():
            return path
    return None


def _find_wps() -> Path | None:
    env_path = os.environ.get("FORMAT_STUDIO_WPS_PATH")
    if env_path and Path(env_path).is_file():
        return Path(env_path)
    local = os.environ.get("LOCALAPPDATA")
    roots: list[Path] = []
    if local:
        roots.append(Path(local) / "Kingsoft" / "WPS Office")
    roots.extend(
        [
            Path(r"C:\Program Files\Kingsoft\WPS Office"),
            Path(r"C:\Program Files (x86)\Kingsoft\WPS Office"),
        ]
    )
    for root in roots:
        if not root.exists():
            continue
        matches = sorted(root.glob("*/office6/wps.exe"), reverse=True)
        if matches:
            return matches[0]
    return None


def _file_version(path: Path) -> str | None:
    if os.name != "nt":
        return None
    try:
        import win32api  # type: ignore

        info = win32api.GetFileVersionInfo(str(path), "\\")
        if not info:
            return None
        ms = info.get("FileVersionMS", 0)
        ls = info.get("FileVersionLS", 0)
        return f"{ms >> 16}.{ms & 0xFFFF}.{ls >> 16}.{ls & 0xFFFF}"
    except Exception:
        return None


def _progid_registered(progid: str) -> bool:
    if os.name != "nt":
        return False
    try:
        import winreg  # type: ignore

        winreg.QueryValue(winreg.HKEY_CLASSES_ROOT, f"{progid}\\CLSID")
        return True
    except Exception:
        return False


def _busy_wps() -> bool:
    """Only treat WPS as busy when it has a visible document window.

    WPS keeps several background helper processes alive after the main window is
    closed; those must not block automation.
    """
    names = ("wps.exe", "et.exe", "wpp.exe")
    candidates: set[int] = set()
    for name in names:
        candidates |= list_process_ids(name)
    if not candidates:
        return False
    try:
        import win32gui  # type: ignore
        import win32process  # type: ignore

        visible: set[int] = set()

        def callback(hwnd, _extra):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return True
                if not win32gui.GetWindowText(hwnd).strip():
                    return True
                _thread, pid = win32process.GetWindowThreadProcessId(hwnd)
                visible.add(int(pid))
            except Exception:
                pass
            return True

        win32gui.EnumWindows(callback, None)
        return bool(candidates & visible)
    except Exception:
        return True


class EngineRegistry:
    def __init__(self) -> None:
        self._lock = threading.RLock()
        self._cache: dict[str, EngineInfo] | None = None
        self._cache_at = 0.0

    def probe(self, force: bool = False) -> dict[str, EngineInfo]:
        with self._lock:
            if self._cache is not None and not force and time.time() - self._cache_at < 15:
                return dict(self._cache)
            self._cache = {
                "libreoffice": self._probe_libreoffice(),
                "wps": self._probe_wps(),
            }
            self._cache_at = time.time()
            return dict(self._cache)

    def _probe_libreoffice(self) -> EngineInfo:
        path = _find_libreoffice()
        if not path:
            return EngineInfo(
                id="libreoffice",
                label="LibreOffice",
                available=False,
                error="未找到 soffice.exe，请安装 LibreOffice 或放置到 vendor/libreoffice",
            )
        try:
            result = run_process([str(path), "--version"], timeout=12)
            version = (result.stdout or result.stderr).strip().splitlines()[0]
            return EngineInfo(
                id="libreoffice",
                label="LibreOffice",
                available=result.returncode == 0,
                path=str(path),
                version=version,
                error=None if result.returncode == 0 else (result.stderr or "").strip()[:500],
            )
        except Exception as exc:  # noqa: BLE001
            return EngineInfo(
                id="libreoffice",
                label="LibreOffice",
                available=False,
                path=str(path),
                error=str(exc),
            )

    def _probe_wps(self) -> EngineInfo:
        path = _find_wps()
        registered = _progid_registered("KWPS.Application")
        if not path:
            return EngineInfo(
                id="wps",
                label="WPS Office",
                available=False,
                error="未检测到 WPS Office 12+，.wps 与 WPS 高保真导出不可用",
            )
        return EngineInfo(
            id="wps",
            label="WPS Office",
            available=registered,
            path=str(path),
            version=_file_version(path),
            busy=_busy_wps(),
            error=None if registered else "WPS 已安装但 COM 接口未注册，请修复或重装 WPS",
        )

    def availability(self, engine_id: str) -> EngineInfo:
        return self.probe().get(engine_id, EngineInfo(engine_id, engine_id, False))


_registry: EngineRegistry | None = None


def registry() -> EngineRegistry:
    global _registry
    if _registry is None:
        _registry = EngineRegistry()
    return _registry


def _lo_profile_dir() -> Path:
    folder = DATA_DIR / "lo-profile" / uuid.uuid4().hex[:12]
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def libreoffice_convert(src: Path, dst: Path, timeout: int = JOB_TIMEOUT_SECONDS) -> None:
    info = registry().availability("libreoffice")
    if not info.available or not info.path:
        raise EngineError("ENGINE_UNAVAILABLE", info.error or "LibreOffice 不可用")
    dst.parent.mkdir(parents=True, exist_ok=True)
    outdir = dst.parent / f".lo-{uuid.uuid4().hex[:8]}"
    outdir.mkdir(parents=True, exist_ok=True)
    profile = _lo_profile_dir()
    profile_url = profile.as_posix()
    if not profile_url.startswith("/"):
        profile_url = "/" + profile_url
    args = [
        info.path,
        f"-env:UserInstallation=file://{profile_url}",
        "--headless",
        "--norestore",
        "--nolockcheck",
        "--nodefault",
        "--nofirststartwizard",
        "--convert-to",
        dst.suffix.lstrip(".") or "pdf",
        "--outdir",
        str(outdir),
        str(src),
    ]
    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"
    try:
        result = run_process(args, timeout=timeout, env=env)
        if result.returncode != 0:
            raise EngineError(
                "LIBREOFFICE_FAILED",
                (result.stderr or result.stdout or "LibreOffice 转换失败").strip()[:800],
            )
        produced = outdir / f"{src.stem}{dst.suffix}"
        if not produced.exists():
            candidates = list(outdir.glob(f"*{dst.suffix}"))
            if not candidates:
                raise EngineError("LIBREOFFICE_FAILED", "LibreOffice 未生成目标文件")
            produced = candidates[0]
        if dst.exists():
            dst.unlink()
        shutil.move(str(produced), str(dst))
    finally:
        shutil.rmtree(outdir, ignore_errors=True)
        shutil.rmtree(profile, ignore_errors=True)


def _wps_constants(suffix: str) -> tuple[str, str]:
    if suffix == ".xlsx":
        return "KET.Application", "excel"
    if suffix == ".pptx":
        return "KWPP.Application", "powerpoint"
    return "KWPS.Application", "word"


def _run_wps_thread(fn: Callable[[], None], timeout: int) -> None:
    before = (
        list_process_ids("wps.exe")
        | list_process_ids("et.exe")
        | list_process_ids("wpp.exe")
    )
    errors: list[BaseException] = []

    def target() -> None:
        try:
            fn()
        except BaseException as exc:  # noqa: BLE001
            errors.append(exc)

    thread = threading.Thread(target=target, daemon=True)
    thread.start()
    thread.join(timeout)
    if thread.is_alive():
        new_ids = (
            list_process_ids("wps.exe")
            | list_process_ids("et.exe")
            | list_process_ids("wpp.exe")
        ) - before
        kill_process_ids(new_ids)
        raise EngineError("WPS_TIMEOUT", "WPS 转换超时，已终止本次启动的进程", retryable=True)
    if errors:
        exc = errors[0]
        raise EngineError("WPS_FAILED", str(exc)[:800], retryable=True)


def _wps_word_convert(src: Path, dst: Path, file_format: int) -> None:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    pythoncom.CoInitialize()
    app = None
    doc = None
    try:
        app = win32com.client.DispatchEx("KWPS.Application")
        try:
            app.Visible = False
        except Exception:
            pass
        try:
            app.DisplayAlerts = 0
        except Exception:
            pass
        doc = app.Documents.Open(str(src), ReadOnly=True, AddToRecentFiles=False)
        try:
            doc.SaveAs2(str(dst), FileFormat=file_format)
        except Exception:
            doc.SaveAs(str(dst), FileFormat=file_format)
    finally:
        if doc is not None:
            try:
                doc.Close(0)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _wps_excel_pdf(src: Path, dst: Path) -> None:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    pythoncom.CoInitialize()
    app = None
    workbook = None
    try:
        app = win32com.client.DispatchEx("KET.Application")
        try:
            app.Visible = False
        except Exception:
            pass
        try:
            app.DisplayAlerts = False
        except Exception:
            pass
        workbook = app.Workbooks.Open(str(src), ReadOnly=True, UpdateLinks=0)
        workbook.ExportAsFixedFormat(0, str(dst))
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _wps_ppt_convert(src: Path, dst: Path, save_format: int) -> None:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    pythoncom.CoInitialize()
    app = None
    presentation = None
    try:
        app = win32com.client.DispatchEx("KWPP.Application")
        try:
            app.Visible = False
        except Exception:
            pass
        presentation = app.Presentations.Open(str(src), ReadOnly=True, WithWindow=False)
        presentation.SaveAs(str(dst), save_format)
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def wps_convert(src: Path, dst: Path, target: str, timeout: int = 300) -> None:
    info = registry().availability("wps")
    if not info.available:
        raise EngineError("ENGINE_UNAVAILABLE", info.error or "WPS Office 不可用")
    if info.busy:
        raise EngineError(
            "WPS_BUSY",
            "检测到 WPS 正在运行。请先保存并关闭所有 WPS 文档后再试，避免影响你的工作。",
            retryable=True,
        )
    dst.parent.mkdir(parents=True, exist_ok=True)
    _progid, kind = _wps_constants(src.suffix.lower())

    def operation() -> None:
        if kind == "excel":
            if target == "pdf":
                _wps_excel_pdf(src, dst)
            else:
                raise EngineError("WPS_FAILED", "WPS 表格仅支持导出 PDF")
        elif kind == "powerpoint":
            if target == "pdf":
                _wps_ppt_convert(src, dst, 32)
            else:
                raise EngineError("WPS_FAILED", "WPS 演示仅支持导出 PDF")
        else:
            file_format = {"pdf": 17, "docx": 16, "wps": 0}.get(target)
            if file_format is None:
                raise EngineError("WPS_FAILED", f"不支持 WPS 目标格式：{target}")
            _wps_word_convert(src, dst, file_format)

    _run_wps_thread(operation, timeout)
    if not dst.exists() or dst.stat().st_size == 0:
        raise EngineError("WPS_FAILED", "WPS 未生成目标文件", retryable=True)


def office_to_pdf(src: Path, dst: Path, preferred: str = "auto") -> str:
    if preferred in {"wps", "libreoffice"}:
        order = [preferred]
        order.extend(item for item in ("wps", "libreoffice") if item != preferred)
    elif src.suffix.lower() in {".xlsx", ".xls", ".csv"}:
        order = ["libreoffice", "wps"]
    else:
        order = ["wps", "libreoffice"]

    errors: list[str] = []
    for engine in order:
        info = registry().availability(engine)
        if not info.available or (engine == "wps" and info.busy):
            errors.append(f"{info.label}: {info.error or '不可用'}")
            continue
        try:
            if engine == "libreoffice":
                libreoffice_convert(src, dst)
            else:
                wps_convert(src, dst, "pdf")
            if dst.exists() and dst.stat().st_size > 0:
                return engine
            errors.append(f"{info.label}: 未生成文件")
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{engine}: {exc}")
            if dst.exists():
                try:
                    dst.unlink()
                except OSError:
                    pass
    raise EngineError("ENGINE_UNAVAILABLE", "；".join(errors) or "没有可用的 Office 转 PDF 引擎")


def to_wps(src_docx: Path, dst: Path) -> str:
    wps_convert(src_docx, dst, "wps")
    return "wps"