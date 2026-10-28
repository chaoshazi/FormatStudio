from __future__ import annotations

import shutil
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from .. import db
from ..config import HISTORY_RETENTION_DAYS, PREVIEW_DIR, TASKS_DIR
from ..utils import unique_path
from .converter import ConversionError, convert_file


class TaskManager:
    def __init__(self) -> None:
        self._executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="convert")
        self._events: dict[str, threading.Event] = {}
        self._lock = threading.RLock()

    def submit(self, task_id: str) -> None:
        with self._lock:
            if task_id in self._events:
                return
            event = threading.Event()
            self._events[task_id] = event
        self._executor.submit(self._run, task_id, event)

    def cancel(self, task_id: str) -> bool:
        with self._lock:
            event = self._events.get(task_id)
            if event:
                event.set()
        task = db.get_task(task_id)
        if not task:
            return False
        if task["status"] == "queued":
            for item in task["items"]:
                if item["status"] == "queued":
                    db.update_item(item["id"], status="canceled", error_code="CANCELED", error_message="任务已取消")
            db.update_task(task_id, status="canceled")
        elif task["status"] == "running":
            for item in task["items"]:
                if item["status"] == "queued":
                    db.update_item(item["id"], status="canceled", error_code="CANCELED", error_message="任务已取消")
        return True

    def retry(self, task_id: str) -> bool:
        task = db.get_task(task_id)
        if not task:
            return False
        retryable = [item for item in task["items"] if item["status"] in {"failed", "canceled", "interrupted"}]
        if not retryable:
            return False
        for item in retryable:
            db.update_item(
                item["id"],
                status="queued",
                progress=0,
                error_code=None,
                error_message=None,
                warning=None,
                outputs=[],
                meta={},
            )
        db.update_task(task_id, status="queued")
        self.submit(task_id)
        return True

    def _publish(self, task_id: str, task: dict[str, Any], task_dir: Path) -> Path | None:
        settings = db.get_settings()
        root = Path(settings["output_dir"]).expanduser()
        try:
            root.mkdir(parents=True, exist_ok=True)
        except Exception:
            return None
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target = unique_path(root / f"{stamp}_{task_id[:6]}")
        try:
            target.mkdir(parents=True, exist_ok=True)
            source = task_dir / "out"
            if source.exists():
                for child in source.iterdir():
                    destination = target / child.name
                    if child.is_dir():
                        if destination.exists():
                            shutil.rmtree(destination, ignore_errors=True)
                        shutil.copytree(child, destination)
                    else:
                        shutil.copy2(child, destination)
            return target
        except Exception:
            return None

    def _run(self, task_id: str, event: threading.Event) -> None:
        task = db.get_task(task_id)
        if not task:
            self._finish(task_id)
            return
        if task["status"] == "canceled":
            self._finish(task_id)
            return
        db.update_task(task_id, status="running")
        options = task.get("options") or {}
        task_dir = TASKS_DIR / task_id
        out_dir = task_dir / "out"
        out_dir.mkdir(parents=True, exist_ok=True)
        for item in task["items"]:
            if item["status"] in {"succeeded", "failed", "canceled"}:
                continue
            if event.is_set():
                db.update_item(
                    item["id"],
                    status="canceled",
                    error_code="CANCELED",
                    error_message="任务已取消",
                )
                continue
            db.update_item(item["id"], status="running", progress=1, error_code=None, error_message=None)
            last_progress = {"value": 0}

            def update_progress(value: int, item_id: str = item["id"]) -> None:
                value = max(0, min(100, int(value)))
                if value == last_progress["value"] and value not in {100}:
                    return
                last_progress["value"] = value
                db.update_item(item_id, progress=value)

            try:
                result = convert_file(
                    Path(item["source_path"]),
                    item["source_format"],
                    item["target_format"],
                    item.get("mode") or "default",
                    out_dir,
                    options,
                    update_progress,
                    event,
                )
                if event.is_set():
                    for output in result.outputs:
                        try:
                            if output.is_file():
                                output.unlink()
                            elif output.is_dir():
                                shutil.rmtree(output, ignore_errors=True)
                        except OSError:
                            pass
                    db.update_item(
                        item["id"],
                        status="canceled",
                        progress=0,
                        error_code="CANCELED",
                        error_message="任务已取消",
                    )
                    continue
                db.update_item(
                    item["id"],
                    status="succeeded",
                    progress=100,
                    engine=result.engine,
                    outputs=[str(path) for path in result.outputs],
                    meta=result.meta,
                    warning="；".join(result.warnings) if result.warnings else None,
                )
            except ConversionError as exc:
                db.update_item(
                    item["id"],
                    status="canceled" if exc.code == "CANCELED" else "failed",
                    progress=0,
                    error_code=exc.code,
                    error_message=exc.message,
                )
            except Exception as exc:  # noqa: BLE001
                db.update_item(
                    item["id"],
                    status="failed",
                    progress=0,
                    error_code="CONVERSION_FAILED",
                    error_message=str(exc)[:800],
                )
            if event.is_set():
                for remaining in db.get_task(task_id)["items"]:
                    if remaining["status"] == "queued":
                        db.update_item(
                            remaining["id"],
                            status="canceled",
                            error_code="CANCELED",
                            error_message="任务已取消",
                        )
                break
        refreshed = db.get_task(task_id)
        if refreshed:
            statuses = [item["status"] for item in refreshed["items"]]
            if event.is_set():
                final = "canceled"
            elif statuses and all(status == "succeeded" for status in statuses):
                final = "succeeded"
            elif statuses and all(status in {"failed", "canceled"} for status in statuses):
                final = "failed"
            elif any(status == "succeeded" for status in statuses):
                final = "partial"
            else:
                final = "failed"
            db.update_task(task_id, status=final)
            if final in {"succeeded", "partial"}:
                published = self._publish(task_id, refreshed, task_dir)
                if published:
                    for item in db.get_task(task_id)["items"]:
                        meta = item.get("meta") or {}
                        meta["published_dir"] = str(published)
                        db.update_item(item["id"], meta=meta)
        self._finish(task_id)

    def _finish(self, task_id: str) -> None:
        with self._lock:
            self._events.pop(task_id, None)


manager = TaskManager()


def mark_interrupted() -> None:
    for task in db.list_tasks(limit=1000):
        if task["status"] in {"queued", "running"}:
            for item in task["items"]:
                if item["status"] in {"queued", "running"}:
                    db.update_item(
                        item["id"],
                        status="interrupted",
                        error_code="INTERRUPTED",
                        error_message="应用退出导致任务中断，可重试",
                    )
            db.update_task(task["id"], status="interrupted")


def cleanup_old_files() -> dict[str, int]:
    settings = db.get_settings()
    days = int(settings.get("retention_days", 7))
    cutoff = datetime.now(timezone.utc) - timedelta(days=days)
    removed = 0
    if TASKS_DIR.exists():
        for folder in TASKS_DIR.iterdir():
            if not folder.is_dir():
                continue
            try:
                modified = datetime.fromtimestamp(folder.stat().st_mtime, tz=timezone.utc)
            except OSError:
                continue
            if modified < cutoff:
                shutil.rmtree(folder, ignore_errors=True)
                removed += 1
    for sub in ("generations", "downloads", "styles", "style-imports"):
        folder = PREVIEW_DIR / sub
        if not folder.exists():
            continue
        for child in folder.iterdir():
            try:
                modified = datetime.fromtimestamp(child.stat().st_mtime, tz=timezone.utc)
            except OSError:
                continue
            if modified < cutoff:
                if child.is_dir():
                    shutil.rmtree(child, ignore_errors=True)
                else:
                    child.unlink(missing_ok=True)
    history_cutoff = datetime.now(timezone.utc) - timedelta(days=HISTORY_RETENTION_DAYS)
    for task in db.list_tasks(limit=1000):
        try:
            created = datetime.fromisoformat(task["created_at"])
        except Exception:
            continue
        if created < history_cutoff:
            with db._lock, db.connect() as conn:  # noqa: SLF001
                conn.execute("DELETE FROM items WHERE task_id = ?", (task["id"],))
                conn.execute("DELETE FROM tasks WHERE id = ?", (task["id"],))
                conn.commit()
    return {"task_dirs_removed": removed}