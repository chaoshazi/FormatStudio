from __future__ import annotations

import json
import sqlite3
import threading
from datetime import datetime, timezone
from typing import Any, Iterable

from .config import DB_PATH, DEFAULT_OUTPUT_DIR, DEFAULT_RETENTION_DAYS
from .presets import builtin_presets

_lock = threading.RLock()

SCHEMA = """
CREATE TABLE IF NOT EXISTS tasks (
    id TEXT PRIMARY KEY,
    status TEXT NOT NULL,
    options TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS items (
    id TEXT PRIMARY KEY,
    task_id TEXT NOT NULL,
    source_name TEXT NOT NULL,
    source_path TEXT NOT NULL,
    source_format TEXT NOT NULL,
    target_format TEXT NOT NULL,
    mode TEXT NOT NULL DEFAULT 'default',
    status TEXT NOT NULL,
    progress INTEGER NOT NULL DEFAULT 0,
    engine TEXT,
    error_code TEXT,
    error_message TEXT,
    warning TEXT,
    outputs TEXT NOT NULL DEFAULT '[]',
    meta TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    FOREIGN KEY(task_id) REFERENCES tasks(id)
);
CREATE INDEX IF NOT EXISTS idx_items_task ON items(task_id);
CREATE TABLE IF NOT EXISTS templates (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    format TEXT NOT NULL,
    path TEXT NOT NULL,
    original_name TEXT NOT NULL,
    fields TEXT NOT NULL DEFAULT '[]',
    builtin INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS styles (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    applies_to TEXT NOT NULL DEFAULT '[]',
    data TEXT NOT NULL,
    builtin INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
"""


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def _json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, default=str)


def _loads(value: str | None, fallback: Any) -> Any:
    if not value:
        return fallback
    try:
        return json.loads(value)
    except (TypeError, ValueError):
        return fallback


def init_db() -> None:
    with _lock, connect() as conn:
        conn.executescript(SCHEMA)
        columns = [row["name"] for row in conn.execute("PRAGMA table_info(templates)")]
        if "builtin" not in columns:
            conn.execute("ALTER TABLE templates ADD COLUMN builtin INTEGER NOT NULL DEFAULT 0")
        conn.execute(
            "INSERT OR IGNORE INTO settings(key, value) VALUES(?, ?)",
            ("output_dir", str(DEFAULT_OUTPUT_DIR)),
        )
        conn.execute(
            "INSERT OR IGNORE INTO settings(key, value) VALUES(?, ?)",
            ("retention_days", str(DEFAULT_RETENTION_DAYS)),
        )
        conn.execute(
            "INSERT OR IGNORE INTO settings(key, value) VALUES(?, ?)",
            ("csv_all_text", "false"),
        )
        for preset in builtin_presets():
            stamp = now_iso()
            conn.execute(
                """INSERT INTO styles
                   (id, name, applies_to, data, builtin, created_at, updated_at)
                   VALUES (?, ?, ?, ?, 1, ?, ?)
                   ON CONFLICT(id) DO UPDATE SET
                     name = excluded.name,
                     applies_to = excluded.applies_to,
                     data = excluded.data,
                     builtin = 1,
                     updated_at = excluded.updated_at""",
                (
                    preset["id"],
                    preset["name"],
                    _json(preset["data"].get("applies_to", [])),
                    _json(preset["data"]),
                    stamp,
                    stamp,
                ),
            )
        conn.commit()


def create_task(task_id: str, options: dict[str, Any], status: str = "queued") -> None:
    stamp = now_iso()
    with _lock, connect() as conn:
        conn.execute(
            "INSERT INTO tasks(id, status, options, created_at, updated_at) VALUES (?, ?, ?, ?, ?)",
            (task_id, status, _json(options), stamp, stamp),
        )
        conn.commit()


def add_item(item: dict[str, Any]) -> None:
    stamp = now_iso()
    with _lock, connect() as conn:
        conn.execute(
            """INSERT INTO items
               (id, task_id, source_name, source_path, source_format, target_format, mode,
                status, progress, outputs, meta, created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0, '[]', '{}', ?, ?)""",
            (
                item["id"],
                item["task_id"],
                item["source_name"],
                item["source_path"],
                item["source_format"],
                item["target_format"],
                item.get("mode", "default"),
                item.get("status", "queued"),
                stamp,
                stamp,
            ),
        )
        conn.commit()


def update_task(task_id: str, **fields: Any) -> None:
    if not fields:
        return
    fields["updated_at"] = now_iso()
    columns = ", ".join(f"{key} = ?" for key in fields)
    values = list(fields.values()) + [task_id]
    with _lock, connect() as conn:
        conn.execute(f"UPDATE tasks SET {columns} WHERE id = ?", values)
        conn.commit()


def update_item(item_id: str, **fields: Any) -> None:
    if not fields:
        return
    if "outputs" in fields and not isinstance(fields["outputs"], str):
        fields["outputs"] = _json(fields["outputs"])
    if "meta" in fields and not isinstance(fields["meta"], str):
        fields["meta"] = _json(fields["meta"])
    fields["updated_at"] = now_iso()
    columns = ", ".join(f"{key} = ?" for key in fields)
    values = list(fields.values()) + [item_id]
    with _lock, connect() as conn:
        conn.execute(f"UPDATE items SET {columns} WHERE id = ?", values)
        conn.commit()


def get_task(task_id: str) -> dict[str, Any] | None:
    with _lock, connect() as conn:
        row = conn.execute("SELECT * FROM tasks WHERE id = ?", (task_id,)).fetchone()
        if not row:
            return None
        task = dict(row)
        task["options"] = _loads(task.get("options"), {})
        items = conn.execute(
            "SELECT * FROM items WHERE task_id = ? ORDER BY created_at", (task_id,)
        ).fetchall()
        task["items"] = [_item_row(item) for item in items]
        return task


def _item_row(row: sqlite3.Row) -> dict[str, Any]:
    item = dict(row)
    item["outputs"] = _loads(item.get("outputs"), [])
    item["meta"] = _loads(item.get("meta"), {})
    return item


def list_tasks(limit: int = 50, offset: int = 0) -> list[dict[str, Any]]:
    with _lock, connect() as conn:
        rows = conn.execute(
            "SELECT * FROM tasks ORDER BY created_at DESC LIMIT ? OFFSET ?",
            (limit, offset),
        ).fetchall()
        result = []
        for row in rows:
            task = dict(row)
            task["options"] = _loads(task.get("options"), {})
            items = conn.execute(
                "SELECT * FROM items WHERE task_id = ? ORDER BY created_at", (task["id"],)
            ).fetchall()
            task["items"] = [_item_row(item) for item in items]
            result.append(task)
        return result


def queue_position(task_id: str) -> int:
    with _lock, connect() as conn:
        rows = conn.execute(
            "SELECT id FROM tasks WHERE status = 'queued' ORDER BY created_at"
        ).fetchall()
        for index, row in enumerate(rows):
            if row["id"] == task_id:
                return index + 1
    return 0


def create_template(record: dict[str, Any]) -> None:
    stamp = now_iso()
    with _lock, connect() as conn:
        conn.execute(
            """INSERT INTO templates
               (id, name, format, path, original_name, fields, builtin, created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                record["id"],
                record["name"],
                record["format"],
                record["path"],
                record["original_name"],
                _json(record.get("fields", [])),
                int(record.get("builtin", 0)),
                stamp,
                stamp,
            ),
        )
        conn.commit()


def list_templates() -> list[dict[str, Any]]:
    with _lock, connect() as conn:
        rows = conn.execute(
            "SELECT * FROM templates ORDER BY builtin DESC, "
            "CASE format WHEN 'docx' THEN 0 WHEN 'xlsx' THEN 1 ELSE 2 END, name COLLATE NOCASE"
        ).fetchall()
        result = []
        for row in rows:
            record = dict(row)
            record["fields"] = _loads(record.get("fields"), [])
            record["builtin"] = bool(record.get("builtin", 0))
            result.append(record)
        return result


def get_template(template_id: str) -> dict[str, Any] | None:
    with _lock, connect() as conn:
        row = conn.execute("SELECT * FROM templates WHERE id = ?", (template_id,)).fetchone()
        if not row:
            return None
        record = dict(row)
        record["fields"] = _loads(record.get("fields"), [])
        record["builtin"] = bool(record.get("builtin", 0))
        return record


def update_template_fields(template_id: str, fields: list[dict[str, Any]]) -> None:
    with _lock, connect() as conn:
        conn.execute(
            "UPDATE templates SET fields = ?, updated_at = ? WHERE id = ?",
            (_json(fields), now_iso(), template_id),
        )
        conn.commit()


def delete_template(template_id: str) -> None:
    with _lock, connect() as conn:
        conn.execute("DELETE FROM templates WHERE id = ?", (template_id,))
        conn.commit()


def list_styles() -> list[dict[str, Any]]:
    with _lock, connect() as conn:
        rows = conn.execute(
            "SELECT * FROM styles ORDER BY builtin DESC, "
            "CASE WHEN applies_to LIKE '%\"docx\"%' THEN 0 "
            "WHEN applies_to LIKE '%\"xlsx\"%' THEN 1 ELSE 2 END, name COLLATE NOCASE"
        ).fetchall()
        result = []
        for row in rows:
            record = dict(row)
            record["applies_to"] = _loads(record.get("applies_to"), [])
            record["data"] = _loads(record.get("data"), {})
            record["builtin"] = bool(record["builtin"])
            result.append(record)
        return result


def get_style(style_id: str) -> dict[str, Any] | None:
    with _lock, connect() as conn:
        row = conn.execute("SELECT * FROM styles WHERE id = ?", (style_id,)).fetchone()
        if not row:
            return None
        record = dict(row)
        record["applies_to"] = _loads(record.get("applies_to"), [])
        record["data"] = _loads(record.get("data"), {})
        record["builtin"] = bool(record["builtin"])
        return record


def save_style(
    style_id: str,
    name: str,
    applies_to: Iterable[str],
    data: dict[str, Any],
    builtin: bool = False,
) -> None:
    stamp = now_iso()
    data = dict(data)
    data["name"] = name
    data["applies_to"] = list(applies_to)
    with _lock, connect() as conn:
        existing = conn.execute("SELECT id FROM styles WHERE id = ?", (style_id,)).fetchone()
        if existing:
            conn.execute(
                "UPDATE styles SET name = ?, applies_to = ?, data = ?, updated_at = ? WHERE id = ?",
                (name, _json(list(applies_to)), _json(data), stamp, style_id),
            )
        else:
            conn.execute(
                """INSERT INTO styles(id, name, applies_to, data, builtin, created_at, updated_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (
                    style_id,
                    name,
                    _json(list(applies_to)),
                    _json(data),
                    int(builtin),
                    stamp,
                    stamp,
                ),
            )
        conn.commit()


def delete_style(style_id: str) -> bool:
    with _lock, connect() as conn:
        row = conn.execute("SELECT builtin FROM styles WHERE id = ?", (style_id,)).fetchone()
        if not row or row["builtin"]:
            return False
        conn.execute("DELETE FROM styles WHERE id = ?", (style_id,))
        conn.commit()
        return True


def get_settings() -> dict[str, Any]:
    with _lock, connect() as conn:
        rows = conn.execute("SELECT key, value FROM settings").fetchall()
    raw = {row["key"]: row["value"] for row in rows}
    return {
        "output_dir": raw.get("output_dir", str(DEFAULT_OUTPUT_DIR)),
        "retention_days": int(raw.get("retention_days", DEFAULT_RETENTION_DAYS)),
        "csv_all_text": raw.get("csv_all_text", "false").lower() == "true",
    }


def update_settings(values: dict[str, Any]) -> dict[str, Any]:
    with _lock, connect() as conn:
        for key, value in values.items():
            if key not in {"output_dir", "retention_days", "csv_all_text"}:
                continue
            if isinstance(value, bool):
                value = "true" if value else "false"
            conn.execute(
                "INSERT INTO settings(key, value) VALUES(?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                (key, str(value)),
            )
        conn.commit()
    return get_settings()


def count_tasks_by_status() -> dict[str, int]:
    with _lock, connect() as conn:
        rows = conn.execute("SELECT status, COUNT(*) AS n FROM tasks GROUP BY status").fetchall()
    return {row["status"]: row["n"] for row in rows}

def upsert_builtin_template(record: dict[str, Any]) -> None:
    stamp = now_iso()
    with _lock, connect() as conn:
        conn.execute(
            """INSERT INTO templates
               (id, name, format, path, original_name, fields, builtin, created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, 1, ?, ?)
               ON CONFLICT(id) DO UPDATE SET
                 name = excluded.name,
                 format = excluded.format,
                 path = excluded.path,
                 original_name = excluded.original_name,
                 fields = excluded.fields,
                 builtin = 1,
                 updated_at = excluded.updated_at""",
            (
                record["id"],
                record["name"],
                record["format"],
                record["path"],
                record["original_name"],
                _json(record.get("fields", [])),
                stamp,
                stamp,
            ),
        )
        conn.commit()
