from __future__ import annotations

import os
from pathlib import Path

APP_NAME = "格式工坊"
APP_SLUG = "FormatStudio"
VERSION = "1.0.0"

ROOT_DIR = Path(__file__).resolve().parent.parent


def _default_data_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA")
    if base:
        return Path(base) / APP_SLUG
    return Path.home() / f".{APP_SLUG.lower()}"


def _default_output_dir() -> Path:
    docs = Path.home() / "Documents"
    if not docs.exists():
        docs = Path.home()
    return docs / APP_SLUG / "输出"


DATA_DIR = Path(os.environ.get("FORMAT_STUDIO_DATA_DIR", _default_data_dir()))
TASKS_DIR = DATA_DIR / "tasks"
TEMPLATES_DIR = DATA_DIR / "templates"
STYLES_DIR = DATA_DIR / "styles"
PREVIEW_DIR = DATA_DIR / "previews"
LOGS_DIR = DATA_DIR / "logs"
DB_PATH = DATA_DIR / "data.db"
DEFAULT_OUTPUT_DIR = Path(
    os.environ.get("FORMAT_STUDIO_OUTPUT_DIR", _default_output_dir())
)

for folder in (
    DATA_DIR,
    TASKS_DIR,
    TEMPLATES_DIR,
    STYLES_DIR,
    PREVIEW_DIR,
    LOGS_DIR,
    DEFAULT_OUTPUT_DIR,
):
    folder.mkdir(parents=True, exist_ok=True)

MAX_FILES = 50
MAX_FILE_BYTES = 200 * 1024 * 1024
MAX_BATCH_BYTES = 2 * 1024 * 1024 * 1024
MAX_TEMPLATE_BYTES = 20 * 1024 * 1024
MAX_TEMPLATE_DATA_BYTES = 50 * 1024 * 1024
MAX_TEMPLATE_ROWS = 200
JOB_TIMEOUT_SECONDS = 10 * 60
DEFAULT_RETENTION_DAYS = 7
HISTORY_RETENTION_DAYS = 90

APP_TOKEN = os.environ.get("FORMAT_STUDIO_TOKEN", "")
TOKEN_DISABLED = os.environ.get("FORMAT_STUDIO_DISABLE_TOKEN", "").lower() in {
    "1",
    "true",
    "yes",
}

FRONTEND_DIST = ROOT_DIR / "web" / "dist"
LOG_FILE = LOGS_DIR / "app.log"