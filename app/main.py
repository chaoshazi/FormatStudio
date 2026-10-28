from __future__ import annotations

import asyncio
import json
import logging
import os
import shutil
import threading
import time
import uuid
import zipfile
from contextlib import asynccontextmanager
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, Form, HTTPException, Query, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from starlette.background import BackgroundTask
from starlette.concurrency import run_in_threadpool

from . import db
from .capabilities import capability_for, formats_payload, full_matrix
from .config import (
    APP_NAME,
    APP_TOKEN,
    DATA_DIR,
    FRONTEND_DIST,
    LOG_FILE,
    MAX_FILE_BYTES,
    MAX_FILES,
    MAX_TEMPLATE_BYTES,
    MAX_TEMPLATE_DATA_BYTES,
    MAX_TEMPLATE_ROWS,
    PREVIEW_DIR,
    TASKS_DIR,
    TEMPLATES_DIR,
    TOKEN_DISABLED,
    VERSION,
)
from .formats import EXTENSIONS, REJECTED_EXTENSIONS
from .services.builtin_templates import ensure_builtin_templates
from .services.engines import office_to_pdf, registry
from .services.pdf_render import render_pages
from .services.styles_service import apply_style, import_style_from_sample, normalize_style
from .services.task_manager import cleanup_old_files, manager, mark_interrupted
from .services.template_service import (
    TemplateError,
    create_sample_document,
    fill_template,
    generate_from_template,
    parse_data,
    scan_template,
)
from .utils import json_dumps, sanitize_filename, unique_path, zip_paths

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s %(message)s")
logger = logging.getLogger("format-studio")
try:
    handler = RotatingFileHandler(LOG_FILE, maxBytes=10 * 1024 * 1024, backupCount=5, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s %(message)s"))
    logger.addHandler(handler)
except Exception:
    pass


class SettingsPayload(BaseModel):
    output_dir: str | None = None
    retention_days: int | None = Field(default=None, ge=1, le=365)
    csv_all_text: bool | None = None


class StylePayload(BaseModel):
    name: str
    applies_to: list[str]
    data: dict[str, Any] = Field(default_factory=dict)


def _cleanup_loop() -> None:
    while True:
        try:
            cleanup_old_files()
        except Exception:
            logger.exception("cleanup failed")
        time.sleep(6 * 60 * 60)


@asynccontextmanager
async def lifespan(app: FastAPI):
    db.init_db()
    try:
        ensure_builtin_templates()
    except Exception:
        logger.exception("builtin template generation failed")
    mark_interrupted()
    threading.Thread(target=_cleanup_loop, daemon=True).start()
    yield


app = FastAPI(title=APP_NAME, version=VERSION, lifespan=lifespan)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173", "http://localhost:4173", "http://127.0.0.1:4173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.middleware("http")
async def token_middleware(request: Request, call_next):
    if not TOKEN_DISABLED and APP_TOKEN and request.url.path.startswith("/api/"):
        token = request.headers.get("x-app-token") or request.query_params.get("token")
        if token != APP_TOKEN:
            return JSONResponse({"code": "UNAUTHORIZED", "message": "无效的本地访问令牌"}, status_code=401)
    return await call_next(request)
def _public_task(task: dict[str, Any]) -> dict[str, Any]:
    result = dict(task)
    options = dict(result.get("options") or {})
    options.pop("pdf_password", None)
    options.pop("style_data", None)
    result["options"] = options
    items = []
    for item in result.get("items", []):
        public = dict(item)
        outputs = [Path(path) for path in public.get("outputs", []) if Path(path).exists()]
        public["output_count"] = len(outputs)
        public["download_url"] = f"/api/v1/tasks/{result['id']}/items/{item['id']}/download"
        public["published_dir"] = (public.get("meta") or {}).get("published_dir")
        public["outputs"] = [path.name for path in outputs]
        if len(outputs) == 1 and outputs[0].is_file():
            public["download_name"] = outputs[0].name
        else:
            public["download_name"] = f"{Path(item['source_name']).stem}_\u8f6c\u6362\u7ed3\u679c.zip"
        items.append(public)
    result["items"] = items
    values = [item.get("progress", 0) for item in items]
    result["progress"] = int(sum(values) / len(values)) if values else 0
    result["queue_position"] = db.queue_position(result["id"]) if result["status"] == "queued" else 0
    return result


@app.get("/api/v1/health")
def health() -> dict[str, Any]:
    return {"ok": True, "name": APP_NAME, "version": VERSION, "time": datetime.now().isoformat()}


@app.get("/api/v1/capabilities")
def capabilities() -> dict[str, Any]:
    return {
        "formats": formats_payload(),
        "matrix": full_matrix(),
        "engines": {key: value.to_dict() for key, value in registry().probe().items()},
    }


@app.get("/api/v1/engines")
def engines() -> dict[str, Any]:
    return {key: value.to_dict() for key, value in registry().probe().items()}


@app.post("/api/v1/engines/probe")
def probe_engines() -> dict[str, Any]:
    return {key: value.to_dict() for key, value in registry().probe(force=True).items()}


def _validate_upload(filename: str, size: int) -> str:
    suffix = Path(filename).suffix.lower()
    if suffix in REJECTED_EXTENSIONS:
        raise HTTPException(422, detail={"code": "UNSUPPORTED_PAIR", "message": REJECTED_EXTENSIONS[suffix]})
    if suffix not in EXTENSIONS:
        raise HTTPException(422, detail={"code": "UNSUPPORTED_PAIR", "message": f"不支持的扩展名：{suffix or '无'}"})
    if size > MAX_FILE_BYTES:
        raise HTTPException(422, detail={"code": "FILE_TOO_LARGE", "message": "单个文件不能超过 200MB"})
    return EXTENSIONS[suffix]


@app.post("/api/v1/tasks", status_code=202)
async def create_task(
    files: list[UploadFile] = File(...),
    target_format: str = Form(...),
    mode: str = Form("default"),
    style_id: str | None = Form(None),
    pdf_password: str | None = Form(None),
    csv_all_text: bool = Form(False),
    sheet_name: str | None = Form(None),
) -> dict[str, Any]:
    if not files:
        raise HTTPException(422, detail={"code": "NO_FILES", "message": "请至少上传一个文件"})
    if len(files) > MAX_FILES:
        raise HTTPException(422, detail={"code": "TOO_MANY_FILES", "message": "单次最多 50 个文件"})
    style_data = None
    if style_id:
        style = db.get_style(style_id)
        if not style:
            raise HTTPException(404, detail={"code": "STYLE_NOT_FOUND", "message": "样式模板不存在"})
        style_data = style["data"]
    task_id = uuid.uuid4().hex
    task_dir = TASKS_DIR / task_id
    in_dir = task_dir / "in"
    out_dir = task_dir / "out"
    logs_dir = task_dir / "logs"
    for folder in (in_dir, out_dir, logs_dir):
        folder.mkdir(parents=True, exist_ok=True)
    options = {
        "target_format": target_format,
        "mode": mode,
        "style_id": style_id,
        "style_data": style_data,
        "pdf_password": pdf_password,
        "csv_all_text": csv_all_text,
        "sheet_name": sheet_name,
    }
    items: list[dict[str, Any]] = []
    total = 0
    try:
        for upload in files:
            data = await upload.read()
            total += len(data)
            source_format = _validate_upload(upload.filename or "file", len(data))
            capability = capability_for(source_format, target_format)
            if not capability:
                raise HTTPException(
                    422,
                    detail={
                        "code": "UNSUPPORTED_PAIR",
                        "message": f"v1 不支持 {source_format.upper()} → {target_format.upper()}",
                    },
                )
            chosen_mode = mode if mode in capability["modes"] else capability["modes"][0]
            suffix = Path(upload.filename or "").suffix.lower()
            safe_name = sanitize_filename(upload.filename or f"file-{len(items) + 1}{suffix}")
            target = unique_path(in_dir / f"{safe_name}{suffix}")
            target.write_bytes(data)
            items.append(
                {
                    "id": uuid.uuid4().hex,
                    "task_id": task_id,
                    "source_name": upload.filename or target.name,
                    "source_path": str(target),
                    "source_format": source_format,
                    "target_format": target_format,
                    "mode": chosen_mode,
                }
            )
        if total > 2 * 1024 * 1024 * 1024:
            raise HTTPException(422, detail={"code": "FILE_TOO_LARGE", "message": "单批文件总大小不能超过 2GB"})
        db.create_task(task_id, options, "queued")
        for item in items:
            db.add_item(item)
        db.update_task(task_id, status="queued")
        manager.submit(task_id)
    except HTTPException:
        shutil.rmtree(task_dir, ignore_errors=True)
        raise
    except Exception as exc:
        shutil.rmtree(task_dir, ignore_errors=True)
        logger.exception("create task failed")
        raise HTTPException(500, detail={"code": "TASK_CREATE_FAILED", "message": str(exc)}) from exc
    task = db.get_task(task_id)
    return _public_task(task) if task else {"id": task_id}
@app.get("/api/v1/tasks")
def list_tasks(limit: int = Query(50, ge=1, le=200), offset: int = Query(0, ge=0)) -> list[dict[str, Any]]:
    return [_public_task(task) for task in db.list_tasks(limit=limit, offset=offset)]


@app.get("/api/v1/tasks/{task_id}")
def get_task(task_id: str) -> dict[str, Any]:
    task = db.get_task(task_id)
    if not task:
        raise HTTPException(404, detail={"code": "TASK_NOT_FOUND", "message": "任务不存在"})
    return _public_task(task)


@app.post("/api/v1/tasks/{task_id}/cancel")
def cancel_task(task_id: str) -> dict[str, Any]:
    if not db.get_task(task_id):
        raise HTTPException(404, detail={"code": "TASK_NOT_FOUND", "message": "任务不存在"})
    manager.cancel(task_id)
    return _public_task(db.get_task(task_id) or {})


@app.post("/api/v1/tasks/{task_id}/retry")
def retry_task(task_id: str) -> dict[str, Any]:
    if not db.get_task(task_id):
        raise HTTPException(404, detail={"code": "TASK_NOT_FOUND", "message": "任务不存在"})
    if not manager.retry(task_id):
        raise HTTPException(409, detail={"code": "NOT_RETRYABLE", "message": "该任务没有可重试的文件"})
    return _public_task(db.get_task(task_id) or {})


@app.get("/api/v1/tasks/{task_id}/events")
async def task_events(task_id: str, request: Request):
    if not db.get_task(task_id):
        raise HTTPException(404, detail={"code": "TASK_NOT_FOUND", "message": "任务不存在"})

    async def stream():
        last = ""
        while True:
            if await request.is_disconnected():
                break
            task = db.get_task(task_id)
            if not task:
                break
            payload = json.dumps(_public_task(task), ensure_ascii=False)
            if payload != last:
                last = payload
                yield f"event: task\ndata: {payload}\n\n"
            if task["status"] in {"succeeded", "failed", "canceled", "interrupted", "partial"}:
                yield f"event: done\ndata: {json.dumps({'status': task['status']}, ensure_ascii=False)}\n\n"
                break
            await asyncio.sleep(0.6)
        yield ": heartbeat\n\n"

    return StreamingResponse(
        stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


def _item_outputs(task_id: str, item_id: str) -> tuple[dict[str, Any], list[Path]]:
    task = db.get_task(task_id)
    if not task:
        raise HTTPException(404, detail={"code": "TASK_NOT_FOUND", "message": "任务不存在"})
    item = next((entry for entry in task["items"] if entry["id"] == item_id), None)
    if not item:
        raise HTTPException(404, detail={"code": "ITEM_NOT_FOUND", "message": "文件记录不存在"})
    outputs = [Path(path) for path in item.get("outputs", []) if Path(path).exists()]
    if not outputs:
        raise HTTPException(404, detail={"code": "OUTPUT_NOT_FOUND", "message": "输出文件不存在或已清理"})
    return item, outputs


def _zip_temp(paths: list[Path], name: str) -> tuple[Path, BackgroundTask]:
    temp_dir = PREVIEW_DIR / "downloads"
    temp_dir.mkdir(parents=True, exist_ok=True)
    archive = temp_dir / f"{uuid.uuid4().hex}_{sanitize_filename(name)}.zip"
    zip_paths(paths, archive)
    return archive, BackgroundTask(lambda: archive.unlink(missing_ok=True))


@app.get("/api/v1/tasks/{task_id}/items/{item_id}/download")
def download_item(task_id: str, item_id: str):
    item, outputs = _item_outputs(task_id, item_id)
    if len(outputs) == 1 and outputs[0].is_file():
        return FileResponse(outputs[0], filename=outputs[0].name, media_type="application/octet-stream")
    archive, background = _zip_temp(outputs, f"{Path(item['source_name']).stem}_转换结果")
    return FileResponse(archive, filename=archive.name, media_type="application/zip", background=background)


@app.get("/api/v1/tasks/{task_id}/download")
def download_task(task_id: str):
    task = db.get_task(task_id)
    if not task:
        raise HTTPException(404, detail={"code": "TASK_NOT_FOUND", "message": "任务不存在"})
    outputs: list[Path] = []
    for item in task["items"]:
        outputs.extend(Path(path) for path in item.get("outputs", []) if Path(path).exists())
    if not outputs:
        raise HTTPException(404, detail={"code": "OUTPUT_NOT_FOUND", "message": "没有可下载的输出"})
    archive, background = _zip_temp(outputs, f"任务_{task_id[:6]}_转换结果")
    return FileResponse(archive, filename=archive.name, media_type="application/zip", background=background)


@app.post("/api/v1/tasks/{task_id}/open-folder")
def open_task_folder(task_id: str) -> dict[str, Any]:
    task = db.get_task(task_id)
    if not task:
        raise HTTPException(404, detail={"code": "TASK_NOT_FOUND", "message": "任务不存在"})
    published = None
    for item in task["items"]:
        published = (item.get("meta") or {}).get("published_dir")
        if published:
            break
    target = Path(published) if published else TASKS_DIR / task_id / "out"
    if not target.exists():
        raise HTTPException(404, detail={"code": "OUTPUT_NOT_FOUND", "message": "输出目录不存在"})
    try:
        os.startfile(str(target))  # type: ignore[attr-defined]
    except Exception as exc:
        raise HTTPException(500, detail={"code": "OPEN_FOLDER_FAILED", "message": str(exc)}) from exc
    return {"ok": True, "path": str(target)}
@app.get("/api/v1/templates")
def list_templates() -> list[dict[str, Any]]:
    return db.list_templates()


@app.post("/api/v1/templates", status_code=201)
async def upload_template(file: UploadFile = File(...), name: str | None = Form(None)) -> dict[str, Any]:
    suffix = Path(file.filename or "").suffix.lower()
    if suffix not in {".docx", ".xlsx", ".pptx"}:
        raise HTTPException(422, detail={"code": "INVALID_TEMPLATE", "message": "数据模板仅支持 docx、xlsx、pptx"})
    data = await file.read()
    if len(data) > MAX_TEMPLATE_BYTES:
        raise HTTPException(422, detail={"code": "FILE_TOO_LARGE", "message": "模板不能超过 20MB"})
    template_id = uuid.uuid4().hex
    folder = TEMPLATES_DIR / template_id
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / f"template{suffix}"
    target.write_bytes(data)
    fmt = suffix.lstrip(".")
    try:
        fields = scan_template(target, fmt)
    except TemplateError as exc:
        shutil.rmtree(folder, ignore_errors=True)
        raise HTTPException(422, detail={"code": exc.code, "message": exc.message}) from exc
    record = {
        "id": template_id,
        "name": (name or Path(file.filename or "模板").stem).strip() or "模板",
        "format": fmt,
        "path": str(target),
        "original_name": file.filename or f"template{suffix}",
        "fields": fields,
    }
    db.create_template(record)
    return db.get_template(template_id) or record


@app.get("/api/v1/templates/{template_id}")
def get_template(template_id: str) -> dict[str, Any]:
    record = db.get_template(template_id)
    if not record:
        raise HTTPException(404, detail={"code": "TEMPLATE_NOT_FOUND", "message": "模板不存在"})
    return record


@app.delete("/api/v1/templates/{template_id}")
def delete_template(template_id: str) -> dict[str, Any]:
    record = db.get_template(template_id)
    if not record:
        raise HTTPException(404, detail={"code": "TEMPLATE_NOT_FOUND", "message": "模板不存在"})
    if record.get("builtin"):
        raise HTTPException(409, detail={"code": "BUILTIN_TEMPLATE", "message": "\u5185\u7f6e\u6a21\u677f\u4e0d\u80fd\u5220\u9664"})
    db.delete_template(template_id)
    shutil.rmtree(Path(record["path"]).parent, ignore_errors=True)
    return {"ok": True}


@app.post("/api/v1/templates/{template_id}/scan")
def scan_template_fields(template_id: str) -> dict[str, Any]:
    record = db.get_template(template_id)
    if not record:
        raise HTTPException(404, detail={"code": "TEMPLATE_NOT_FOUND", "message": "模板不存在"})
    try:
        fields = scan_template(Path(record["path"]), record["format"])
    except TemplateError as exc:
        raise HTTPException(422, detail={"code": exc.code, "message": exc.message}) from exc
    db.update_template_fields(template_id, fields)
    return {"fields": fields}


async def _save_temp_upload(upload: UploadFile, folder: Path, max_bytes: int) -> Path:
    data = await upload.read()
    if len(data) > max_bytes:
        raise HTTPException(422, detail={"code": "FILE_TOO_LARGE", "message": "上传文件超过大小限制"})
    folder.mkdir(parents=True, exist_ok=True)
    suffix = Path(upload.filename or "").suffix.lower()
    target = folder / f"data{suffix}"
    target.write_bytes(data)
    return target
@app.post("/api/v1/templates/{template_id}/preview")
async def preview_template(template_id: str, data_file: UploadFile | None = File(None)) -> dict[str, Any]:
    record = db.get_template(template_id)
    if not record:
        raise HTTPException(404, detail={"code": "TEMPLATE_NOT_FOUND", "message": "模板不存在"})
    folder = PREVIEW_DIR / f"template-{template_id}-{uuid.uuid4().hex[:8]}"
    folder.mkdir(parents=True, exist_ok=True)
    values: dict[str, str] = {}
    warnings: list[str] = []
    if data_file is not None and data_file.filename:
        data_path = await _save_temp_upload(data_file, folder / "data", MAX_TEMPLATE_DATA_BYTES)
        try:
            headers, rows = parse_data(data_path)
            if rows:
                values = rows[0]
                warnings.append(f"数据源字段：{', '.join(headers[:20])}")
        except TemplateError as exc:
            raise HTTPException(422, detail={"code": exc.code, "message": exc.message}) from exc
    if not values:
        values = {field["name"]: f"示例{field['name']}" for field in record["fields"]}
    target = folder / f"preview.{record['format']}"
    shutil.copy2(record["path"], target)
    missing = fill_template(target, record["format"], values)
    if missing:
        warnings.append(f"缺少字段：{', '.join(sorted(missing))}")
    pages: list[str] = []
    pdf_path = folder / "preview.pdf"
    try:
        await run_in_threadpool(office_to_pdf, target, pdf_path, "auto")
        rendered = await run_in_threadpool(render_pages, pdf_path, folder, 120, None, 3)
        pages = [f"/api/v1/previews/{folder.name}/{path['path'].name}" for path in rendered]
    except Exception as exc:
        warnings.append(f"预览渲染不可用：{exc}")
    return {"pages": pages, "warnings": warnings, "values": values}


@app.post("/api/v1/templates/{template_id}/generate")
async def generate_template(
    template_id: str,
    data_file: UploadFile = File(...),
    style_id: str | None = Form(None),
) -> dict[str, Any]:
    record = db.get_template(template_id)
    if not record:
        raise HTTPException(404, detail={"code": "TEMPLATE_NOT_FOUND", "message": "模板不存在"})
    generation_id = uuid.uuid4().hex[:12]
    folder = PREVIEW_DIR / "generations" / generation_id
    folder.mkdir(parents=True, exist_ok=True)
    data_path = await _save_temp_upload(data_file, folder / "data", MAX_TEMPLATE_DATA_BYTES)
    try:
        headers, rows = await run_in_threadpool(parse_data, data_path)
    except TemplateError as exc:
        raise HTTPException(422, detail={"code": exc.code, "message": exc.message}) from exc
    if len(rows) > MAX_TEMPLATE_ROWS:
        raise HTTPException(422, detail={"code": "TOO_MANY_ROWS", "message": f"单次最多生成 {MAX_TEMPLATE_ROWS} 份文件"})
    style_data = None
    if style_id:
        style = db.get_style(style_id)
        if style:
            style_data = style["data"]
    out_dir = folder / "out"
    try:
        result = await run_in_threadpool(
            generate_from_template,
            Path(record["path"]),
            record["format"],
            data_path,
            out_dir,
            style_data,
        )
    except TemplateError as exc:
        raise HTTPException(422, detail={"code": exc.code, "message": exc.message}) from exc
    files = [
        {
            "name": path.name,
            "url": f"/api/v1/previews/generations/{generation_id}/out/{path.name}",
        }
        for path in result["outputs"]
    ]
    archive_url = None
    archive_name = None
    if result.get("archive"):
        archive = Path(result["archive"])
        archive_url = f"/api/v1/previews/generations/{generation_id}/out/{archive.name}"
        archive_name = archive.name
    return {
        "generation_id": generation_id,
        "files": files,
        "archive_url": archive_url,
        "archive_name": archive_name,
        "generated": result["generated"],
        "rows": result["rows"],
        "warnings": result["warnings"],
        "headers": headers,
    }
@app.get("/api/v1/styles")
def list_styles() -> list[dict[str, Any]]:
    return db.list_styles()


@app.get("/api/v1/styles/{style_id}")
def get_style(style_id: str) -> dict[str, Any]:
    style = db.get_style(style_id)
    if not style:
        raise HTTPException(404, detail={"code": "STYLE_NOT_FOUND", "message": "样式不存在"})
    return style


@app.post("/api/v1/styles", status_code=201)
def create_style(payload: StylePayload) -> dict[str, Any]:
    style_id = uuid.uuid4().hex
    data = normalize_style({**payload.data, "name": payload.name, "applies_to": payload.applies_to})
    db.save_style(style_id, payload.name, payload.applies_to, data)
    return db.get_style(style_id) or {}


@app.put("/api/v1/styles/{style_id}")
def update_style(style_id: str, payload: StylePayload) -> dict[str, Any]:
    existing = db.get_style(style_id)
    if not existing:
        raise HTTPException(404, detail={"code": "STYLE_NOT_FOUND", "message": "样式不存在"})
    if existing.get("builtin"):
        raise HTTPException(409, detail={"code": "BUILTIN_STYLE", "message": "内置样式不能修改，请复制后编辑"})
    data = normalize_style({**payload.data, "name": payload.name, "applies_to": payload.applies_to})
    db.save_style(style_id, payload.name, payload.applies_to, data)
    return db.get_style(style_id) or {}


@app.delete("/api/v1/styles/{style_id}")
def delete_style(style_id: str) -> dict[str, Any]:
    if not db.delete_style(style_id):
        raise HTTPException(409, detail={"code": "STYLE_DELETE_FAILED", "message": "样式不存在或为内置样式"})
    return {"ok": True}


@app.post("/api/v1/styles/import", status_code=201)
async def import_style(file: UploadFile = File(...), name: str | None = Form(None)) -> dict[str, Any]:
    suffix = Path(file.filename or "").suffix.lower()
    if suffix not in {".docx", ".xlsx", ".pptx"}:
        raise HTTPException(422, detail={"code": "INVALID_TEMPLATE", "message": "仅支持从 docx、xlsx、pptx 导入样式"})
    folder = PREVIEW_DIR / "style-imports" / uuid.uuid4().hex
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / f"sample{suffix}"
    target.write_bytes(await file.read())
    fmt = suffix.lstrip(".")
    try:
        data, warnings = await run_in_threadpool(
            import_style_from_sample, target, fmt, (name or Path(file.filename or "导入样式").stem).strip()
        )
    except Exception as exc:
        raise HTTPException(422, detail={"code": "STYLE_IMPORT_FAILED", "message": str(exc)}) from exc
    style_id = uuid.uuid4().hex
    db.save_style(style_id, data["name"], data["applies_to"], data)
    return {**(db.get_style(style_id) or {}), "warnings": warnings}


@app.post("/api/v1/styles/{style_id}/preview")
async def preview_style(style_id: str, fmt: str | None = Query(None)) -> dict[str, Any]:
    style = db.get_style(style_id)
    if not style:
        raise HTTPException(404, detail={"code": "STYLE_NOT_FOUND", "message": "样式不存在"})
    applies = style.get("applies_to") or ["docx"]
    folder = PREVIEW_DIR / "styles" / f"{style_id}-{uuid.uuid4().hex[:8]}"
    folder.mkdir(parents=True, exist_ok=True)
    warnings = []
    pages: list[str] = []
    if fmt == "pdf":
        chosen = "pdf"
        sample = folder / "sample.docx"
        await run_in_threadpool(create_sample_document, "docx", sample)
        pdf_path = folder / "preview.pdf"
        try:
            await run_in_threadpool(office_to_pdf, sample, pdf_path, "auto")
            warnings.extend(await run_in_threadpool(apply_style, pdf_path, "pdf", style["data"]))
            rendered = await run_in_threadpool(render_pages, pdf_path, folder, 120, None, 2)
            pages = [f"/api/v1/previews/styles/{folder.name}/{path['path'].name}" for path in rendered]
        except Exception as exc:
            warnings.append(f"preview unavailable: {exc}")
    else:
        chosen = fmt if fmt in applies else applies[0]
        sample = folder / f"sample.{chosen}"
        await run_in_threadpool(create_sample_document, chosen, sample)
        try:
            warnings.extend(await run_in_threadpool(apply_style, sample, chosen, style["data"]))
        except Exception as exc:
            warnings.append(f"style apply failed: {exc}")
        pdf_path = folder / "preview.pdf"
        try:
            await run_in_threadpool(office_to_pdf, sample, pdf_path, "auto")
            rendered = await run_in_threadpool(render_pages, pdf_path, folder, 120, None, 2)
            pages = [f"/api/v1/previews/styles/{folder.name}/{path['path'].name}" for path in rendered]
        except Exception as exc:
            warnings.append(f"preview unavailable: {exc}")
    return {"pages": pages, "warnings": warnings, "format": chosen}
@app.get("/api/v1/settings")
def get_settings() -> dict[str, Any]:
    return db.get_settings()


@app.put("/api/v1/settings")
def update_settings(payload: SettingsPayload) -> dict[str, Any]:
    values = payload.model_dump(exclude_none=True)
    if "output_dir" in values:
        path = Path(values["output_dir"]).expanduser()
        try:
            path.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            raise HTTPException(422, detail={"code": "INVALID_OUTPUT_DIR", "message": f"无法创建目录：{exc}"}) from exc
        values["output_dir"] = str(path)
    return db.update_settings(values)


@app.post("/api/v1/system/cleanup")
def cleanup() -> dict[str, Any]:
    return cleanup_old_files()


@app.get("/api/v1/system/diagnostics")
def diagnostics() -> Response:
    folder = PREVIEW_DIR / "diagnostics"
    folder.mkdir(parents=True, exist_ok=True)
    archive = folder / f"diagnostics-{uuid.uuid4().hex[:8]}.zip"
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("settings.json", json_dumps(db.get_settings()))
        zf.writestr("engines.json", json_dumps({key: value.to_dict() for key, value in registry().probe().items()}))
        zf.writestr("version.json", json_dumps({"name": APP_NAME, "version": VERSION, "data_dir": str(DATA_DIR)}))
        if LOG_FILE.exists():
            zf.write(LOG_FILE, "logs/app.log")
    return FileResponse(
        archive,
        filename=archive.name,
        media_type="application/zip",
        background=BackgroundTask(lambda: archive.unlink(missing_ok=True)),
    )


@app.get("/api/v1/previews/{path:path}")
def preview_file(path: str):
    base = PREVIEW_DIR.resolve()
    target = (base / path).resolve()
    if base not in target.parents and target != base:
        raise HTTPException(403, detail={"code": "FORBIDDEN", "message": "非法路径"})
    if not target.is_file():
        raise HTTPException(404, detail={"code": "NOT_FOUND", "message": "预览文件不存在"})
    media = "image/png" if target.suffix.lower() == ".png" else "application/pdf" if target.suffix.lower() == ".pdf" else "application/octet-stream"
    return FileResponse(target, media_type=media)


if FRONTEND_DIST.exists():
    assets = FRONTEND_DIST / "assets"
    if assets.exists():
        app.mount("/assets", StaticFiles(directory=assets), name="assets")


@app.get("/{full_path:path}", include_in_schema=False)
def spa(full_path: str):
    if full_path.startswith("api/"):
        raise HTTPException(404, detail={"code": "NOT_FOUND", "message": "接口不存在"})
    index = FRONTEND_DIST / "index.html"
    if index.exists():
        return FileResponse(index)
    return JSONResponse(
        {
            "name": APP_NAME,
            "version": VERSION,
            "message": "前端尚未构建。请在 web/ 目录运行 npm install 与 npm run build，或使用 npm run dev。",
            "api": "/docs",
        }
    )
