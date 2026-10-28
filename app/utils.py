from __future__ import annotations

import csv
import hashlib
import io
import json
import os
import re
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Iterable

from charset_normalizer import from_bytes

INVALID_FILENAME = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def sanitize_filename(name: str, fallback: str = "file", max_length: int = 80) -> str:
    stem = Path(name).stem if Path(name).suffix else name
    stem = INVALID_FILENAME.sub("_", stem).strip(" .")
    stem = re.sub(r"\s+", " ", stem)
    if not stem:
        stem = fallback
    return stem[:max_length]


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    index = 2
    while True:
        candidate = path.with_name(f"{path.stem}_{index}{path.suffix}")
        if not candidate.exists():
            return candidate
        index += 1


def human_size(size: int) -> str:
    value = float(size)
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:.0f} {unit}" if unit == "B" else f"{value:.1f} {unit}"
        value /= 1024
    return f"{size} B"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def detect_and_read_text(path: Path) -> tuple[str, str]:
    data = path.read_bytes()
    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig"), "utf-8-sig"
    try:
        return data.decode("utf-8"), "utf-8"
    except UnicodeDecodeError:
        pass
    if data.startswith(b"\xff\xfe") or data.startswith(b"\xfe\xff"):
        try:
            return data.decode("utf-16"), "utf-16"
        except UnicodeDecodeError:
            pass
    best = from_bytes(data).best()
    if best is not None:
        try:
            return str(best), best.encoding or "utf-8"
        except (UnicodeDecodeError, LookupError):
            pass
    try:
        return data.decode("gb18030"), "gb18030"
    except UnicodeDecodeError:
        return data.decode("utf-8", errors="replace"), "utf-8(替换非法字符)"


def write_text(path: Path, text: str, encoding: str = "utf-8") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if encoding == "utf-8-sig":
        path.write_text(text, encoding="utf-8-sig")
    else:
        path.write_text(text, encoding="utf-8", newline="\n")


def sniff_csv(text: str) -> tuple[str, str]:
    sample = text[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        delimiter = dialect.delimiter
    except csv.Error:
        delimiter = ","
    return delimiter, text


def read_csv_rows(path: Path) -> tuple[list[str], list[dict[str, str]], str]:
    text, encoding = detect_and_read_text(path)
    delimiter, _ = sniff_csv(text)
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    rows = list(reader)
    if not rows:
        return [], [], encoding
    headers = [str(value).strip() or f"字段{index + 1}" for index, value in enumerate(rows[0])]
    records: list[dict[str, str]] = []
    for raw in rows[1:]:
        if not any(str(value).strip() for value in raw):
            continue
        record = {}
        for index, header in enumerate(headers):
            record[header] = str(raw[index]) if index < len(raw) else ""
        records.append(record)
    return headers, records, encoding


def write_csv(path: Path, headers: list[str], rows: Iterable[Iterable[Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        for row in rows:
            writer.writerow(["" if value is None else value for value in row])


def zip_paths(paths: list[Path], destination: Path, arcname_prefix: str = "") -> Path:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in paths:
            if path.is_file():
                arcname = f"{arcname_prefix}{path.name}" if arcname_prefix else path.name
                archive.write(path, arcname)
            elif path.is_dir():
                for child in path.rglob("*"):
                    if child.is_file():
                        relative = child.relative_to(path)
                        archive.write(child, str(Path(arcname_prefix) / relative))
    return destination


def run_process(
    args: list[str],
    timeout: int,
    cwd: str | Path | None = None,
    env: dict[str, str] | None = None,
) -> subprocess.CompletedProcess[str]:
    process = subprocess.Popen(
        args,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=str(cwd) if cwd else None,
        env=env,
    )
    try:
        stdout, stderr = process.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        if os.name == "nt":
            subprocess.run(
                ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                capture_output=True,
                text=True,
            )
        else:
            process.kill()
        process.communicate()
        raise TimeoutError(f"命令执行超时：{Path(args[0]).name}")
    return subprocess.CompletedProcess(args, process.returncode, stdout, stderr)


def list_process_ids(image_name: str) -> set[int]:
    if os.name != "nt":
        return set()
    try:
        result = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {image_name}", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=10,
        )
    except Exception:
        return set()
    ids: set[int] = set()
    for line in result.stdout.splitlines():
        if not line.startswith('"'):
            continue
        parts = next(csv.reader([line]))
        if len(parts) >= 2 and parts[1].isdigit():
            ids.add(int(parts[1]))
    return ids


def kill_process_ids(ids: Iterable[int]) -> None:
    if os.name != "nt":
        return
    for pid in ids:
        try:
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/T", "/F"],
                capture_output=True,
                text=True,
                timeout=10,
            )
        except Exception:
            pass


def json_dumps(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, indent=2, default=str)


def temp_dir(prefix: str = "fmt-") -> tempfile.TemporaryDirectory[str]:
    return tempfile.TemporaryDirectory(prefix=prefix)