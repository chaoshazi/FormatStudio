from __future__ import annotations

from pathlib import Path

FORMATS: dict[str, dict[str, str]] = {
    "txt": {"label": "纯文本", "kind": "document", "icon": "FileText"},
    "md": {"label": "Markdown", "kind": "document", "icon": "FileCode2"},
    "docx": {"label": "Word 文档", "kind": "document", "icon": "FileType2"},
    "pdf": {"label": "PDF 文档", "kind": "document", "icon": "File"},
    "wps": {"label": "WPS 文字", "kind": "document", "icon": "FilePen"},
    "csv": {"label": "CSV 表格", "kind": "table", "icon": "Table2"},
    "xlsx": {"label": "Excel 表格", "kind": "table", "icon": "Sheet"},
    "pptx": {"label": "PowerPoint", "kind": "presentation", "icon": "Presentation"},
}

EXTENSIONS: dict[str, str] = {
    ".txt": "txt",
    ".text": "txt",
    ".md": "md",
    ".markdown": "md",
    ".docx": "docx",
    ".pdf": "pdf",
    ".wps": "wps",
    ".csv": "csv",
    ".xlsx": "xlsx",
    ".pptx": "pptx",
}

REJECTED_EXTENSIONS = {
    ".docm": "不支持启用宏的 Word 文档",
    ".xlsm": "不支持启用宏的 Excel 文档",
    ".pptm": "不支持启用宏的 PowerPoint 文档",
    ".et": "v1 不支持 WPS 表格格式 .et",
    ".dps": "v1 不支持 WPS 演示格式 .dps",
}


def detect_format(filename: str) -> str | None:
    return EXTENSIONS.get(Path(filename).suffix.lower())


def extension_for(fmt: str) -> str:
    return {
        "txt": ".txt",
        "md": ".md",
        "docx": ".docx",
        "pdf": ".pdf",
        "wps": ".wps",
        "csv": ".csv",
        "xlsx": ".xlsx",
        "pptx": ".pptx",
    }[fmt]


def safe_stem(filename: str) -> str:
    stem = Path(filename).stem.strip() or "document"
    for char in '\\/:*?"<>|':
        stem = stem.replace(char, "_")
    return stem[:80]