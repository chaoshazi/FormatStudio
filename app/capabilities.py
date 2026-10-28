from __future__ import annotations

from typing import Any

from .formats import FORMATS

# source -> target -> capability
SUPPORTED: dict[str, dict[str, dict[str, Any]]] = {
    "txt": {
        "md": {"quality": "normal", "engines": ["native"]},
        "docx": {"quality": "normal", "engines": ["native"]},
        "pdf": {"quality": "normal", "engines": ["auto"]},
        "wps": {"quality": "normal", "engines": ["wps"]},
    },
    "md": {
        "txt": {"quality": "content", "engines": ["native"]},
        "docx": {"quality": "normal", "engines": ["native"]},
        "pdf": {"quality": "normal", "engines": ["auto"]},
        "wps": {"quality": "normal", "engines": ["wps"]},
    },
    "docx": {
        "txt": {"quality": "content", "engines": ["native"]},
        "md": {"quality": "normal", "engines": ["native"]},
        "pdf": {"quality": "high", "engines": ["wps", "libreoffice"]},
        "wps": {"quality": "normal", "engines": ["wps"]},
        "xlsx": {"quality": "normal", "engines": ["native"]},
    },
    "pdf": {
        "txt": {"quality": "content", "engines": ["native"]},
        "md": {"quality": "content", "engines": ["native"]},
        "docx": {
            "quality": "high",
            "engines": ["native"],
            "modes": ["editable", "layout"],
        },
        "pptx": {
            "quality": "high",
            "engines": ["native"],
            "modes": ["editable", "layout"],
        },
        "wps": {"quality": "normal", "engines": ["native", "wps"]},
        "xlsx": {"quality": "beta", "engines": ["native"]},
    },
    "wps": {
        "txt": {"quality": "content", "engines": ["wps"]},
        "md": {"quality": "content", "engines": ["wps", "native"]},
        "docx": {"quality": "normal", "engines": ["wps"]},
        "pdf": {"quality": "high", "engines": ["wps"]},
    },
    "csv": {
        "xlsx": {"quality": "normal", "engines": ["native"]},
        "pdf": {"quality": "normal", "engines": ["libreoffice", "wps"]},
        "docx": {"quality": "normal", "engines": ["native"]},
        "md": {"quality": "normal", "engines": ["native"]},
    },
    "xlsx": {
        "csv": {"quality": "normal", "engines": ["native"]},
        "pdf": {"quality": "normal", "engines": ["libreoffice", "wps"]},
        "docx": {"quality": "normal", "engines": ["native"]},
        "md": {"quality": "normal", "engines": ["native"]},
    },
    "pptx": {
        "pdf": {"quality": "high", "engines": ["wps", "libreoffice"]},
    },
}

DISABLED_REASON = "v1 不承诺此方向，请先转换到中间格式（例如 DOCX 或 PDF）"


def capability_for(source: str, target: str) -> dict[str, Any] | None:
    spec = SUPPORTED.get(source, {}).get(target)
    if not spec:
        return None
    modes = spec.get("modes") or ["default"]
    return {
        "source": source,
        "target": target,
        "status": "beta" if spec["quality"] == "beta" else "supported",
        "quality": spec["quality"],
        "engines": list(spec["engines"]),
        "modes": modes,
        "reason": None,
    }


def full_matrix() -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for source in FORMATS:
        for target in FORMATS:
            if source == target:
                continue
            cap = capability_for(source, target)
            if cap:
                rows.append(cap)
            else:
                rows.append(
                    {
                        "source": source,
                        "target": target,
                        "status": "disabled",
                        "quality": None,
                        "engines": [],
                        "modes": [],
                        "reason": DISABLED_REASON,
                    }
                )
    return rows


def formats_payload() -> list[dict[str, str]]:
    return [{"id": key, **value} for key, value in FORMATS.items()]