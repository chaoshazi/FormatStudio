from __future__ import annotations

import shutil
import tempfile
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from ..capabilities import capability_for
from ..formats import extension_for, safe_stem
from ..utils import unique_path
from .engines import EngineError, office_to_pdf, to_wps, wps_convert
from .metadata import inspect_document
from .pdf_conv import (
    PdfConversionError,
    pdf_to_docx_editable,
    pdf_to_docx_layout,
    pdf_to_md,
    pdf_to_pptx_editable,
    pdf_to_pptx_layout,
    pdf_to_txt,
    pdf_to_xlsx,
)
from .styles_service import apply_style
from .table_conv import (
    csv_to_docx,
    csv_to_md,
    csv_to_xlsx,
    docx_to_xlsx,
    xlsx_to_csv,
    xlsx_to_docx,
    xlsx_to_md,
)
from .text_conv import docx_to_md, docx_to_txt, md_to_docx, md_to_txt, txt_to_docx, txt_to_md


class ConversionError(RuntimeError):
    def __init__(self, code: str, message: str, retryable: bool = False):
        super().__init__(message)
        self.code = code
        self.message = message
        self.retryable = retryable


@dataclass
class ConversionResult:
    outputs: list[Path] = field(default_factory=list)
    engine: str = "native"
    warnings: list[str] = field(default_factory=list)
    meta: dict[str, Any] = field(default_factory=dict)


Progress = Callable[[int], None]


def _check_cancel(cancel_event: threading.Event | None) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise ConversionError("CANCELED", "任务已取消")


def _style_for(options: dict[str, Any]) -> dict[str, Any] | None:
    style = options.get("style_data")
    return style if isinstance(style, dict) else None


def _apply_style(path: Path, fmt: str, options: dict[str, Any], password: str | None = None) -> list[str]:
    style = _style_for(options)
    if not style:
        return []
    try:
        return apply_style(path, fmt, style, password)
    except Exception as exc:  # noqa: BLE001
        return [f"样式模板应用失败：{exc}"]


def _to_docx_intermediate(
    source: Path,
    source_fmt: str,
    mode: str,
    temp_dir: Path,
    options: dict[str, Any],
) -> Path:
    target = temp_dir / f"{safe_stem(source.name)}.docx"
    if source_fmt == "docx":
        shutil.copy2(source, target)
        return target
    if source_fmt == "txt":
        txt_to_docx(source, target)
        return target
    if source_fmt == "md":
        md_to_docx(source, target)
        return target
    if source_fmt == "pdf":
        password = options.get("pdf_password")
        if mode == "layout":
            pdf_to_docx_layout(source, target, password)
        else:
            pdf_to_docx_editable(source, target, password)
        return target
    if source_fmt == "wps":
        wps_convert(source, target, "docx")
        return target
    raise ConversionError("UNSUPPORTED_PAIR", f"无法把 {source_fmt} 转为 WPS 所需的 DOCX 中间格式")


def _dispatch(
    source: Path,
    source_fmt: str,
    target_fmt: str,
    mode: str,
    destination: Path,
    temp_dir: Path,
    options: dict[str, Any],
    progress: Progress,
    cancel_event: threading.Event | None,
) -> ConversionResult:
    result = ConversionResult()
    password = options.get("pdf_password")
    progress(10)
    _check_cancel(cancel_event)

    if source_fmt == "wps":
        if target_fmt == "docx":
            wps_convert(source, destination, "docx")
            result.engine = "wps"
            result.warnings.extend(_apply_style(destination, "docx", options))
        elif target_fmt == "pdf":
            wps_convert(source, destination, "pdf")
            result.engine = "wps"
            result.warnings.extend(_apply_style(destination, "pdf", options, password))
        elif target_fmt == "txt":
            intermediate = temp_dir / f"{safe_stem(source.name)}.docx"
            wps_convert(source, intermediate, "docx")
            docx_to_txt(intermediate, destination)
            result.engine = "wps"
        elif target_fmt == "md":
            intermediate = temp_dir / f"{safe_stem(source.name)}.docx"
            wps_convert(source, intermediate, "docx")
            docx_to_md(intermediate, destination)
            result.engine = "wps"
        else:
            raise ConversionError("UNSUPPORTED_PAIR", "WPS 源文件仅支持转为 docx/pdf/txt/md")
        result.outputs = [destination]
        return result

    if target_fmt == "wps":
        intermediate = _to_docx_intermediate(source, source_fmt, mode, temp_dir, options)
        result.warnings.extend(_apply_style(intermediate, "docx", options))
        to_wps(intermediate, destination)
        result.engine = "wps"
        result.outputs = [destination]
        return result

    if source_fmt == "txt":
        if target_fmt == "md":
            txt_to_md(source, destination)
        elif target_fmt == "docx":
            txt_to_docx(source, destination)
            result.warnings.extend(_apply_style(destination, "docx", options))
        elif target_fmt == "pdf":
            intermediate = temp_dir / f"{safe_stem(source.name)}.docx"
            txt_to_docx(source, intermediate)
            result.engine = office_to_pdf(intermediate, destination)
            result.warnings.extend(_apply_style(destination, "pdf", options, password))
        else:
            raise ConversionError("UNSUPPORTED_PAIR", f"不支持 txt → {target_fmt}")
        result.outputs = [destination]

    elif source_fmt == "md":
        if target_fmt == "txt":
            md_to_txt(source, destination)
        elif target_fmt == "docx":
            md_to_docx(source, destination)
            result.warnings.extend(_apply_style(destination, "docx", options))
        elif target_fmt == "pdf":
            intermediate = temp_dir / f"{safe_stem(source.name)}.docx"
            md_to_docx(source, intermediate)
            result.engine = office_to_pdf(intermediate, destination)
            result.warnings.extend(_apply_style(destination, "pdf", options, password))
        else:
            raise ConversionError("UNSUPPORTED_PAIR", f"不支持 md → {target_fmt}")
        result.outputs = [destination]

    elif source_fmt == "docx":
        if target_fmt == "txt":
            docx_to_txt(source, destination)
        elif target_fmt == "md":
            docx_to_md(source, destination)
            result.outputs = [destination]
            assets = destination.parent / "assets"
            if assets.exists():
                result.outputs.append(assets)
            return result
        elif target_fmt == "xlsx":
            try:
                count = docx_to_xlsx(source, destination)
                result.meta["tables"] = count
            except ValueError as exc:
                raise ConversionError("NO_TABLE", str(exc)) from exc
        elif target_fmt == "pdf":
            result.engine = office_to_pdf(source, destination)
            result.warnings.extend(_apply_style(destination, "pdf", options, password))
        else:
            raise ConversionError("UNSUPPORTED_PAIR", f"不支持 docx → {target_fmt}")
        result.outputs = [destination]

    elif source_fmt == "pdf":
        password = options.get("pdf_password")
        if target_fmt == "txt":
            pdf_to_txt(source, destination, password)
        elif target_fmt == "md":
            pdf_to_md(source, destination, password)
        elif target_fmt == "docx":
            if mode == "layout":
                result.warnings.extend(pdf_to_docx_layout(source, destination, password))
            else:
                result.warnings.extend(pdf_to_docx_editable(source, destination, password))
            result.warnings.extend(_apply_style(destination, "docx", options))
        elif target_fmt == "pptx":
            if mode == "layout":
                result.warnings.extend(pdf_to_pptx_layout(source, destination, password))
            else:
                result.warnings.extend(pdf_to_pptx_editable(source, destination, password))
        elif target_fmt == "xlsx":
            try:
                count = pdf_to_xlsx(source, destination, password)
                result.meta["tables"] = count
            except PdfConversionError as exc:
                raise ConversionError(exc.code, exc.message) from exc
        else:
            raise ConversionError("UNSUPPORTED_PAIR", f"不支持 pdf → {target_fmt}")
        result.outputs = [destination]

    elif source_fmt == "csv":
        if target_fmt == "xlsx":
            csv_to_xlsx(source, destination, bool(options.get("csv_all_text")))
        elif target_fmt == "docx":
            csv_to_docx(source, destination)
            result.warnings.extend(_apply_style(destination, "docx", options))
        elif target_fmt == "md":
            csv_to_md(source, destination)
        elif target_fmt == "pdf":
            intermediate = temp_dir / f"{safe_stem(source.name)}.xlsx"
            csv_to_xlsx(source, intermediate, bool(options.get("csv_all_text")))
            result.engine = office_to_pdf(intermediate, destination)
            result.warnings.extend(_apply_style(destination, "pdf", options, password))
        else:
            raise ConversionError("UNSUPPORTED_PAIR", f"不支持 csv → {target_fmt}")
        result.outputs = [destination]

    elif source_fmt == "xlsx":
        sheet = options.get("sheet_name") or None
        if target_fmt == "csv":
            xlsx_to_csv(source, destination, sheet)
        elif target_fmt == "docx":
            xlsx_to_docx(source, destination, sheet)
            result.warnings.extend(_apply_style(destination, "docx", options))
        elif target_fmt == "md":
            xlsx_to_md(source, destination, sheet)
        elif target_fmt == "pdf":
            result.engine = office_to_pdf(source, destination)
            result.warnings.extend(_apply_style(destination, "pdf", options, password))
        else:
            raise ConversionError("UNSUPPORTED_PAIR", f"不支持 xlsx → {target_fmt}")
        result.outputs = [destination]

    elif source_fmt == "pptx":
        if target_fmt == "pdf":
            result.engine = office_to_pdf(source, destination)
            result.warnings.extend(_apply_style(destination, "pdf", options, password))
        else:
            raise ConversionError("UNSUPPORTED_PAIR", f"不支持 pptx → {target_fmt}")
        result.outputs = [destination]

    else:
        raise ConversionError("UNSUPPORTED_PAIR", f"不支持的源格式：{source_fmt}")

    return result


def convert_file(
    source: Path,
    source_fmt: str,
    target_fmt: str,
    mode: str,
    output_dir: Path,
    options: dict[str, Any] | None = None,
    progress: Progress | None = None,
    cancel_event: threading.Event | None = None,
) -> ConversionResult:
    options = dict(options or {})
    capability = capability_for(source_fmt, target_fmt)
    if not capability:
        raise ConversionError("UNSUPPORTED_PAIR", f"v1 不支持 {source_fmt} → {target_fmt}")
    if mode not in capability["modes"]:
        mode = capability["modes"][0]
    progress = progress or (lambda _value: None)
    output_dir.mkdir(parents=True, exist_ok=True)
    destination = unique_path(output_dir / f"{safe_stem(source.name)}{extension_for(target_fmt)}")
    with tempfile.TemporaryDirectory(prefix="format-studio-") as temp:
        temp_dir = Path(temp)
        try:
            result = _dispatch(
                source,
                source_fmt,
                target_fmt,
                mode,
                destination,
                temp_dir,
                options,
                progress,
                cancel_event,
            )
        except ConversionError:
            raise
        except EngineError as exc:
            raise ConversionError(exc.code, exc.message, exc.retryable) from exc
        except PdfConversionError as exc:
            raise ConversionError(exc.code, exc.message) from exc
        except PermissionError as exc:
            raise ConversionError("OUTPUT_WRITE_FAILED", f"无法写入输出文件：{exc}") from exc
        except Exception as exc:  # noqa: BLE001
            raise ConversionError("CONVERSION_FAILED", f"转换失败：{exc}") from exc
    progress(96)
    _check_cancel(cancel_event)
    if not result.outputs:
        raise ConversionError("CONVERSION_FAILED", "转换未生成输出文件")
    if not result.outputs[0].exists() or result.outputs[0].stat().st_size == 0:
        raise ConversionError("CONVERSION_FAILED", "输出文件为空或不存在")
    result.meta.update(inspect_document(result.outputs[0], target_fmt))
    progress(100)
    return result