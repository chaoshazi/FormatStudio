from __future__ import annotations

import io
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas

_FONT_READY = False


def _ensure_font() -> None:
    global _FONT_READY
    if not _FONT_READY:
        try:
            pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
        except Exception:
            pass
        _FONT_READY = True


def open_pdf(path: Path, password: str | None = None):
    import pypdfium2 as pdfium

    kwargs = {}
    if password:
        kwargs["password"] = password
    return pdfium.PdfDocument(str(path), **kwargs)


def render_pages(
    path: Path,
    out_dir: Path,
    dpi: int = 200,
    password: str | None = None,
    max_pages: int | None = None,
) -> list[dict]:
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf = open_pdf(path, password)
    results: list[dict] = []
    try:
        count = len(pdf)
        if max_pages is not None:
            count = min(count, max_pages)
        scale = dpi / 72.0
        for index in range(count):
            page = pdf[index]
            width, height = page.get_size()
            bitmap = page.render(scale=scale)
            image = bitmap.to_pil().convert("RGB")
            target = out_dir / f"page-{index + 1:03d}.png"
            image.save(target, "PNG", optimize=True)
            results.append(
                {
                    "index": index,
                    "path": target,
                    "width": width,
                    "height": height,
                    "width_px": image.width,
                    "height_px": image.height,
                }
            )
            image.close()
            bitmap.close()
            page.close()
    finally:
        pdf.close()
    return results


def render_page_image(
    path: Path,
    out_path: Path,
    page_index: int = 0,
    dpi: int = 200,
    password: str | None = None,
) -> dict:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    pdf = open_pdf(path, password)
    try:
        page = pdf[page_index]
        width, height = page.get_size()
        bitmap = page.render(scale=dpi / 72.0)
        image = bitmap.to_pil().convert("RGB")
        image.save(out_path, "PNG", optimize=True)
        result = {
            "path": out_path,
            "width": width,
            "height": height,
            "width_px": image.width,
            "height_px": image.height,
        }
        image.close()
        bitmap.close()
        page.close()
        return result
    finally:
        pdf.close()


def add_watermark(
    source: Path,
    destination: Path,
    text: str,
    opacity: float = 0.15,
    angle: int = 45,
    size: int = 48,
    color: str = "94A3B8",
    tile: bool = False,
    password: str | None = None,
) -> None:
    _ensure_font()
    reader = PdfReader(str(source))
    if reader.is_encrypted and password:
        reader.decrypt(password)
    writer = PdfWriter()
    overlay_buffer = io.BytesIO()
    first = reader.pages[0]
    width = float(first.mediabox.width)
    height = float(first.mediabox.height)
    pdf_canvas = canvas.Canvas(overlay_buffer, pagesize=(width, height))
    try:
        rgb = HexColor(f"#{color.lstrip('#')}")
    except Exception:
        rgb = HexColor("#94A3B8")
    for page in reader.pages:
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)
        pdf_canvas.setPageSize((page_width, page_height))
        pdf_canvas.saveState()
        try:
            pdf_canvas.setFillColor(rgb)
            pdf_canvas.setFillAlpha(max(0.0, min(1.0, opacity)))
        except Exception:
            pass
        pdf_canvas.setFont("STSong-Light", size)
        if tile:
            step_x = max(width * 0.45, 220)
            step_y = max(height * 0.28, 160)
            for x in range(0, int(page_width + step_x), int(step_x)):
                for y in range(0, int(page_height + step_y), int(step_y)):
                    pdf_canvas.saveState()
                    pdf_canvas.translate(x, y)
                    pdf_canvas.rotate(angle)
                    pdf_canvas.drawCentredString(0, 0, text)
                    pdf_canvas.restoreState()
        else:
            pdf_canvas.saveState()
            pdf_canvas.translate(page_width / 2, page_height / 2)
            pdf_canvas.rotate(angle)
            pdf_canvas.drawCentredString(0, 0, text)
            pdf_canvas.restoreState()
        pdf_canvas.restoreState()
        pdf_canvas.showPage()
    pdf_canvas.save()
    overlay_buffer.seek(0)
    overlay = PdfReader(overlay_buffer)
    for index, page in enumerate(reader.pages):
        if index < len(overlay.pages):
            page.merge_page(overlay.pages[index])
        writer.add_page(page)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("wb") as handle:
        writer.write(handle)