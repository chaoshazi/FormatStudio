from __future__ import annotations

from copy import deepcopy

DOCX_BASE = {
    "font_cn": "微软雅黑",
    "font_en": "Calibri",
    "title_size": 22,
    "h1_size": 18,
    "h2_size": 15,
    "body_size": 11,
    "line_spacing": 1.5,
    "space_before": 0,
    "space_after": 6,
    "first_line_indent": 0,
    "margin_top_cm": 2.54,
    "margin_bottom_cm": 2.54,
    "margin_left_cm": 3.18,
    "margin_right_cm": 3.18,
    "orientation": "portrait",
    "header_text": "",
    "footer_text": "",
    "page_number": True,
    "force_direct_format": False,
    "primary_color": "2563EB",
    "accent_color": "60A5FA",
    "light_color": "EFF6FF",
    "heading_color": "1E3A8A",
    "text_color": "1F2937",
    "header_fill": "2563EB",
    "page_background": "FFFFFF",
}

XLSX_BASE = {
    "font_name": "微软雅黑",
    "font_size": 11,
    "header_bold": True,
    "header_fill": "DBEAFE",
    "header_color": "1E3A8A",
    "border_color": "CBD5E1",
    "zebra_fill": "F8FAFC",
    "freeze_header": True,
    "auto_width": True,
    "landscape": True,
    "fit_to_width": True,
    "tab_color": "2563EB",
}

PPTX_BASE = {
    "title_font": "微软雅黑",
    "body_font": "微软雅黑",
    "theme_colors": ["2563EB", "0F172A", "64748B", "DBEAFE", "F8FAFC", "16A34A"],
    "slide_width_inches": 13.333,
    "slide_height_inches": 7.5,
    "background_color": "F8FAFC",
    "primary_color": "2563EB",
    "accent_color": "60A5FA",
    "text_color": "0F172A",
}

PDF_BASE = {
    "watermark": {
        "enabled": False,
        "text": "内部资料",
        "size": 48,
        "color": "94A3B8",
        "opacity": 0.15,
        "angle": 45,
        "tile": False,
    }
}


def _base(name: str, applies_to: list[str]) -> dict:
    return {
        "version": 1,
        "name": name,
        "applies_to": applies_to,
        "docx": deepcopy(DOCX_BASE),
        "xlsx": deepcopy(XLSX_BASE),
        "pptx": deepcopy(PPTX_BASE),
        "pdf": deepcopy(PDF_BASE),
    }


def _base_builtin_presets() -> list[dict]:
    report = _base("简洁报告", ["docx"])
    report["docx"].update(
        {
            "font_cn": "微软雅黑",
            "title_size": 24,
            "h1_size": 18,
            "h2_size": 15,
            "body_size": 11,
            "line_spacing": 1.6,
            "space_after": 8,
            "first_line_indent": 0,
            "header_text": "",
            "footer_text": "机密 · 内部使用",
            "page_number": True,
        }
    )

    business = _base("商务蓝", ["docx", "pptx"])
    business["docx"].update(
        {
            "font_cn": "微软雅黑",
            "title_size": 26,
            "h1_size": 20,
            "h2_size": 16,
            "body_size": 11,
            "line_spacing": 1.5,
            "space_after": 10,
            "header_text": "BUSINESS REPORT",
            "footer_text": "",
            "page_number": True,
        }
    )
    business["pptx"].update(
        {
            "title_font": "微软雅黑",
            "body_font": "微软雅黑",
            "theme_colors": ["1D4ED8", "0F172A", "475569", "DBEAFE", "EFF6FF", "0EA5E9"],
        }
    )

    official = _base("公文规范", ["docx"])
    official["docx"].update(
        {
            "font_cn": "仿宋",
            "font_en": "Times New Roman",
            "title_size": 22,
            "h1_size": 16,
            "h2_size": 16,
            "body_size": 16,
            "line_spacing": 1.5,
            "space_before": 0,
            "space_after": 0,
            "first_line_indent": 2,
            "margin_top_cm": 3.7,
            "margin_bottom_cm": 3.5,
            "margin_left_cm": 2.8,
            "margin_right_cm": 2.6,
            "header_text": "",
            "footer_text": "",
            "page_number": True,
        }
    )

    grid = _base("清晰网格", ["xlsx"])
    grid["xlsx"].update(
        {
            "header_fill": "1E3A8A",
            "header_color": "FFFFFF",
            "border_color": "94A3B8",
            "zebra_fill": "F1F5F9",
            "freeze_header": True,
            "landscape": True,
        }
    )

    zebra = _base("斑马纹", ["xlsx"])
    zebra["xlsx"].update(
        {
            "header_fill": "0F766E",
            "header_color": "FFFFFF",
            "border_color": "CCFBF1",
            "zebra_fill": "F0FDFA",
            "freeze_header": True,
            "landscape": False,
        }
    )

    minimal = _base("极简灰", ["pptx"])
    minimal["pptx"].update(
        {
            "title_font": "微软雅黑",
            "body_font": "微软雅黑",
            "theme_colors": ["374151", "111827", "6B7280", "E5E7EB", "F9FAFB", "9CA3AF"],
        }
    )

    report["docx"].update({"primary_color": "2563EB", "accent_color": "60A5FA", "light_color": "EFF6FF", "heading_color": "1E3A8A", "text_color": "1F2937", "header_fill": "2563EB", "page_background": "F8FAFC"})
    business["docx"].update({"primary_color": "1D4ED8", "accent_color": "3B82F6", "light_color": "EFF6FF", "heading_color": "1E3A8A", "text_color": "1F2937", "header_fill": "1D4ED8", "page_background": "EFF6FF"})
    business["pptx"].update({"background_color": "F8FAFC", "primary_color": "1D4ED8", "accent_color": "3B82F6", "text_color": "0F172A"})
    official["docx"].update({"primary_color": "991B1B", "accent_color": "B91C1C", "light_color": "FEF2F2", "heading_color": "7F1D1D", "text_color": "111827", "header_fill": "991B1B", "page_background": "FEF2F2"})
    grid["xlsx"].update({"tab_color": "1E3A8A"})
    zebra["xlsx"].update({"tab_color": "0F766E"})
    minimal["pptx"].update({"background_color": "F8FAFC", "primary_color": "374151", "accent_color": "9CA3AF", "text_color": "111827"})
    return [
        {"id": "builtin-report", "name": "简洁报告", "data": report},
        {"id": "builtin-business", "name": "商务蓝", "data": business},
        {"id": "builtin-official", "name": "公文规范", "data": official},
        {"id": "builtin-grid", "name": "清晰网格", "data": grid},
        {"id": "builtin-zebra", "name": "斑马纹", "data": zebra},
        {"id": "builtin-minimal", "name": "极简灰", "data": minimal},
    ]

def _docx_style(name: str, font_cn: str, font_en: str, body: float, title: float, h1: float, h2: float, line: float, top: float, left: float, header: str, footer: str, *, indent: float = 0) -> dict:
    data = _base(name, ["docx"])
    data["docx"].update(
        {
            "font_cn": font_cn,
            "font_en": font_en,
            "body_size": body,
            "title_size": title,
            "h1_size": h1,
            "h2_size": h2,
            "line_spacing": line,
            "margin_top_cm": top,
            "margin_bottom_cm": top,
            "margin_left_cm": left,
            "margin_right_cm": left,
            "header_text": header,
            "footer_text": footer,
            "first_line_indent": indent,
            "page_number": True,
        }
    )
    return data


def _xlsx_style(name: str, fill: str, header_color: str, border: str, zebra: str, font: str = "微软雅黑", *, font_color: str = "0F172A", body_fill: str = "", landscape: bool = True) -> dict:
    data = _base(name, ["xlsx"])
    data["xlsx"].update(
        {
            "font_name": font,
            "font_color": font_color,
            "body_fill": body_fill,
            "header_fill": fill,
            "header_color": header_color,
            "border_color": border,
            "zebra_fill": zebra,
            "header_bold": True,
            "freeze_header": True,
            "auto_width": True,
            "fit_to_width": True,
            "landscape": landscape,
        }
    )
    return data


def _pptx_style(name: str, title_font: str, body_font: str, colors: list[str]) -> dict:
    data = _base(name, ["pptx"])
    data["pptx"].update(
        {
            "title_font": title_font,
            "body_font": body_font,
            "theme_colors": colors,
            "slide_width_inches": 13.333,
            "slide_height_inches": 7.5,
        }
    )
    return data


def _extra_style_presets() -> list[dict]:
    n = {
        "academic": "\u5b66\u672f\u8bba\u6587",
        "ink_green": "\u58a8\u7eff\u5546\u52a1",
        "warm_orange": "\u6696\u6a59\u63d0\u6848",
        "minimal_gray": "\u6781\u7b80\u7070",
        "tech_purple": "\u79d1\u6280\u7d2b",
        "finance_navy": "\u91d1\u878d\u6df1\u84dd",
        "creative_red": "\u521b\u610f\u7ea2",
        "business_blue": "\u5546\u52a1\u84dd\u8868",
        "finance_green": "\u8d22\u52a1\u7eff",
        "sales_orange": "\u9500\u552e\u6a59",
        "project_board": "\u9879\u76ee\u770b\u677f",
        "minimal_line": "\u6781\u7b80\u7ebf\u6761",
        "dark_dashboard": "\u6df1\u8272\u4eea\u8868",
        "soft_purple": "\u67d4\u548c\u7d2b",
        "high_contrast": "\u9ad8\u5bf9\u6bd4",
        "vibrant_orange": "\u6d3b\u529b\u6a59",
        "ink_green_ppt": "\u58a8\u7eff",
        "fresh_cyan": "\u6e05\u65b0\u9752",
        "warm_sand": "\u6696\u6c99",
        "black_white": "\u9ed1\u767d",
    }
    cn = "\u5fae\u8f6f\u96c5\u9ed1"
    song = "\u5b8b\u4f53"
    internal = "\u5185\u90e8\u8d44\u6599"
    items = [
        {"id": "builtin-docx-academic", "name": n["academic"], "data": _docx_style(n["academic"], song, "Times New Roman", 12, 22, 16, 14, 1.5, 2.54, 3.0, "", "", indent=2)},
        {"id": "builtin-docx-ink-green", "name": n["ink_green"], "data": _docx_style(n["ink_green"], cn, "Calibri", 11, 24, 18, 15, 1.65, 2.2, 2.8, "GREEN REPORT", internal)},
        {"id": "builtin-docx-warm-orange", "name": n["warm_orange"], "data": _docx_style(n["warm_orange"], cn, "Calibri", 11, 26, 19, 16, 1.6, 2.3, 2.9, "PROPOSAL", "Confidential")},
        {"id": "builtin-docx-minimal-gray", "name": n["minimal_gray"], "data": _docx_style(n["minimal_gray"], cn, "Calibri", 10.5, 22, 16, 13, 1.45, 2.0, 2.2, "", "")},
        {"id": "builtin-docx-tech-purple", "name": n["tech_purple"], "data": _docx_style(n["tech_purple"], cn, "Calibri", 11, 26, 18, 15, 1.55, 2.2, 2.7, "TECH REPORT", "Innovation")},
        {"id": "builtin-docx-finance-navy", "name": n["finance_navy"], "data": _docx_style(n["finance_navy"], song, "Times New Roman", 12, 24, 17, 14, 1.6, 2.5, 3.0, "FINANCIAL REPORT", internal)},
        {"id": "builtin-docx-creative-red", "name": n["creative_red"], "data": _docx_style(n["creative_red"], cn, "Calibri", 11, 28, 19, 16, 1.5, 2.1, 2.6, "CREATIVE BRIEF", "")},
        {"id": "builtin-xlsx-business-blue", "name": n["business_blue"], "data": _xlsx_style(n["business_blue"], "1D4ED8", "FFFFFF", "BFDBFE", "EFF6FF")},
        {"id": "builtin-xlsx-finance-green", "name": n["finance_green"], "data": _xlsx_style(n["finance_green"], "047857", "FFFFFF", "A7F3D0", "ECFDF5")},
        {"id": "builtin-xlsx-sales-orange", "name": n["sales_orange"], "data": _xlsx_style(n["sales_orange"], "C2410C", "FFFFFF", "FED7AA", "FFF7ED")},
        {"id": "builtin-xlsx-project-board", "name": n["project_board"], "data": _xlsx_style(n["project_board"], "4338CA", "FFFFFF", "C7D2FE", "EEF2FF")},
        {"id": "builtin-xlsx-minimal-line", "name": n["minimal_line"], "data": _xlsx_style(n["minimal_line"], "334155", "FFFFFF", "CBD5E1", "F8FAFC", landscape=False)},
        {"id": "builtin-xlsx-dark-dashboard", "name": n["dark_dashboard"], "data": _xlsx_style(n["dark_dashboard"], "0F172A", "F8FAFC", "334155", "1E293B", font_color="F8FAFC", body_fill="0F172A")},
        {"id": "builtin-xlsx-soft-purple", "name": n["soft_purple"], "data": _xlsx_style(n["soft_purple"], "6D28D9", "FFFFFF", "DDD6FE", "F5F3FF")},
        {"id": "builtin-xlsx-high-contrast", "name": n["high_contrast"], "data": _xlsx_style(n["high_contrast"], "111827", "FFFFFF", "F59E0B", "FEF3C7")},
        {"id": "builtin-pptx-tech-purple", "name": n["tech_purple"], "data": _pptx_style(n["tech_purple"], cn, cn, ["6D28D9", "2E1065", "8B5CF6", "DDD6FE", "F5F3FF", "06B6D4"])},
        {"id": "builtin-pptx-vibrant-orange", "name": n["vibrant_orange"], "data": _pptx_style(n["vibrant_orange"], cn, cn, ["C2410C", "431407", "F97316", "FED7AA", "FFF7ED", "0EA5E9"])},
        {"id": "builtin-pptx-ink-green", "name": n["ink_green_ppt"], "data": _pptx_style(n["ink_green_ppt"], cn, cn, ["047857", "064E3B", "10B981", "A7F3D0", "ECFDF5", "0F766E"])},
        {"id": "builtin-pptx-finance-navy", "name": n["finance_navy"], "data": _pptx_style(n["finance_navy"], cn, cn, ["1E3A8A", "0F172A", "3B82F6", "BFDBFE", "EFF6FF", "0F766E"])},
        {"id": "builtin-pptx-creative-red", "name": n["creative_red"], "data": _pptx_style(n["creative_red"], cn, cn, ["BE123C", "4C0519", "F43F5E", "FECDD3", "FFF1F2", "F59E0B"])},
        {"id": "builtin-pptx-fresh-cyan", "name": n["fresh_cyan"], "data": _pptx_style(n["fresh_cyan"], cn, cn, ["0E7490", "164E63", "06B6D4", "A5F3FC", "ECFEFF", "14B8A6"])},
        {"id": "builtin-pptx-warm-sand", "name": n["warm_sand"], "data": _pptx_style(n["warm_sand"], cn, cn, ["92400E", "451A03", "D97706", "FDE68A", "FFFBEB", "B45309"])},
        {"id": "builtin-pptx-black-white", "name": n["black_white"], "data": _pptx_style(n["black_white"], cn, cn, ["111827", "000000", "4B5563", "D1D5DB", "F9FAFB", "6B7280"])},
    ]


    docx_themes = {
        "builtin-docx-academic": ("374151", "9CA3AF", "F3F4F6", "111827", "1F2937"),
        "builtin-docx-ink-green": ("047857", "10B981", "ECFDF5", "064E3B", "064E3B"),
        "builtin-docx-warm-orange": ("C2410C", "F97316", "FFF7ED", "431407", "431407"),
        "builtin-docx-minimal-gray": ("334155", "94A3B8", "F8FAFC", "0F172A", "1F2937"),
        "builtin-docx-tech-purple": ("6D28D9", "8B5CF6", "F5F3FF", "2E1065", "2E1065"),
        "builtin-docx-finance-navy": ("1E3A8A", "3B82F6", "EFF6FF", "0F172A", "0F172A"),
        "builtin-docx-creative-red": ("BE123C", "F43F5E", "FFF1F2", "4C0519", "4C0519"),
    }
    for item in items:
        data = item["data"]
        if item["id"] in docx_themes:
            primary, accent, light, heading, text = docx_themes[item["id"]]
            data["docx"].update(
                {
                    "primary_color": primary,
                    "accent_color": accent,
                    "light_color": light,
                    "heading_color": heading,
                    "text_color": text,
                    "header_fill": primary,
                    "page_background": light,
                }
            )
        if data["applies_to"] == ["xlsx"] and item["id"] != "builtin-xlsx-dark-dashboard":
            data["xlsx"]["tab_color"] = data["xlsx"]["header_fill"]
            data["xlsx"]["body_fill"] = "F8FAFC"
        if data["applies_to"] == ["pptx"]:
            colors = data["pptx"]["theme_colors"]
            data["pptx"].update(
                {
                    "background_color": colors[4],
                    "primary_color": colors[0],
                    "accent_color": colors[2],
                    "text_color": colors[1],
                }
            )
    return items

def builtin_presets() -> list[dict]:
    return _base_builtin_presets() + _extra_style_presets()