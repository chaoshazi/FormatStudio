from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from .. import db
from ..config import TEMPLATES_DIR
from .template_service import scan_template

LOGGER = logging.getLogger("format-studio.builtin-templates")
BUILTIN_VERSION = 8

PALETTES: dict[str, dict[str, str]] = {
    "blue": {"primary": "1D4ED8", "dark": "0F172A", "accent": "3B82F6", "light": "EFF6FF", "muted": "64748B", "white": "FFFFFF"},
    "teal": {"primary": "0F766E", "dark": "134E4A", "accent": "14B8A6", "light": "F0FDFA", "muted": "64748B", "white": "FFFFFF"},
    "emerald": {"primary": "047857", "dark": "064E3B", "accent": "10B981", "light": "ECFDF5", "muted": "64748B", "white": "FFFFFF"},
    "orange": {"primary": "C2410C", "dark": "431407", "accent": "F97316", "light": "FFF7ED", "muted": "78716C", "white": "FFFFFF"},
    "purple": {"primary": "6D28D9", "dark": "2E1065", "accent": "8B5CF6", "light": "F5F3FF", "muted": "64748B", "white": "FFFFFF"},
    "rose": {"primary": "BE123C", "dark": "4C0519", "accent": "F43F5E", "light": "FFF1F2", "muted": "64748B", "white": "FFFFFF"},
    "slate": {"primary": "334155", "dark": "0F172A", "accent": "64748B", "light": "F8FAFC", "muted": "64748B", "white": "FFFFFF"},
    "indigo": {"primary": "4338CA", "dark": "1E1B4B", "accent": "6366F1", "light": "EEF2FF", "muted": "64748B", "white": "FFFFFF"},
    "amber": {"primary": "B45309", "dark": "451A03", "accent": "F59E0B", "light": "FFFBEB", "muted": "78716C", "white": "FFFFFF"},
    "cyan": {"primary": "0E7490", "dark": "164E63", "accent": "06B6D4", "light": "ECFEFF", "muted": "64748B", "white": "FFFFFF"},
}
for _palette in PALETTES.values():
    _palette.setdefault("border", "E2E8F0")


def _rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def _shade_cell(cell, color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), color)
    shd.set(qn("w:val"), "clear")


def _shade_paragraph(paragraph, color: str) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    shd = p_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        p_pr.append(shd)
    shd.set(qn("w:fill"), color)
    shd.set(qn("w:val"), "clear")


def _bottom_border(paragraph, color: str, size: int = 8) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = p_pr.find(qn("w:pBdr"))
    if p_bdr is None:
        p_bdr = OxmlElement("w:pBdr")
        p_pr.append(p_bdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)


def _set_run(run, size: float, color: str, bold: bool = False, font: str = "微软雅黑") -> None:
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = _rgb(color)
    r_pr = run._element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = r_pr.makeelement(qn("w:rFonts"), {})
        r_pr.append(r_fonts)
    r_fonts.set(qn("w:eastAsia"), font)
    r_fonts.set(qn("w:ascii"), font)
    r_fonts.set(qn("w:hAnsi"), font)


def _title_band(document, title: str, subtitle: str, palette: dict[str, str]) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    _shade_paragraph(paragraph, palette["primary"])
    run = paragraph.add_run(f"  {title}")
    _set_run(run, 24, palette["white"], True)
    sub = document.add_paragraph()
    sub.paragraph_format.space_before = Pt(6)
    sub.paragraph_format.space_after = Pt(10)
    run = sub.add_run(subtitle)
    _set_run(run, 11, palette["muted"])


def _meta_table(document, pairs: list[tuple[str, str]], palette: dict[str, str]) -> None:
    rows = (len(pairs) + 1) // 2
    table = document.add_table(rows=rows, cols=4)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for index, (label, value) in enumerate(pairs):
        row = index // 2
        col = (index % 2) * 2
        label_cell = table.cell(row, col)
        value_cell = table.cell(row, col + 1)
        label_cell.text = ""
        value_cell.text = ""
        label_paragraph = label_cell.paragraphs[0]
        value_paragraph = value_cell.paragraphs[0]
        run = label_paragraph.add_run(label)
        _set_run(run, 9.5, palette["muted"], True)
        run = value_paragraph.add_run(value)
        _set_run(run, 10, palette["dark"])
        _shade_cell(label_cell, palette["light"])
    document.add_paragraph().paragraph_format.space_after = Pt(2)


def _section_heading(document, text: str, palette: dict[str, str]) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(12)
    paragraph.paragraph_format.space_after = Pt(4)
    run = paragraph.add_run(text)
    _set_run(run, 13, palette["primary"], True)
    _bottom_border(paragraph, palette["accent"], 6)


def _text(document, text: str, palette: dict[str, str]) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.line_spacing = 1.6
    paragraph.paragraph_format.space_after = Pt(4)
    run = paragraph.add_run(text)
    _set_run(run, 10.5, palette["dark"])


def _bullets(document, items: list[str], palette: dict[str, str]) -> None:
    for item in items:
        paragraph = document.add_paragraph(style="List Bullet")
        paragraph.paragraph_format.line_spacing = 1.5
        run = paragraph.add_run(item)
        _set_run(run, 10.5, palette["dark"])


def _table(document, headers: list[str], rows: list[list[str]], palette: dict[str, str]) -> None:
    table = document.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for index, header in enumerate(headers):
        cell = table.cell(0, index)
        cell.text = ""
        run = cell.paragraphs[0].add_run(header)
        _set_run(run, 9.5, palette["white"], True)
        _shade_cell(cell, palette["primary"])
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for row_index, row in enumerate(rows, start=1):
        for col_index, value in enumerate(row):
            cell = table.cell(row_index, col_index)
            cell.text = ""
            run = cell.paragraphs[0].add_run(value)
            _set_run(run, 9.3, palette["dark"])
            if row_index % 2 == 0:
                _shade_cell(cell, palette["light"])
    document.add_paragraph().paragraph_format.space_after = Pt(2)


def _footer(document, text: str, palette: dict[str, str]) -> None:
    for section in document.sections:
        footer = section.footer
        paragraph = footer.paragraphs[0]
        paragraph.text = ""
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(text)
        _set_run(run, 8.5, palette["muted"])
DOCX_SPECS: list[dict[str, Any]] = [
    {
        "id": "builtin-docx-business-report",
        "name": "商务报告 · 深海蓝",
        "format": "docx",
        "palette": "blue",
        "title": "{{title}}",
        "subtitle": "{{company}} · {{department}}",
        "meta": [
            ("报告人", "{{author}}"),
            ("日期", "{{date}}"),
            ("版本", "{{version}}"),
            ("密级", "{{confidentiality}}"),
        ],
        "sections": [
            {"heading": "执行摘要", "text": "{{summary}}"},
            {
                "heading": "核心指标",
                "table": {
                    "headers": ["指标", "本期", "上期", "变化"],
                    "rows": [
                        ["{{kpi1_name}}", "{{kpi1_value}}", "{{kpi1_previous}}", "{{kpi1_delta}}"],
                        ["{{kpi2_name}}", "{{kpi2_value}}", "{{kpi2_previous}}", "{{kpi2_delta}}"],
                        ["{{kpi3_name}}", "{{kpi3_value}}", "{{kpi3_previous}}", "{{kpi3_delta}}"],
                    ],
                },
            },
            {"heading": "业务进展", "bullets": ["{{progress1}}", "{{progress2}}", "{{progress3}}"]},
            {"heading": "问题与风险", "text": "{{risk}}"},
            {"heading": "下一步计划", "text": "{{next_plan}}"},
            {"heading": "结论", "text": "{{conclusion}}"},
        ],
        "footer": "{{company}} · 内部资料",
    },
    {
        "id": "builtin-docx-weekly-report",
        "name": "项目周报 · 青绿",
        "format": "docx",
        "palette": "teal",
        "title": "项目周报 · {{project}}",
        "subtitle": "{{week}}  |  负责人：{{owner}}",
        "meta": [
            ("项目", "{{project}}"),
            ("周期", "{{week}}"),
            ("负责人", "{{owner}}"),
            ("日期", "{{date}}"),
        ],
        "sections": [
            {"heading": "本周概述", "text": "{{summary}}"},
            {
                "heading": "本周进展",
                "table": {
                    "headers": ["任务", "负责人", "状态", "完成度"],
                    "rows": [
                        ["{{task1}}", "{{task1_owner}}", "{{task1_status}}", "{{task1_progress}}"],
                        ["{{task2}}", "{{task2_owner}}", "{{task2_status}}", "{{task2_progress}}"],
                        ["{{task3}}", "{{task3_owner}}", "{{task3_status}}", "{{task3_progress}}"],
                    ],
                },
            },
            {"heading": "风险与问题", "text": "{{risk}}"},
            {"heading": "下周计划", "bullets": ["{{next1}}", "{{next2}}", "{{next3}}"]},
            {"heading": "需协调事项", "text": "{{support}}"},
        ],
        "footer": "{{project}} · 周报",
    },
    {
        "id": "builtin-docx-meeting-minutes",
        "name": "会议纪要 · 石板灰",
        "format": "docx",
        "palette": "slate",
        "title": "{{title}}",
        "subtitle": "会议纪要  |  {{date}}",
        "meta": [
            ("时间", "{{time}}"),
            ("地点", "{{location}}"),
            ("主持", "{{host}}"),
            ("记录", "{{recorder}}"),
        ],
        "sections": [
            {
                "heading": "参会人员",
                "table": {"headers": ["姓名", "部门", "角色"], "rows": [["{{attendee1}}", "{{dept1}}", "{{role1}}"], ["{{attendee2}}", "{{dept2}}", "{{role2}}"], ["{{attendee3}}", "{{dept3}}", "{{role3}}"]]},
            },
            {"heading": "会议议程", "bullets": ["{{agenda1}}", "{{agenda2}}", "{{agenda3}}"]},
            {"heading": "讨论要点", "text": "{{discussion}}"},
            {"heading": "会议决议", "bullets": ["{{decision1}}", "{{decision2}}", "{{decision3}}"]},
            {
                "heading": "行动项",
                "table": {"headers": ["事项", "负责人", "截止日期", "状态"], "rows": [["{{action1}}", "{{action1_owner}}", "{{action1_due}}", "{{action1_status}}"], ["{{action2}}", "{{action2_owner}}", "{{action2_due}}", "{{action2_status}}"]]},
            },
        ],
        "footer": "{{company}} · 会议纪要",
    },
]
DOCX_SPECS.extend([
    {
        "id": "builtin-docx-contract",
        "name": "合同协议 · 深海军蓝",
        "format": "docx",
        "palette": "indigo",
        "title": "{{title}}",
        "subtitle": "合同编号：{{contract_no}}",
        "meta": [
            ("甲方", "{{party_a}}"),
            ("乙方", "{{party_b}}"),
            ("签订日期", "{{sign_date}}"),
            ("合同金额", "{{amount}}"),
        ],
        "sections": [
            {"heading": "合作内容", "text": "{{scope}}"},
            {"heading": "合同期限", "text": "{{term}}"},
            {"heading": "付款方式", "text": "{{payment_terms}}"},
            {"heading": "双方权利与义务", "bullets": ["{{obligation1}}", "{{obligation2}}", "{{obligation3}}"]},
            {"heading": "违约责任", "text": "{{liability}}"},
            {
                "heading": "签署",
                "table": {
                    "headers": ["甲方（盖章）", "乙方（盖章）"],
                    "rows": [["授权代表：{{party_a_sign}}", "授权代表：{{party_b_sign}}"], ["日期：{{party_a_date}}", "日期：{{party_b_date}}"]],
                },
            },
        ],
        "footer": "合同编号：{{contract_no}}",
    },
    {
        "id": "builtin-docx-resume",
        "name": "个人简历 · 极简墨绿",
        "format": "docx",
        "palette": "emerald",
        "title": "{{name}}",
        "subtitle": "{{job_title}}",
        "meta": [
            ("电话", "{{phone}}"),
            ("邮箱", "{{email}}"),
            ("城市", "{{city}}"),
            ("工作年限", "{{experience_years}}"),
        ],
        "sections": [
            {"heading": "个人简介", "text": "{{summary}}"},
            {"heading": "工作经历", "bullets": ["{{experience1}}", "{{experience2}}", "{{experience3}}"]},
            {"heading": "项目经历", "bullets": ["{{project1}}", "{{project2}}"]},
            {"heading": "教育背景", "text": "{{education}}"},
            {"heading": "专业技能", "text": "{{skills}}"},
        ],
        "footer": "{{name}} · {{job_title}}",
    },
    {
        "id": "builtin-docx-quotation",
        "name": "报价单 · 暖橙",
        "format": "docx",
        "palette": "orange",
        "title": "报价单",
        "subtitle": "报价编号：{{quote_no}}",
        "meta": [
            ("客户", "{{client}}"),
            ("联系人", "{{contact}}"),
            ("日期", "{{date}}"),
            ("有效期", "{{valid_until}}"),
        ],
        "sections": [
            {
                "heading": "报价明细",
                "table": {
                    "headers": ["项目", "规格", "数量", "单价", "金额"],
                    "rows": [
                        ["{{item1}}", "{{spec1}}", "{{qty1}}", "{{price1}}", "{{amount1}}"],
                        ["{{item2}}", "{{spec2}}", "{{qty2}}", "{{price2}}", "{{amount2}}"],
                        ["{{item3}}", "{{spec3}}", "{{qty3}}", "{{price3}}", "{{amount3}}"],
                        ["合计", "", "", "", "{{total}}"],
                    ],
                },
            },
            {"heading": "商务条款", "bullets": ["{{terms1}}", "{{terms2}}", "{{terms3}}"]},
            {"heading": "备注", "text": "{{notes}}"},
        ],
        "footer": "{{company}} · 报价单",
    },
    {
        "id": "builtin-docx-leave-request",
        "name": "请假申请 · 青碧",
        "format": "docx",
        "palette": "cyan",
        "title": "请假申请单",
        "subtitle": "{{department}} · {{name}}",
        "meta": [
            ("姓名", "{{name}}"),
            ("岗位", "{{position}}"),
            ("请假类型", "{{leave_type}}"),
            ("申请日期", "{{apply_date}}"),
            ("开始日期", "{{start_date}}"),
            ("结束日期", "{{end_date}}"),
            ("请假天数", "{{days}}"),
            ("紧急联系人", "{{emergency_contact}}"),
        ],
        "sections": [
            {"heading": "请假事由", "text": "{{reason}}"},
            {"heading": "工作交接", "text": "{{handover}}"},
            {
                "heading": "审批意见",
                "table": {"headers": ["角色", "姓名", "审批意见", "日期"], "rows": [["直属主管", "{{manager}}", "{{manager_comment}}", "{{manager_date}}"], ["部门负责人", "{{director}}", "{{director_comment}}", "{{director_date}}"]]},
            },
        ],
        "footer": "{{company}} · 人事表单",
    },
])
DOCX_SPECS.extend([
    {
        "id": "builtin-docx-prd",
        "name": "产品需求文档 · 科技紫",
        "format": "docx",
        "palette": "purple",
        "title": "{{product}} 产品需求文档",
        "subtitle": "版本 {{version}}  |  负责人 {{owner}}",
        "meta": [
            ("日期", "{{date}}"),
            ("状态", "{{status}}"),
            ("优先级", "{{priority}}"),
            ("评审人", "{{reviewer}}"),
        ],
        "sections": [
            {"heading": "背景与问题", "text": "{{background}}"},
            {"heading": "产品目标", "text": "{{goal}}"},
            {"heading": "范围说明", "text": "{{scope}}"},
            {
                "heading": "需求清单",
                "table": {
                    "headers": ["编号", "需求", "优先级", "验收标准"],
                    "rows": [
                        ["R-01", "{{requirement1}}", "{{priority1}}", "{{acceptance1}}"],
                        ["R-02", "{{requirement2}}", "{{priority2}}", "{{acceptance2}}"],
                        ["R-03", "{{requirement3}}", "{{priority3}}", "{{acceptance3}}"],
                    ],
                },
            },
            {
                "heading": "里程碑",
                "table": {"headers": ["阶段", "交付物", "负责人", "日期"], "rows": [["{{phase1}}", "{{deliverable1}}", "{{owner1}}", "{{date1}}"], ["{{phase2}}", "{{deliverable2}}", "{{owner2}}", "{{date2}}"]]},
            },
            {"heading": "风险与依赖", "text": "{{risk}}"},
        ],
        "footer": "{{product}} · PRD v{{version}}",
    },
    {
        "id": "builtin-docx-training-notice",
        "name": "培训通知 · 活力橙",
        "format": "docx",
        "palette": "amber",
        "title": "培训通知 · {{title}}",
        "subtitle": "{{department}}  |  {{date}}",
        "meta": [
            ("讲师", "{{trainer}}"),
            ("时间", "{{time}}"),
            ("地点", "{{location}}"),
            ("对象", "{{audience}}"),
        ],
        "sections": [
            {"heading": "培训简介", "text": "{{summary}}"},
            {
                "heading": "课程议程",
                "table": {"headers": ["时间", "主题", "内容", "讲师"], "rows": [["{{slot1_time}}", "{{slot1_topic}}", "{{slot1_content}}", "{{slot1_trainer}}"], ["{{slot2_time}}", "{{slot2_topic}}", "{{slot2_content}}", "{{slot2_trainer}}"], ["{{slot3_time}}", "{{slot3_topic}}", "{{slot3_content}}", "{{slot3_trainer}}"]]},
            },
            {"heading": "报名方式", "text": "{{registration}}"},
            {"heading": "联系人", "text": "{{contact}}"},
        ],
        "footer": "{{department}} · 培训通知",
    },
    {
        "id": "builtin-docx-retrospective",
        "name": "项目复盘 · 玫瑰红",
        "format": "docx",
        "palette": "rose",
        "title": "项目复盘 · {{project}}",
        "subtitle": "{{date}}  |  负责人 {{owner}}",
        "meta": [
            ("项目周期", "{{period}}"),
            ("参与人", "{{team}}"),
            ("项目目标", "{{goal}}"),
            ("项目结果", "{{result}}"),
        ],
        "sections": [
            {
                "heading": "关键指标",
                "table": {"headers": ["指标", "目标", "实际", "达成率"], "rows": [["{{metric1}}", "{{target1}}", "{{actual1}}", "{{rate1}}"], ["{{metric2}}", "{{target2}}", "{{actual2}}", "{{rate2}}"], ["{{metric3}}", "{{target3}}", "{{actual3}}", "{{rate3}}"]]},
            },
            {"heading": "做得好的地方", "bullets": ["{{good1}}", "{{good2}}", "{{good3}}"]},
            {"heading": "需要改进的地方", "bullets": ["{{improve1}}", "{{improve2}}", "{{improve3}}"]},
            {"heading": "经验教训", "text": "{{lesson}}"},
            {
                "heading": "行动项",
                "table": {"headers": ["行动", "负责人", "截止日期", "状态"], "rows": [["{{action1}}", "{{action1_owner}}", "{{action1_due}}", "{{action1_status}}"], ["{{action2}}", "{{action2_owner}}", "{{action2_due}}", "{{action2_status}}"]]},
            },
        ],
        "footer": "{{project}} · 项目复盘",
    },
])


def _no_cell_borders(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "nil")
        borders.append(element)


def _cell_margins(cell, top: int = 120, left: int = 180, bottom: int = 120, right: int = 180) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    margins = tc_pr.find(qn("w:tcMar"))
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tc_pr.append(margins)
    for edge, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        margins.append(node)


def _cover_text(cell, text: str, size: float, color: str, bold: bool = False, space_before: int = 0, space_after: int = 0) -> None:
    paragraph = cell.paragraphs[0] if not cell.paragraphs[0].text else cell.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(space_before)
    paragraph.paragraph_format.space_after = Pt(space_after)
    paragraph.paragraph_format.line_spacing = 1.15
    run = paragraph.add_run(text)
    _set_run(run, size, color, bold)


def _docx_cover(document, spec: dict[str, Any], palette: dict[str, str], variant: str) -> None:
    title = spec["title"]
    subtitle = spec["subtitle"]
    name_parts = spec["name"].split(" ? ")
    kicker = (name_parts[0] if name_parts else spec["name"]).upper()
    meta = spec["meta"]
    if variant == "dark":
        table = document.add_table(rows=1, cols=1)
        table.autofit = False
        table.columns[0].width = Cm(16.0)
        row = table.rows[0]
        row.height = Cm(20.5)
        cell = table.cell(0, 0)
        _shade_cell(cell, palette["dark"])
        _no_cell_borders(cell)
        _cell_margins(cell, 700, 700, 500, 700)
        _cover_text(cell, kicker, 10, palette["accent"], True, 0, 18)
        _cover_text(cell, title, 34, palette["white"], True, 20, 12)
        _cover_text(cell, subtitle, 14, palette["light"], False, 4, 30)
        accent = cell.add_paragraph()
        _bottom_border(accent, palette["accent"], 12)
        for label, value in meta[:4]:
            _cover_text(cell, f"{label}   {value}", 10, palette["light"], False, 2, 2)
        _cover_text(cell, "{{company}}  |  {{date}}", 9, palette["muted"], False, 50, 0)
    elif variant == "split":
        table = document.add_table(rows=1, cols=2)
        table.autofit = False
        table.columns[0].width = Cm(6.2)
        table.columns[1].width = Cm(9.8)
        row = table.rows[0]
        row.height = Cm(20.5)
        left, right = table.cell(0, 0), table.cell(0, 1)
        _shade_cell(left, palette["primary"])
        _shade_cell(right, palette["light"])
        for cell in (left, right):
            _no_cell_borders(cell)
            _cell_margins(cell, 600, 500, 400, 500)
        _cover_text(left, "FORMAT STUDIO", 9, palette["white"], True, 0, 12)
        _cover_text(left, kicker, 16, palette["white"], True, 10, 10)
        _cover_text(left, "{{company}}", 11, palette["light"], False, 10, 6)
        _cover_text(left, "{{date}}", 10, palette["light"], False, 4, 0)
        _cover_text(right, "REPORT / {{version}}", 9, palette["primary"], True, 0, 20)
        _cover_text(right, title, 24, palette["dark"], True, 20, 10)
        _cover_text(right, subtitle, 13, palette["muted"], False, 0, 28)
        for label, value in meta[:4]:
            _cover_text(right, f"{label}    {value}", 9.5, palette["dark"], False, 2, 2)
    elif variant == "minimal":
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(10)
        paragraph.paragraph_format.space_after = Pt(10)
        _bottom_border(paragraph, palette["primary"], 18)
        _text(document, kicker, palette)
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(50)
        run = paragraph.add_run(title)
        _set_run(run, 38, palette["dark"], True)
        paragraph = document.add_paragraph()
        run = paragraph.add_run(subtitle)
        _set_run(run, 14, palette["muted"])
        document.add_paragraph()
        table = document.add_table(rows=2, cols=4)
        table.autofit = True
        for index, (label, value) in enumerate(meta[:4]):
            cell = table.cell(0, index)
            cell.text = ""
            _shade_cell(cell, palette["light"])
            _no_cell_borders(cell)
            _cover_text(cell, label, 8.5, palette["muted"], True, 4, 2)
            value_cell = table.cell(1, index)
            value_cell.text = ""
            _no_cell_borders(value_cell)
            _cover_text(value_cell, value, 10, palette["dark"], False, 2, 8)
    elif variant == "magazine":
        _text(document, "INSIGHT  /  {{date}}", palette)
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(20)
        run = paragraph.add_run(title)
        _set_run(run, 42, palette["dark"], True)
        band = document.add_table(rows=1, cols=1)
        band.autofit = False
        band.columns[0].width = Cm(16.0)
        cell = band.cell(0, 0)
        _shade_cell(cell, palette["primary"])
        _no_cell_borders(cell)
        _cell_margins(cell, 220, 280, 220, 280)
        _cover_text(cell, subtitle, 14, palette["white"], True, 0, 0)
        document.add_paragraph()
        table = document.add_table(rows=2, cols=3)
        for index, (label, value) in enumerate(meta[:3]):
            label_cell = table.cell(0, index)
            value_cell = table.cell(1, index)
            label_cell.text = ""
            value_cell.text = ""
            _shade_cell(label_cell, palette["light"])
            _no_cell_borders(label_cell)
            _no_cell_borders(value_cell)
            _cover_text(label_cell, label, 8.5, palette["muted"], True, 10, 4)
            _cover_text(value_cell, value, 11, palette["dark"], True, 4, 10)
    else:
        _title_band(document, title, subtitle, palette)
        table = document.add_table(rows=1, cols=3)
        for index, (label, value) in enumerate(meta[:3]):
            cell = table.cell(0, index)
            cell.text = ""
            _shade_cell(cell, palette["light"])
            _no_cell_borders(cell)
            _cell_margins(cell, 180, 220, 180, 220)
            _cover_text(cell, label, 8.5, palette["muted"], True, 2, 2)
            _cover_text(cell, value, 11, palette["dark"], True, 2, 4)
    document.add_page_break()


def _section_banner(document, number: int, text: str, palette: dict[str, str]) -> None:
    table = document.add_table(rows=1, cols=2)
    table.autofit = False
    table.columns[0].width = Cm(1.4)
    table.columns[1].width = Cm(14.6)
    left, right = table.cell(0, 0), table.cell(0, 1)
    _shade_cell(left, palette["primary"])
    _shade_cell(right, palette["light"])
    for cell in (left, right):
        _no_cell_borders(cell)
        _cell_margins(cell, 100, 120, 100, 160)
    left.text = ""
    paragraph = left.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(f"{number:02d}")
    _set_run(run, 11, palette["white"], True)
    right.text = ""
    run = right.paragraphs[0].add_run(text)
    _set_run(run, 12.5, palette["primary"], True)


def _callout(document, text: str, palette: dict[str, str]) -> None:
    table = document.add_table(rows=1, cols=1)
    table.autofit = False
    table.columns[0].width = Cm(16.0)
    cell = table.cell(0, 0)
    _shade_cell(cell, palette["light"])
    _no_cell_borders(cell)
    _cell_margins(cell, 180, 240, 180, 240)
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.line_spacing = 1.55
    run = paragraph.add_run(text)
    _set_run(run, 10.5, palette["dark"])


def build_docx(spec: dict[str, Any], path: Path) -> None:
    palette = PALETTES[spec["palette"]]
    document = Document()
    section = document.sections[0]
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(2.4)
    section.right_margin = Cm(2.4)
    section.different_first_page_header_footer = True
    section.first_page_header.paragraphs[0].text = ""
    section.first_page_footer.paragraphs[0].text = ""
    normal = document.styles["Normal"]
    normal.font.name = "????"
    normal.font.size = Pt(10.5)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "????")
    variants = ["dark", "split", "minimal", "band", "magazine"]
    try:
        index = DOCX_SPECS.index(spec)
    except ValueError:
        index = 0
    _docx_cover(document, spec, palette, spec.get("variant", variants[index % len(variants)]))
    header = section.header
    header_paragraph = header.paragraphs[0]
    header_paragraph.text = ""
    run = header_paragraph.add_run("{{company}}  |  " + spec["name"].split(" ? ")[0])
    _set_run(run, 8.5, palette["muted"])
    _bottom_border(header_paragraph, palette["light"], 4)
    for number, item in enumerate(spec["sections"], start=1):
        if item.get("page_break"):
            document.add_page_break()
            continue
        _section_banner(document, number, item["heading"], palette)
        if "text" in item:
            _callout(document, item["text"], palette)
        if "bullets" in item:
            _bullets(document, item["bullets"], palette)
        if "table" in item:
            _table(document, item["table"]["headers"], item["table"]["rows"], palette)
        document.add_paragraph().paragraph_format.space_after = Pt(2)
    _footer(document, spec["footer"], palette)
    path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(path))


XLSX_SPECS: list[dict[str, Any]] = [
    {
        "id": "builtin-xlsx-project-plan",
        "name": "项目计划 · 蓝",
        "format": "xlsx",
        "palette": "blue",
        "title": "{{project}} 项目计划",
        "subtitle": "负责人：{{owner}}  |  周期：{{start_date}} 至 {{end_date}}",
        "meta": [("项目", "{{project}}"), ("负责人", "{{owner}}"), ("开始日期", "{{start_date}}"), ("结束日期", "{{end_date}}")],
        "sheet_name": "任务计划",
        "headers": ["任务", "负责人", "开始日期", "截止日期", "状态", "进度"],
        "rows": [
            ["需求梳理", "张三", "2026-01-05", "2026-01-08", "已完成", 1.0],
            ["方案设计", "李四", "2026-01-09", "2026-01-15", "进行中", 0.65],
            ["开发实施", "王五", "2026-01-16", "2026-02-10", "未开始", 0.0],
            ["测试验收", "赵六", "2026-02-11", "2026-02-18", "未开始", 0.0],
        ],
        "kpis": [("任务总数", "4"), ("进行中", "1"), ("已完成", "1")],
        "chart": "bar",
        "notes": "每周更新任务状态与进度；进度列使用 0-1 小数表示。",
    },
    {
        "id": "builtin-xlsx-budget",
        "name": "预算管理 · 墨绿",
        "format": "xlsx",
        "palette": "emerald",
        "title": "{{department}} 预算管理",
        "subtitle": "预算期间：{{period}}",
        "meta": [("部门", "{{department}}"), ("期间", "{{period}}"), ("编制人", "{{owner}}"), ("币种", "{{currency}}")],
        "sheet_name": "预算明细",
        "headers": ["类别", "预算金额", "实际支出", "差异", "使用率"],
        "rows": [
            ["人力成本", 120000, 82000, "=B4-C4", "=IF(B4=0,0,C4/B4)"],
            ["市场推广", 80000, 61000, "=B5-C5", "=IF(B5=0,0,C5/B5)"],
            ["差旅费用", 30000, 18000, "=B6-C6", "=IF(B6=0,0,C6/B6)"],
            ["办公采购", 20000, 9000, "=B7-C7", "=IF(B7=0,0,C7/B7)"],
        ],
        "kpis": [("总预算", "250000"), ("已使用", "170000"), ("执行率", "68%")],
        "chart": "pie",
        "notes": "实际支出按财务口径填写；差异为预算减实际。",
    },
    {
        "id": "builtin-xlsx-sales-dashboard",
        "name": "销售看板 · 橙",
        "format": "xlsx",
        "palette": "orange",
        "title": "{{region}} 销售看板",
        "subtitle": "期间：{{period}}  |  负责人：{{owner}}",
        "meta": [("区域", "{{region}}"), ("期间", "{{period}}"), ("负责人", "{{owner}}"), ("币种", "{{currency}}")],
        "sheet_name": "销售数据",
        "headers": ["月份", "销售额", "目标", "达成率"],
        "rows": [
            ["1月", 820000, 900000, "=B4/C4"],
            ["2月", 910000, 900000, "=B5/C5"],
            ["3月", 1050000, 1000000, "=B6/C6"],
            ["4月", 980000, 1050000, "=B7/C7"],
        ],
        "kpis": [("累计销售", "3,760,000"), ("平均达成率", "96%"), ("目标差额", "-240,000")],
        "chart": "bar",
        "notes": "达成率 = 销售额 / 目标；建议每周刷新数据。",
    },
    {
        "id": "builtin-xlsx-inventory",
        "name": "库存管理 · 青",
        "format": "xlsx",
        "palette": "teal",
        "title": "{{warehouse}} 库存管理",
        "subtitle": "盘点日期：{{date}}  |  负责人：{{owner}}",
        "meta": [("仓库", "{{warehouse}}"), ("盘点日期", "{{date}}"), ("负责人", "{{owner}}"), ("币种", "{{currency}}")],
        "sheet_name": "库存明细",
        "headers": ["SKU", "名称", "类别", "库存", "安全库存", "状态"],
        "rows": [
            ["SKU-001", "无线鼠标", "外设", 120, 50, '=IF(D4<E4,"需补货","正常")'],
            ["SKU-002", "机械键盘", "外设", 35, 40, '=IF(D5<E5,"需补货","正常")'],
            ["SKU-003", "USB-C 线", "配件", 260, 100, '=IF(D6<E6,"需补货","正常")'],
            ["SKU-004", "显示器支架", "配件", 18, 25, '=IF(D7<E7,"需补货","正常")'],
        ],
        "kpis": [("SKU 数", "4"), ("需补货", "2"), ("库存总量", "433")],
        "chart": "bar",
        "notes": "库存低于安全库存时状态自动标记为需补货。",
    },
]
XLSX_SPECS.extend([
    {
        "id": "builtin-xlsx-expense",
        "name": "费用报销 · 玫瑰",
        "format": "xlsx",
        "palette": "rose",
        "title": "{{employee}} 费用报销",
        "subtitle": "部门：{{department}}  |  月份：{{month}}",
        "meta": [("申请人", "{{employee}}"), ("部门", "{{department}}"), ("月份", "{{month}}"), ("审批人", "{{approver}}")],
        "sheet_name": "报销明细",
        "headers": ["日期", "类别", "说明", "金额"],
        "rows": [
            ["2026-01-06", "交通", "客户拜访打车", 86.0],
            ["2026-01-09", "餐饮", "项目组加班餐", 240.0],
            ["2026-01-15", "住宿", "出差住宿", 680.0],
            ["2026-01-20", "办公", "打印耗材", 120.0],
        ],
        "kpis": [("报销笔数", "4"), ("合计金额", "1126"), ("待审批", "4")],
        "chart": "bar",
        "notes": "请附发票或凭证；金额单位为人民币元。",
    },
    {
        "id": "builtin-xlsx-quotation",
        "name": "报价单 · 琥珀",
        "format": "xlsx",
        "palette": "amber",
        "title": "报价单 {{quote_no}}",
        "subtitle": "客户：{{client}}  |  日期：{{date}}",
        "meta": [("\u5ba2\u6237", "{{client}}"), ("\u8054\u7cfb\u4eba", "{{contact}}"), ("\u62a5\u4ef7\u65e5\u671f", "{{date}}"), ("\u6709\u6548\u671f", "{{valid_until}}")],
        "sheet_name": "报价明细",
        "headers": ["项目", "规格", "数量", "单价", "金额"],
        "rows": [
            ["产品 A", "标准版", 10, 1200, "=C4*D4"],
            ["产品 B", "专业版", 5, 2600, "=C5*D5"],
            ["实施服务", "人天", 8, 1500, "=C6*D6"],
            ["合计", "", "", "", "=SUM(E4:E6)"],
        ],
        "kpis": [("项目数", "3"), ("报价总额", "39,000"), ("有效期", "30 天")],
        "chart": "bar",
        "notes": "报价含税/不含税请在商务条款中注明；金额单位：元。",
    },
    {
        "id": "builtin-xlsx-attendance",
        "name": "考勤表 · 紫",
        "format": "xlsx",
        "palette": "purple",
        "title": "{{department}} 考勤表",
        "subtitle": "月份：{{month}}  |  负责人：{{owner}}",
        "meta": [("部门", "{{department}}"), ("月份", "{{month}}"), ("负责人", "{{owner}}"), ("工作日", "{{workdays}}")],
        "sheet_name": "考勤明细",
        "headers": ["姓名", "应出勤", "实出勤", "请假", "加班", "出勤率"],
        "rows": [
            ["张三", 22, 21, 1, 4, "=IF(B4=0,0,C4/B4)"],
            ["李四", 22, 22, 0, 2, "=IF(B5=0,0,C5/B5)"],
            ["王五", 22, 20, 2, 0, "=IF(B6=0,0,C6/B6)"],
            ["赵六", 22, 22, 0, 6, "=IF(B7=0,0,C7/B7)"],
        ],
        "kpis": [("团队人数", "4"), ("平均出勤率", "96%"), ("加班合计", "12h")],
        "chart": "bar",
        "notes": "出勤率 = 实出勤 / 应出勤；加班单位：小时。",
    },
    {
        "id": "builtin-xlsx-kpi",
        "name": "KPI 考核 · 靛蓝",
        "format": "xlsx",
        "palette": "indigo",
        "title": "{{employee}} KPI 考核",
        "subtitle": "部门：{{department}}  |  周期：{{period}}",
        "meta": [("姓名", "{{employee}}"), ("部门", "{{department}}"), ("周期", "{{period}}"), ("评估人", "{{reviewer}}")],
        "sheet_name": "KPI 明细",
        "headers": ["指标", "权重", "目标", "实际", "得分", "加权得分"],
        "rows": [
            ["销售额", 0.4, 1000000, 920000, "=MIN(1.2,D4/C4)", "=B4*E4"],
            ["客户满意度", 0.3, 95, 92, "=MIN(1.2,D5/C5)", "=B5*E5"],
            ["项目交付", 0.2, 10, 10, "=MIN(1.2,D6/C6)", "=B6*E6"],
            ["团队协作", 0.1, 100, 88, "=MIN(1.2,D7/C7)", "=B7*E7"],
        ],
        "kpis": [("综合得分", "0.93"), ("权重合计", "100%"), ("评级", "A")],
        "chart": "bar",
        "notes": "权重合计应为 100%；评分上限 120%。",
    },
    {
        "id": "builtin-xlsx-crm",
        "name": "客户管理 · 青碧",
        "format": "xlsx",
        "palette": "cyan",
        "title": "{{owner}} 客户管理",
        "subtitle": "月份：{{month}}  |  团队：{{team}}",
        "meta": [("负责人", "{{owner}}"), ("月份", "{{month}}"), ("团队", "{{team}}"), ("目标", "{{target}}")],
        "sheet_name": "客户明细",
        "headers": ["客户", "联系人", "阶段", "预计金额", "下一步", "下次跟进"],
        "rows": [
            ["星河科技", "李经理", "需求确认", 120000, "提交方案", "2026-02-03"],
            ["云帆制造", "王总", "商务谈判", 260000, "价格确认", "2026-02-05"],
            ["青禾教育", "赵老师", "初步接触", 80000, "需求访谈", "2026-02-06"],
            ["远山物流", "陈总", "已成交", 180000, "交付跟进", "2026-02-10"],
        ],
        "kpis": [("客户数", "4"), ("预计金额", "640,000"), ("本月目标", "{{target}}")],
        "chart": "bar",
        "notes": "阶段建议统一为：初步接触、需求确认、商务谈判、已成交、已流失。",
    },
    {
        "id": "builtin-xlsx-content-calendar",
        "name": "内容日历 · 石板灰",
        "format": "xlsx",
        "palette": "slate",
        "title": "{{brand}} 内容日历",
        "subtitle": "月份：{{month}}  |  负责人：{{owner}}",
        "meta": [("品牌", "{{brand}}"), ("月份", "{{month}}"), ("负责人", "{{owner}}"), ("渠道", "{{channel}}")],
        "sheet_name": "内容排期",
        "headers": ["日期", "渠道", "主题", "形式", "负责人", "状态"],
        "rows": [
            ["2026-02-02", "公众号", "春季新品预告", "图文", "张三", "已发布"],
            ["2026-02-05", "视频号", "产品使用技巧", "短视频", "李四", "制作中"],
            ["2026-02-09", "小红书", "用户案例分享", "图文", "王五", "待审核"],
            ["2026-02-12", "抖音", "直播预告", "短视频", "赵六", "待排期"],
        ],
        "kpis": [("计划内容", "4"), ("已发布", "1"), ("制作中", "1")],
        "chart": "bar",
        "notes": "状态建议：待排期、制作中、待审核、已发布。",
    },
])
def _legacy_build_xlsx(spec: dict[str, Any], path: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, PieChart, Reference
    from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.properties import PageSetupProperties

    palette = PALETTES[spec["palette"]]
    thin = Side(style="thin", color="E2E8F0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    workbook = Workbook()
    cover = workbook.active
    cover.title = "概览"
    cover.sheet_view.showGridLines = False
    for column, width in zip("ABCDEF", (18, 18, 18, 18, 18, 18)):
        cover.column_dimensions[column].width = width
    cover.merge_cells("A1:F2")
    cell = cover["A1"]
    cell.value = spec["title"]
    cell.fill = PatternFill("solid", fgColor=palette["primary"])
    cell.font = Font(name="微软雅黑", size=22, bold=True, color=palette["white"])
    cell.alignment = Alignment(vertical="center", horizontal="left", indent=2)
    cover.merge_cells("A3:F3")
    cell = cover["A3"]
    cell.value = spec["subtitle"]
    cell.fill = PatternFill("solid", fgColor=palette["light"])
    cell.font = Font(name="微软雅黑", size=11, color=palette["dark"])
    cell.alignment = Alignment(vertical="center", horizontal="left", indent=2)
    for index, (label, value) in enumerate(spec["meta"]):
        row = 5 + index // 2
        col = 1 + (index % 2) * 3
        label_cell = cover.cell(row=row, column=col, value=label)
        value_cell = cover.cell(row=row, column=col + 1, value=value)
        cover.merge_cells(start_row=row, start_column=col + 1, end_row=row, end_column=col + 2)
        label_cell.font = Font(name="微软雅黑", size=9, bold=True, color=palette["muted"])
        label_cell.fill = PatternFill("solid", fgColor=palette["light"])
        label_cell.border = border
        value_cell.font = Font(name="微软雅黑", size=10, color=palette["dark"])
        value_cell.border = border
        value_cell.alignment = Alignment(vertical="center")
    for index, (label, value) in enumerate(spec["kpis"][:3]):
        start_col = 1 + index * 2
        cover.merge_cells(start_row=8, start_column=start_col, end_row=8, end_column=start_col + 1)
        label_cell = cover.cell(row=8, column=start_col, value=label)
        label_cell.fill = PatternFill("solid", fgColor=palette["light"])
        label_cell.font = Font(name="微软雅黑", size=9, color=palette["muted"])
        label_cell.alignment = Alignment(horizontal="center")
        cover.merge_cells(start_row=9, start_column=start_col, end_row=10, end_column=start_col + 1)
        value_cell = cover.cell(row=9, column=start_col, value=value)
        value_cell.fill = PatternFill("solid", fgColor=palette["white"])
        value_cell.font = Font(name="微软雅黑", size=16, bold=True, color=palette["primary"])
        value_cell.alignment = Alignment(horizontal="center", vertical="center")
        for row in range(8, 11):
            for col in range(start_col, start_col + 2):
                cover.cell(row=row, column=col).border = border
    cover.merge_cells("A12:F13")
    note = cover["A12"]
    note.value = spec.get("notes", "")
    note.font = Font(name="微软雅黑", size=9, color=palette["muted"])
    note.alignment = Alignment(wrap_text=True, vertical="top")
    cover.print_area = "A1:F13"
    cover.page_setup.orientation = "portrait"
    cover.page_setup.fitToWidth = 1
    cover.page_setup.fitToHeight = 1
    cover.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    data = workbook.create_sheet(spec["sheet_name"])
    data.sheet_view.showGridLines = False
    headers = spec["headers"]
    last_col = len(headers)
    data.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    title_cell = data.cell(row=1, column=1, value=spec["title"])
    title_cell.fill = PatternFill("solid", fgColor=palette["dark"])
    title_cell.font = Font(name="微软雅黑", size=15, bold=True, color=palette["white"])
    title_cell.alignment = Alignment(vertical="center", indent=1)
    data.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    subtitle_cell = data.cell(row=2, column=1, value=spec["subtitle"])
    subtitle_cell.font = Font(name="微软雅黑", size=9, color=palette["muted"])
    subtitle_cell.alignment = Alignment(vertical="center", indent=1)
    for col_index, header in enumerate(headers, start=1):
        cell = data.cell(row=3, column=col_index, value=header)
        cell.fill = PatternFill("solid", fgColor=palette["primary"])
        cell.font = Font(name="微软雅黑", size=9.5, bold=True, color=palette["white"])
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    for row_index, row in enumerate(spec["rows"], start=4):
        for col_index, value in enumerate(row, start=1):
            cell = data.cell(row=row_index, column=col_index, value=value)
            cell.font = Font(name="微软雅黑", size=9.5, color=palette["dark"])
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if row_index % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=palette["light"])
            header = headers[col_index - 1]
            if "率" in header or "进度" in header or "达成" in header:
                cell.number_format = "0%"
            elif any(marker in header for marker in ("金额", "预算", "支出", "单价", "销售", "预计", "差异")):
                cell.number_format = "#,##0"
    last_row = 3 + len(spec["rows"])
    data.freeze_panes = "A4"
    data.auto_filter.ref = f"A3:{get_column_letter(last_col)}{last_row}"
    data.print_title_rows = "3:3"
    data.print_area = f"A1:{get_column_letter(last_col)}{last_row}"
    data.page_setup.orientation = "landscape"
    data.page_setup.fitToWidth = 1
    data.page_setup.fitToHeight = 1
    data.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    for col_index, header in enumerate(headers, start=1):
        letter = get_column_letter(col_index)
        width = max(12, min(30, sum(2 if ord(ch) > 127 else 1 for ch in header) + 6))
        data.column_dimensions[letter].width = width
        if any(marker in header for marker in ("率", "进度", "达成")):
            data.conditional_formatting.add(
                f"{letter}4:{letter}{last_row}",
                DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=palette["accent"]),
            )
        if "状态" in header:
            from openpyxl.formatting.rule import CellIsRule

            data.conditional_formatting.add(
                f"{letter}4:{letter}{last_row}",
                CellIsRule(operator="equal", formula=['"已完成"'], fill=PatternFill("solid", fgColor="DCFCE7")),
            )
    if spec.get("chart") == "bar":
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = spec["sheet_name"]
        chart.y_axis.title = headers[1]
        chart.x_axis.title = headers[0]
        chart.height = 8
        chart.width = 16
        chart.add_data(Reference(data, min_col=2, max_col=min(3, last_col), min_row=3, max_row=last_row), titles_from_data=True)
        chart.set_categories(Reference(data, min_col=1, min_row=4, max_row=last_row))
        data.add_chart(chart, f"{get_column_letter(last_col + 2)}3")
    else:
        chart = PieChart()
        chart.title = headers[1]
        chart.height = 8
        chart.width = 14
        chart.add_data(Reference(data, min_col=2, min_row=3, max_row=last_row), titles_from_data=True)
        chart.set_categories(Reference(data, min_col=1, min_row=4, max_row=last_row))
        data.add_chart(chart, f"{get_column_letter(last_col + 2)}3")

    guide = workbook.create_sheet("使用说明")
    guide.sheet_view.showGridLines = False
    guide.column_dimensions["A"].width = 100
    guide["A1"] = "模板使用说明"
    guide["A1"].font = Font(name="微软雅黑", size=16, bold=True, color=palette["primary"])
    guide["A3"] = "1. 在数据源中为 {{字段}} 提供对应值，一行数据生成一份文件。"
    guide["A4"] = "2. 表格中的示例数据用于展示样式，可在生成后替换为实际业务数据。"
    guide["A5"] = f"3. {spec.get('notes', '')}"
    for row in range(3, 6):
        guide.cell(row=row, column=1).font = Font(name="微软雅黑", size=10, color=palette["dark"])
        guide.cell(row=row, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    guide.row_dimensions[3].height = 28
    guide.row_dimensions[4].height = 28
    guide.print_area = "A1:A5"
    guide.page_setup.orientation = "portrait"
    guide.page_setup.fitToWidth = 1
    guide.page_setup.fitToHeight = 1
    guide.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(str(path))
PPTX_SPECS: list[dict[str, Any]] = [
    {
        "id": "builtin-pptx-business-proposal",
        "name": "商务提案 · 深海蓝",
        "format": "pptx",
        "palette": "blue",
        "title": "{{title}}",
        "subtitle": "商务提案  |  {{company}}",
        "slides": [
            {"title": "项目背景", "bullets": ["{{background}}", "当前业务面临的关键挑战", "为什么现在必须行动"]},
            {"title": "解决方案", "bullets": ["{{solution}}", "核心能力与差异化", "交付方式与保障"]},
            {"title": "实施计划", "bullets": ["{{plan}}", "阶段划分与关键里程碑", "双方协作机制"]},
            {"title": "预期收益", "kpis": [("效率提升", "{{efficiency}}"), ("成本优化", "{{cost}}"), ("上线周期", "{{timeline}}")], "bullets": ["{{benefit}}"]},
        ],
    },
    {
        "id": "builtin-pptx-quarterly-review",
        "name": "季度汇报 · 墨绿",
        "format": "pptx",
        "palette": "emerald",
        "title": "{{quarter}} 季度汇报",
        "subtitle": "{{department}}  |  {{presenter}}",
        "slides": [
            {"title": "核心数据", "kpis": [("营收", "{{revenue}}"), ("增长", "{{growth}}"), ("完成率", "{{completion}}")], "bullets": ["{{highlight}}"]},
            {"title": "业务进展", "bullets": ["{{progress1}}", "{{progress2}}", "{{progress3}}"]},
            {"title": "问题与改进", "bullets": ["{{issue1}}", "{{issue2}}", "{{improvement}}"]},
            {"title": "下季度目标", "bullets": ["{{goal1}}", "{{goal2}}", "{{goal3}}"]},
        ],
    },
    {
        "id": "builtin-pptx-product-launch",
        "name": "产品发布 · 活力橙",
        "format": "pptx",
        "palette": "orange",
        "title": "{{product}} 产品发布",
        "subtitle": "{{slogan}}  |  {{date}}",
        "slides": [
            {"title": "用户痛点", "bullets": ["{{pain1}}", "{{pain2}}", "{{pain3}}"]},
            {"title": "产品亮点", "kpis": [("效率", "{{efficiency}}"), ("体验", "{{experience}}"), ("成本", "{{cost}}")], "bullets": ["{{highlight}}"]},
            {"title": "应用场景", "bullets": ["{{scenario1}}", "{{scenario2}}", "{{scenario3}}"]},
            {"title": "上市计划", "bullets": ["{{launch1}}", "{{launch2}}", "{{launch3}}"]},
        ],
    },
    {
        "id": "builtin-pptx-project-kickoff",
        "name": "项目启动 · 青",
        "format": "pptx",
        "palette": "teal",
        "title": "{{project}} 项目启动会",
        "subtitle": "{{department}}  |  {{date}}",
        "slides": [
            {"title": "项目目标", "bullets": ["{{goal}}", "成功标准：{{success}}", "关键收益：{{benefit}}"]},
            {"title": "范围与里程碑", "bullets": ["{{scope}}", "{{milestone1}}", "{{milestone2}}"]},
            {"title": "团队分工", "bullets": ["{{owner1}}：{{role1}}", "{{owner2}}：{{role2}}", "{{owner3}}：{{role3}}"]},
            {"title": "风险与机制", "bullets": ["{{risk1}}", "{{risk2}}", "沟通机制：{{cadence}}"]},
        ],
    },
    {
        "id": "builtin-pptx-company-profile",
        "name": "公司介绍 · 深石板",
        "format": "pptx",
        "palette": "slate",
        "title": "{{company}}",
        "subtitle": "{{slogan}}",
        "slides": [
            {"title": "关于我们", "bullets": ["{{about}}", "成立时间：{{founded}}", "服务客户：{{clients}}"]},
            {"title": "核心能力", "bullets": ["{{capability1}}", "{{capability2}}", "{{capability3}}"]},
            {"title": "客户案例", "bullets": ["{{case1}}", "{{case2}}", "{{case3}}"]},
            {"title": "联系我们", "bullets": ["电话：{{phone}}", "邮箱：{{email}}", "地址：{{address}}"]},
        ],
    },
]
PPTX_SPECS.extend([
    {
        "id": "builtin-pptx-training",
        "name": "培训课件 · 紫",
        "format": "pptx",
        "palette": "purple",
        "title": "{{course}} 培训课件",
        "subtitle": "讲师：{{trainer}}  |  {{date}}",
        "slides": [
            {"title": "学习目标", "bullets": ["{{objective1}}", "{{objective2}}", "{{objective3}}"]},
            {"title": "核心概念", "bullets": ["{{concept1}}", "{{concept2}}", "{{concept3}}"]},
            {"title": "案例练习", "bullets": ["{{case}}", "练习任务：{{exercise}}", "讨论问题：{{question}}"]},
            {"title": "总结回顾", "kpis": [("知识点", "{{points}}"), ("练习", "{{exercises}}"), ("时长", "{{duration}}")], "bullets": ["{{summary}}"]},
        ],
    },
    {
        "id": "builtin-pptx-market-analysis",
        "name": "市场分析 · 靛蓝",
        "format": "pptx",
        "palette": "indigo",
        "title": "{{market}} 市场分析",
        "subtitle": "{{period}}  |  {{team}}",
        "slides": [
            {"title": "市场规模", "kpis": [("市场规模", "{{size}}"), ("增长率", "{{growth}}"), ("渗透率", "{{penetration}}")], "bullets": ["{{insight}}"]},
            {"title": "竞争格局", "bullets": ["{{competitor1}}", "{{competitor2}}", "{{competitor3}}"]},
            {"title": "用户洞察", "bullets": ["{{need1}}", "{{need2}}", "{{need3}}"]},
            {"title": "策略建议", "bullets": ["{{strategy1}}", "{{strategy2}}", "{{strategy3}}"]},
        ],
    },
    {
        "id": "builtin-pptx-annual-summary",
        "name": "年度总结 · 玫瑰",
        "format": "pptx",
        "palette": "rose",
        "title": "{{year}} 年度总结",
        "subtitle": "{{company}}  |  {{department}}",
        "slides": [
            {"title": "年度成绩", "kpis": [("营收", "{{revenue}}"), ("客户", "{{clients}}"), ("增长", "{{growth}}")], "bullets": ["{{highlight}}"]},
            {"title": "重点项目", "bullets": ["{{project1}}", "{{project2}}", "{{project3}}"]},
            {"title": "团队成长", "bullets": ["{{team1}}", "{{team2}}", "{{team3}}"]},
            {"title": "明年规划", "bullets": ["{{plan1}}", "{{plan2}}", "{{plan3}}"]},
        ],
    },
    {
        "id": "builtin-pptx-minimal-pitch",
        "name": "极简路演 · 黑白",
        "format": "pptx",
        "palette": "slate",
        "title": "{{project}}",
        "subtitle": "{{slogan}}",
        "slides": [
            {"title": "问题", "bullets": ["{{problem}}", "现有方案为何不够好", "用户真正需要什么"]},
            {"title": "方案", "bullets": ["{{solution}}", "核心差异化", "为什么是我们"]},
            {"title": "商业模式", "kpis": [("客单价", "{{price}}"), ("毛利率", "{{margin}}"), ("市场规模", "{{market}}")], "bullets": ["{{model}}"]},
            {"title": "团队", "bullets": ["{{team1}}", "{{team2}}", "{{team3}}"]},
        ],
    },
    {
        "id": "builtin-pptx-team-intro",
        "name": "团队介绍 · 琥珀",
        "format": "pptx",
        "palette": "amber",
        "title": "{{team}} 团队介绍",
        "subtitle": "{{company}}  |  {{date}}",
        "slides": [
            {"title": "团队概览", "kpis": [("成员", "{{members}}"), ("项目", "{{projects}}"), ("经验", "{{experience}}")], "bullets": ["{{mission}}"]},
            {"title": "核心成员", "bullets": ["{{member1}}：{{role1}}", "{{member2}}：{{role2}}", "{{member3}}：{{role3}}"]},
            {"title": "协作方式", "bullets": ["{{process1}}", "{{process2}}", "{{process3}}"]},
            {"title": "加入我们", "bullets": ["{{hiring}}", "联系方式：{{contact}}", "工作地点：{{location}}"]},
        ],
    },
])


def _legacy_build_pptx(spec: dict[str, Any], path: Path) -> None:
    from pptx import Presentation
    from pptx.dml.color import RGBColor as PptRGB
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Inches, Pt as PptPt

    palette = PALETTES[spec["palette"]]
    path.parent.mkdir(parents=True, exist_ok=True)
    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    blank = presentation.slide_layouts[6]
    width = presentation.slide_width
    height = presentation.slide_height

    def rgb(hex_value: str):
        return PptRGB.from_string(hex_value)

    def rect(slide, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE):
        item = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        item.fill.solid()
        item.fill.fore_color.rgb = rgb(color)
        item.line.fill.background()
        return item

    def text(slide, value, x, y, w, h, size, color, bold=False, align=PP_ALIGN.LEFT, font="微软雅黑"):
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        frame = box.text_frame
        frame.word_wrap = True
        frame.margin_left = 0
        frame.margin_right = 0
        frame.margin_top = 0
        frame.margin_bottom = 0
        frame.vertical_anchor = MSO_ANCHOR.TOP
        paragraph = frame.paragraphs[0]
        paragraph.alignment = align
        run = paragraph.add_run()
        run.text = value
        run.font.size = PptPt(size)
        run.font.bold = bold
        run.font.name = font
        run.font.color.rgb = rgb(color)
        return box

    def footer(slide, page: int) -> None:
        text(slide, spec["subtitle"], 0.9, 7.02, 8.0, 0.25, 8, palette["muted"])
        text(slide, str(page), 12.2, 7.02, 0.4, 0.25, 8, palette["muted"], align=PP_ALIGN.RIGHT)

    title_slide = presentation.slides.add_slide(blank)
    rect(title_slide, 0, 0, 13.333, 7.5, palette["dark"])
    rect(title_slide, 0, 0, 0.28, 7.5, palette["accent"])
    rect(title_slide, 10.9, 0.8, 1.7, 1.7, palette["primary"], MSO_SHAPE.OVAL)
    text(title_slide, spec["title"], 0.9, 2.0, 9.8, 1.5, 36, palette["white"], True)
    text(title_slide, spec["subtitle"], 0.95, 3.65, 9.5, 0.5, 17, palette["accent"])
    text(title_slide, "{{presenter}}  |  {{date}}", 0.95, 6.4, 7.0, 0.35, 11, palette["muted"])
    for index, slide_spec in enumerate(spec["slides"], start=2):
        slide = presentation.slides.add_slide(blank)
        rect(slide, 0, 0, 13.333, 7.5, "FFFFFF")
        rect(slide, 0, 0, 13.333, 0.12, palette["primary"])
        rect(slide, 0.65, 0.55, 0.12, 0.9, palette["accent"])
        text(slide, slide_spec["title"], 0.95, 0.55, 10.5, 0.75, 27, palette["dark"], True)
        text(slide, "{{company}}  ·  {{presenter}}", 0.97, 1.28, 8.0, 0.3, 10, palette["muted"])
        if slide_spec.get("kpis"):
            for card_index, (label, value) in enumerate(slide_spec["kpis"][:3]):
                x = 0.9 + card_index * 3.95
                rect(slide, x, 2.0, 3.45, 1.55, palette["light"], MSO_SHAPE.ROUNDED_RECTANGLE)
                text(slide, label, x + 0.25, 2.22, 2.9, 0.3, 10, palette["muted"])
                text(slide, value, x + 0.25, 2.62, 2.9, 0.65, 25, palette["primary"], True)
            bullet_top = 4.0
        else:
            bullet_top = 2.1
        bullets = slide_spec.get("bullets", [])
        for bullet_index, bullet in enumerate(bullets[:4]):
            y = bullet_top + bullet_index * 0.75
            rect(slide, 0.95, y + 0.11, 0.12, 0.12, palette["accent"], MSO_SHAPE.OVAL)
            text(slide, bullet, 1.25, y, 10.8, 0.55, 16, palette["dark"])
        footer(slide, index)
    closing = presentation.slides.add_slide(blank)
    rect(closing, 0, 0, 13.333, 7.5, palette["primary"])
    text(closing, "谢谢", 0.9, 2.5, 6.0, 1.0, 44, palette["white"], True)
    text(closing, "{{company}}  |  {{presenter}}", 0.95, 3.8, 8.0, 0.4, 16, palette["light"])
    text(closing, "{{contact}}", 0.95, 4.4, 8.0, 0.4, 13, palette["light"])
    presentation.save(str(path))
def build_xlsx(spec: dict[str, Any], path: Path) -> None:
    from .xlsx_builder import build_xlsx_rich

    try:
        index = XLSX_SPECS.index(spec)
    except ValueError:
        index = 0
    variant = spec.get("variant", ["dashboard", "executive", "tracker"][index % 3])
    chart_columns = {
        "builtin-xlsx-project-plan": [6],
        "builtin-xlsx-budget": [2],
        "builtin-xlsx-sales-dashboard": [2, 3],
        "builtin-xlsx-inventory": [4, 5],
        "builtin-xlsx-expense": [4],
        "builtin-xlsx-quotation": [5],
        "builtin-xlsx-attendance": [6],
        "builtin-xlsx-kpi": [6],
        "builtin-xlsx-crm": [4],
        "builtin-xlsx-content-calendar": [],
    }.get(spec["id"], [])
    rich_spec = dict(spec)
    rich_spec["chart_columns"] = chart_columns
    if not chart_columns:
        rich_spec["chart"] = None
    build_xlsx_rich(rich_spec, path, PALETTES, variant)


def build_pptx(spec: dict[str, Any], path: Path) -> None:
    from .pptx_builder import build_pptx_rich

    try:
        index = PPTX_SPECS.index(spec)
    except ValueError:
        index = 0
    build_pptx_rich(spec, path, PALETTES, index)


BUILTIN_SPECS: list[dict[str, Any]] = DOCX_SPECS + XLSX_SPECS + PPTX_SPECS
BUILDERS = {"docx": build_docx, "xlsx": build_xlsx, "pptx": build_pptx}


def ensure_builtin_templates(force: bool = False) -> int:
    root = TEMPLATES_DIR / "builtin"
    root.mkdir(parents=True, exist_ok=True)
    manifest_path = root / "manifest.json"
    manifest: dict[str, Any] = {}
    if manifest_path.exists():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except Exception:
            manifest = {}
    regenerate = force or manifest.get("version") != BUILTIN_VERSION
    count = 0
    for spec in BUILTIN_SPECS:
        folder = root / spec["id"]
        path = folder / f"template.{spec['format']}"
        try:
            if regenerate or not path.exists():
                BUILDERS[spec["format"]](spec, path)
            fields = scan_template(path, spec["format"])
            db.upsert_builtin_template(
                {
                    "id": spec["id"],
                    "name": spec["name"],
                    "format": spec["format"],
                    "path": str(path),
                    "original_name": f"{spec['name']}.{spec['format']}",
                    "fields": fields,
                }
            )
            count += 1
        except Exception:
            LOGGER.exception("failed to build builtin template %s", spec["id"])
    manifest_path.write_text(json.dumps({"version": BUILTIN_VERSION, "count": count}, ensure_ascii=False), encoding="utf-8")
    LOGGER.info("builtin templates ready: %s", count)
    return count
