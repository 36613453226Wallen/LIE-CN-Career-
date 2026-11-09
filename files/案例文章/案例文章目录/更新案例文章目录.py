#!/usr/bin/env python3
"""Generate 案例文章目录.docx from 案例目录.json."""

from __future__ import annotations

import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, RGBColor

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parents[1] / "指令库"))
from docx_style import (  # noqa: E402
    INK,
    NAVY,
    SLATE,
    add_bottom_border,
    compact,
    set_cell_margins,
    set_run_font,
    shade_cell,
)

DATA_PATH = ROOT / "案例目录.json"
OUTPUT_PATH = ROOT / "案例文章目录.docx"


def add_text(paragraph, text, *, name="宋体", size=10.5, bold=False, color=INK):
    run = paragraph.add_run(text)
    set_run_font(run, name, size, bold, color)
    return run


def generate():
    data = json.loads(DATA_PATH.read_text(encoding="utf-8"))
    items = data["articles"]

    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    kicker = doc.add_paragraph()
    compact(kicker, after=2, line=1.0)
    add_text(kicker, "LIE-CN-Career  ·  案例文章", name="黑体", size=10, bold=True, color=NAVY)

    title = doc.add_paragraph()
    compact(title, after=4, line=1.1)
    add_bottom_border(title, "1F4E79", "16")
    add_text(title, "案例文章目录", name="黑体", size=20, bold=True, color=NAVY)

    intro = doc.add_paragraph()
    compact(intro, before=6, after=10, line=1.15)
    add_text(
        intro,
        "每篇收录三件套：原版 PDF、原版 DOCX、分类好文段组 DOCX。"
        "分类文段另有可修改链接，不以聊天宣布为准。",
        size=10.5,
        color=INK,
    )

    table = doc.add_table(rows=1, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    headers = ["序号", "短标题", "快照", "缺项", "分类"]
    widths = [0.55, 2.4, 1.1, 1.6, 1.25]
    for i, (header, width) in enumerate(zip(headers, widths)):
        cell = table.rows[0].cells[i]
        cell.width = Inches(width)
        cell.text = ""
        shade_cell(cell, "1F4E79")
        set_cell_margins(cell, top=50, bottom=50, left=60, right=60)
        p = cell.paragraphs[0]
        compact(p, after=0, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_text(p, header, name="黑体", size=9, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

    for index, item in enumerate(items, start=1):
        row = table.add_row()
        missing = "；".join(item.get("missing") or []) or "未见缺项"
        values = [
            str(item["id"]),
            item["short_title"],
            item.get("snapshot_at") or "未见",
            missing,
            item.get("classified") or "事实与观点",
        ]
        for i, (value, width) in enumerate(zip(values, widths)):
            cell = row.cells[i]
            cell.width = Inches(width)
            cell.text = ""
            shade_cell(cell, "F4F7FA" if index % 2 == 0 else "FFFFFF")
            set_cell_margins(cell, top=50, bottom=50, left=70, right=70)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            p = cell.paragraphs[0]
            align = WD_ALIGN_PARAGRAPH.CENTER if i in (0, 2) else WD_ALIGN_PARAGRAPH.LEFT
            compact(p, after=0, line=1.08, align=align)
            add_text(p, value, name="宋体", size=9, color=INK)

    for item in items:
        heading = doc.add_paragraph()
        compact(heading, before=12, after=4, line=1.1)
        add_bottom_border(heading, "1F4E79", "10")
        add_text(heading, f"{item['id']}. {item['short_title']}", name="黑体", size=13, bold=True, color=NAVY)
        fields = [
            ("文件夹", item.get("folder") or ""),
            ("原版 PDF", item.get("pdf") or ""),
            ("原版 DOCX", item.get("docx") or ""),
            ("分类文段", item.get("classified_docx") or ""),
            ("可修改稿", item.get("classified_md") or ""),
            ("展示网页", item.get("classified_html") or ""),
            ("01V2 网页", item.get("v2_html") or ""),
            ("01V2 人名分色 OCR", item.get("v2_ocr_color_docx") or ""),
            ("可修改链接", item.get("classified_canvas") or ""),
            ("备份", item.get("backup") or ""),
            ("原文链接", item.get("original_url") or "未见"),
        ]
        for label, value in fields:
            p = doc.add_paragraph()
            compact(p, after=3, line=1.15)
            add_text(p, f"{label}：", name="黑体", size=10.5, bold=True, color=NAVY)
            add_text(p, value, size=10.5, color=INK)

    footer = doc.add_paragraph()
    compact(footer, before=12, after=0, line=1.15)
    add_text(
        footer,
        f"共 {len(items)} 篇｜更新 {data.get('updated_at')}｜数据源 案例目录.json",
        size=9,
        color=SLATE,
    )
    doc.core_properties.title = "案例文章目录"
    doc.save(OUTPUT_PATH)
    print(f"已生成：{OUTPUT_PATH}")


if __name__ == "__main__":
    generate()
