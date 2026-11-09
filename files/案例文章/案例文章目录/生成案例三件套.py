#!/usr/bin/env python3
"""Build original PDF/DOCX and classified fact/opinion files for one case."""

from __future__ import annotations

import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from PIL import Image

ROOT = Path(__file__).resolve().parent
LIB = ROOT.parents[1] / "指令库"
sys.path.insert(0, str(LIB))
from docx_style import (  # noqa: E402
    FACT,
    INK,
    MIXED,
    NAVY,
    OPINION,
    SLATE,
    UNSURE,
    add_bottom_border,
    compact,
    set_run_font,
)

KIND_COLOR = {
    "事实": FACT,
    "观点": OPINION,
    "事实+观点": MIXED,
    "态度或追问": RGBColor(0x2E, 0x6B, 0x4A),
    "存疑": UNSURE,
}

KIND_ORDER = ["事实", "观点", "事实+观点", "态度或追问", "存疑"]


def add_text(paragraph, text, *, name="宋体", size=10.5, bold=False, color=INK, italic=False):
    run = paragraph.add_run(text)
    set_run_font(run, name, size, bold, color, italic=italic)
    return run


def set_doc_margins(doc):
    section = doc.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)


def heading(doc, text, size=16):
    p = doc.add_paragraph()
    compact(p, before=8, after=6, line=1.1)
    add_bottom_border(p, "1F4E79", "12")
    add_text(p, text, name="黑体", size=size, bold=True, color=NAVY)
    return p


def body(doc, text, *, size=10.5, color=INK, before=0, after=6):
    p = doc.add_paragraph()
    compact(p, before=before, after=after, line=1.18)
    add_text(p, text, size=size, color=color)
    return p


def label_value(doc, label, value):
    p = doc.add_paragraph()
    compact(p, after=4, line=1.15)
    add_text(p, f"{label}：", name="黑体", size=10.5, bold=True, color=NAVY)
    add_text(p, value, size=10.5, color=INK)
    return p


def add_picture_keep_ratio(doc, path: Path, max_width_in=2.4):
    im = Image.open(path)
    w, h = im.size
    width = max_width_in
    height = width * (h / w)
    # cap height so Word can still open a tall phone screenshot
    if height > 8.8:
        height = 8.8
        width = height * (w / h)
    p = doc.add_paragraph()
    compact(p, after=8, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
    p.add_run().add_picture(str(path), width=Inches(width), height=Inches(height))
    cap = doc.add_paragraph()
    compact(cap, after=10, line=1.1, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(cap, path.name, size=9, color=SLATE, italic=True)


def write_pdf(images: list[Path], dest: Path):
    pages = [Image.open(p).convert("RGB") for p in images]
    first, rest = pages[0], pages[1:]
    first.save(dest, save_all=True, append_images=rest, resolution=150)
    for im in pages:
        im.close()


def write_original_docx(case: dict, case_dir: Path, images: list[Path], dest: Path):
    doc = Document()
    set_doc_margins(doc)
    heading(doc, f"{case['id']} {case['short_title']}", 18)
    body(doc, "原版 DOCX。评论截图贴在 OCR 文稿之后，保留头像，不重排原图。", size=10, color=SLATE)

    heading(doc, "缺项提醒", 13)
    for item in case.get("missing") or ["未见缺项"]:
        body(doc, f"缺：{item}", after=3)
    body(doc, "以上缺项不编造。本篇只有评论回复截图，没有完整回答正文。", size=10, color=SLATE)

    heading(doc, "著录", 13)
    for label, key in [
        ("问题标题", "question_title"),
        ("问题描述", "question_description"),
        ("原作者", "author"),
        ("原文链接", "original_url"),
        ("首次发布", "first_published"),
        ("后续编辑", "edited_at"),
        ("快照时间", "snapshot_at"),
        ("备份文件夹", "backup_folder"),
        ("OCR 说明", "ocr_note"),
    ]:
        label_value(doc, label, case.get(key) or "未见")

    heading(doc, "OCR 文稿（按发言条）", 13)
    body(doc, "机器识别不稳。低置信条不要当引用原文。", size=10, color=SLATE)
    for u in case["utterances"]:
        p = doc.add_paragraph()
        compact(p, before=8, after=2, line=1.12)
        add_text(p, f"{u['id']}  {u['speaker']}", name="黑体", size=11, bold=True, color=NAVY)
        if u.get("reply_to"):
            add_text(p, f"  → {u['reply_to']}", size=10, color=SLATE)
        add_text(p, f"  · {u['source']} · 置信{u['ocr_confidence']}", size=9, color=SLATE)
        body(doc, u["text"], after=2)
        if u.get("note"):
            body(doc, f"注：{u['note']}", size=9, color=SLATE, after=6)

    heading(doc, "评论原图", 13)
    body(doc, "下面是备份1里的原图，不转录为唯一正文。", size=10, color=SLATE)
    for img in images:
        add_picture_keep_ratio(doc, img)

    heading(doc, "原文链接", 13)
    body(doc, case.get("original_url") or "未见")

    doc.core_properties.title = f"{case['id']} {case['short_title']}"
    doc.save(dest)


def split_clauses(u: dict) -> list[dict]:
    """Keep whole utterance as one clause; mixed kinds stay mixed unless note splits them."""
    return [u]


def write_classified_docx(case: dict, dest: Path):
    doc = Document()
    set_doc_margins(doc)
    heading(doc, f"{case['id']} 分类好文段 · 事实与观点", 18)
    body(
        doc,
        "可改稿以 Canvas 和同名 Markdown 为准。这份 DOCX 是归档副本。"
        "事实=发言里当成可核验的陈述（对错另标待核）；观点=判断、态度、建议；"
        "OCR 不稳进存疑，不升格成事实。",
        size=10.5,
        color=INK,
    )
    label_value(doc, "短标题", case["short_title"])
    label_value(doc, "快照", case["snapshot_at"])
    label_value(doc, "可改 Markdown", "01分类好文段-事实与观点.md")

    grouped: dict[str, list] = {k: [] for k in KIND_ORDER}
    for u in case["utterances"]:
        grouped.setdefault(u["kind"], []).append(u)

    for kind in KIND_ORDER:
        items = grouped.get(kind) or []
        if not items:
            continue
        heading(doc, f"{kind}（{len(items)}）", 13)
        for u in items:
            p = doc.add_paragraph()
            compact(p, before=8, after=2, line=1.12)
            add_text(p, u["speaker"], name="黑体", size=11, bold=True, color=KIND_COLOR.get(kind, INK))
            if u.get("reply_to"):
                add_text(p, f"  → {u['reply_to']}", size=10, color=SLATE)
            add_text(p, f"  · {u['id']} · {u['source']} · 置信{u['ocr_confidence']}", size=9, color=SLATE)
            q = doc.add_paragraph()
            compact(q, after=2, line=1.18)
            add_text(q, u["text"], size=10.5, color=INK)
            meta = []
            if u.get("subkind"):
                meta.append(u["subkind"])
            if u.get("note"):
                meta.append(u["note"])
            if meta:
                body(doc, "；".join(meta), size=9, color=SLATE, after=6)

    footer = doc.add_paragraph()
    compact(footer, before=12, after=0, line=1.15)
    add_text(footer, "改分类请改 JSON 或 Canvas，再跑 生成案例三件套.py。", size=9, color=SLATE)
    doc.core_properties.title = f"{case['id']} 分类好文段-事实与观点"
    doc.save(dest)


def write_classified_md(case: dict, dest: Path):
    lines = [
        f"# {case['id']} 分类好文段 · 事实与观点",
        "",
        f"短标题：**{case['short_title']}**  ·  快照 {case['snapshot_at']}",
        "",
        "这是分类文段的可改稿。改完类别、OCR 或备注后，把对应改动写回 `案例数据.json`，或直接在对话里说改哪一条。",
        "",
        "规则：",
        "- **事实**：发言里当成可核验的陈述。对错另标待核，不因为可疑就改成观点。",
        "- **观点**：判断、态度、建议、价值判断。",
        "- **事实+观点**：同一条里两句功能不同。",
        "- **态度或追问**：疑问、反问、调侃，没有新的可核验内容。",
        "- **存疑**：OCR 不稳，不能当引用。",
        "",
    ]
    grouped: dict[str, list] = {k: [] for k in KIND_ORDER}
    for u in case["utterances"]:
        grouped.setdefault(u["kind"], []).append(u)
    for kind in KIND_ORDER:
        items = grouped.get(kind) or []
        if not items:
            continue
        lines.append(f"## {kind}（{len(items)}）")
        lines.append("")
        for u in items:
            reply = f" → {u['reply_to']}" if u.get("reply_to") else ""
            lines.append(f"### {u['id']} {u['speaker']}{reply}")
            lines.append("")
            lines.append(f"> {u['text']}")
            lines.append("")
            lines.append(
                f"来源 {u['source']}  ·  置信 {u['ocr_confidence']}  ·  {u.get('subkind') or ''}"
            )
            if u.get("note"):
                lines.append("")
                lines.append(f"注：{u['note']}")
            lines.append("")
    dest.write_text("\n".join(lines), encoding="utf-8")


def write_ocr_md(case: dict, dest: Path):
    lines = [
        f"# {case['id']} OCR 原稿",
        "",
        case.get("ocr_note") or "",
        "",
    ]
    for u in case["utterances"]:
        reply = f" → {u['reply_to']}" if u.get("reply_to") else ""
        lines.append(f"## {u['id']} {u['speaker']}{reply}")
        lines.append("")
        lines.append(u["text"])
        lines.append("")
        lines.append(f"来源 {u['source']}  ·  置信 {u['ocr_confidence']}  ·  初分 {u['kind']}")
        if u.get("note"):
            lines.append("")
            lines.append(f"注：{u['note']}")
        lines.append("")
    dest.write_text("\n".join(lines), encoding="utf-8")


def main():
    case_dir = ROOT / "01巴掌大的地方没有存在的必要"
    case = json.loads((case_dir / "案例数据.json").read_text(encoding="utf-8"))
    backup = ROOT.parents[0] / "备份1"
    images = [backup / "图1-评论区.jpg", backup / "图2-评论区.jpg"]
    pdf = case_dir / f"{case['year']}-{case['id']:02d}{case['short_title']}.pdf"
    original_docx = case_dir / f"{case['id']:02d}{case['short_title']}.docx"
    classified_docx = case_dir / "01分类好文段-事实与观点.docx"
    classified_md = case_dir / "01分类好文段-事实与观点.md"
    ocr_md = case_dir / "ocr-原文.md"

    write_pdf(images, pdf)
    write_original_docx(case, case_dir, images, original_docx)
    write_classified_docx(case, classified_docx)
    write_classified_md(case, classified_md)
    write_ocr_md(case, ocr_md)
    print("pdf", pdf)
    print("docx", original_docx)
    print("classified", classified_docx)
    print("md", classified_md)


if __name__ == "__main__":
    main()
