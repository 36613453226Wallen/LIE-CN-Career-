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
    body(doc, "下面是备份1-手机滚动截图太长识别会出错2张 里的原图，不转录为唯一正文。", size=10, color=SLATE)
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


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__</title>
<style>
  :root {
    --bg: #f7f7f5;
    --card: #ffffff;
    --ink: #1a1a1a;
    --muted: #667085;
    --line: #e6e6e1;
    --accent: #1f4e79;
    --fact: #1b5e8a;
    --opinion: #8a4b12;
    --warn: #7a5c00;
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    font: 15px/1.55 "Source Han Sans SC", "Noto Sans SC", "PingFang SC", "Microsoft YaHei", sans-serif;
    background: var(--bg);
    color: var(--ink);
  }
  main { max-width: 1080px; margin: 0 auto; padding: 32px 20px 64px; }
  h1 { font-size: 28px; font-weight: 650; margin: 0 0 6px; letter-spacing: -0.02em; }
  .sub { color: var(--muted); margin-bottom: 20px; }
  .note {
    background: #fff8e8;
    border: 1px solid #f0e0b8;
    border-radius: 10px;
    padding: 12px 14px;
    color: #5c4a12;
    margin-bottom: 20px;
  }
  .stats { display: flex; gap: 28px; flex-wrap: wrap; margin: 8px 0 20px; }
  .stat b { display: block; font-size: 28px; line-height: 1.1; }
  .stat span { color: var(--muted); font-size: 12px; }
  .stat.fact b { color: var(--fact); }
  .stat.opinion b { color: var(--opinion); }
  .stat.unsure b { color: var(--warn); }
  .toolbar { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; margin-bottom: 16px; }
  button.pill {
    border: 1px solid var(--line);
    background: var(--card);
    border-radius: 999px;
    padding: 6px 12px;
    cursor: pointer;
    color: var(--ink);
  }
  button.pill[aria-pressed="true"] { background: #222; color: #fff; border-color: #222; }
  button.linkish { margin-left: auto; background: none; border: 0; color: var(--muted); cursor: pointer; }
  table { width: 100%; border-collapse: collapse; background: var(--card); border: 1px solid var(--line); border-radius: 12px; overflow: hidden; }
  th, td { text-align: left; padding: 10px 12px; border-bottom: 1px solid var(--line); vertical-align: top; }
  th { font-size: 12px; color: var(--muted); font-weight: 600; background: #fafafa; }
  tr:last-child td { border-bottom: 0; }
  select, textarea {
    width: 100%;
    font: inherit;
    border: 1px solid var(--line);
    border-radius: 8px;
    padding: 6px 8px;
    background: #fff;
  }
  textarea { min-height: 72px; resize: vertical; }
  .speaker { font-weight: 600; }
  .muted { color: var(--muted); font-size: 12px; }
  .editor { margin-top: 24px; background: var(--card); border: 1px solid var(--line); border-radius: 12px; padding: 16px; }
  .editor h2 { margin: 0 0 12px; font-size: 18px; }
  footer { margin-top: 28px; color: var(--muted); font-size: 12px; }
</style>
</head>
<body>
<main>
  <h1>__H1__</h1>
  <p class="sub">__SUB__</p>
  <div class="note">这不是 Markdown。本页是一份<strong>自包含 HTML</strong>，任何浏览器都能打开筛选和改分类。数据写在页面里，不依赖 Cursor。改动只存在这台电脑的浏览器里，不会自动写回 Git。</div>
  <div class="stats" id="stats"></div>
  <div class="toolbar" id="filters"></div>
  <table>
    <thead>
      <tr><th>编号</th><th>发言人</th><th>类别</th><th>置信</th><th>摘录</th></tr>
    </thead>
    <tbody id="tbody"></tbody>
  </table>
  <section class="editor">
    <h2>改这一条</h2>
    <p class="muted" id="edit-meta"></p>
    <p><select id="pick"></select></p>
    <p><select id="kind"></select></p>
    <p><textarea id="text" placeholder="OCR 原文"></textarea></p>
    <p><textarea id="note" placeholder="备注、待核、OCR 校正"></textarea></p>
  </section>
  <footer>
    数据源：案例数据.json 生成。归档还有 DOCX 和 Markdown。Cursor Canvas 只是同内容的编辑器，不是给外人看的网页。
  </footer>
</main>
<script id="case-data" type="application/json">__JSON__</script>
<script>
const KINDS = ["事实", "观点", "事实+观点", "态度或追问", "存疑"];
const CASE = JSON.parse(document.getElementById("case-data").textContent);
const STORAGE_KEY = "lie-cn-case-" + CASE.id + "-rows";
const original = CASE.utterances.map(u => ({...u}));
let rows = loadRows();
let filter = "全部";
let selectedId = rows[0] ? rows[0].id : "";

function loadRows() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "null");
    if (Array.isArray(saved) && saved.length) return saved;
  } catch (e) {}
  return original.map(u => ({...u}));
}
function saveRows() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
}
function count(kind) { return rows.filter(r => r.kind === kind).length; }
function visible() { return filter === "全部" ? rows : rows.filter(r => r.kind === filter); }
function clip(t, n) { t = (t || "").replace(/\s+/g, " "); return t.length > n ? t.slice(0, n) + "…" : t; }
function speaker(u) { return u.reply_to ? u.speaker + " → " + u.reply_to : u.speaker; }

function renderStats() {
  document.getElementById("stats").innerHTML = [
    ["发言条", rows.length, ""],
    ["事实", count("事实"), "fact"],
    ["观点", count("观点"), "opinion"],
    ["事实+观点", count("事实+观点"), ""],
    ["存疑", count("存疑"), "unsure"]
  ].map(([label, n, cls]) => `<div class="stat ${cls}"><b>${n}</b><span>${label}</span></div>`).join("");
}
function renderFilters() {
  const labels = ["全部", ...KINDS];
  document.getElementById("filters").innerHTML =
    labels.map(l => `<button class="pill" data-filter="${l}" aria-pressed="${filter===l}">${l}</button>`).join("") +
    `<button class="linkish" id="reset">恢复初分</button>`;
  document.querySelectorAll("[data-filter]").forEach(btn => {
    btn.onclick = () => { filter = btn.dataset.filter; render(); };
  });
  document.getElementById("reset").onclick = () => {
    rows = original.map(u => ({...u}));
    saveRows();
    render();
  };
}
function renderTable() {
  const list = visible();
  document.getElementById("tbody").innerHTML = list.map(u => `
    <tr data-id="${u.id}" style="cursor:pointer">
      <td>${u.id}</td>
      <td class="speaker">${speaker(u)}</td>
      <td>${u.kind}</td>
      <td>${u.ocr_confidence || ""}</td>
      <td>${clip(u.text, 42)}</td>
    </tr>`).join("");
  document.querySelectorAll("tr[data-id]").forEach(tr => {
    tr.onclick = () => { selectedId = tr.dataset.id; renderEditor(); };
  });
}
function renderEditor() {
  const list = visible();
  const pick = document.getElementById("pick");
  pick.innerHTML = list.map(u => `<option value="${u.id}">${u.id} ${speaker(u)}</option>`).join("");
  if (!list.find(u => u.id === selectedId) && list[0]) selectedId = list[0].id;
  pick.value = selectedId;
  const u = rows.find(r => r.id === selectedId);
  if (!u) return;
  document.getElementById("edit-meta").textContent = `${u.source || ""} · 置信${u.ocr_confidence || ""} · ${u.subkind || ""}`;
  const kind = document.getElementById("kind");
  kind.innerHTML = KINDS.map(k => `<option${k===u.kind?" selected":""}>${k}</option>`).join("");
  document.getElementById("text").value = u.text || "";
  document.getElementById("note").value = u.note || "";
}
function bindEditor() {
  document.getElementById("pick").onchange = e => { selectedId = e.target.value; renderEditor(); };
  document.getElementById("kind").onchange = e => patch("kind", e.target.value);
  document.getElementById("text").oninput = e => patch("text", e.target.value);
  document.getElementById("note").oninput = e => patch("note", e.target.value);
}
function patch(field, value) {
  rows = rows.map(r => r.id === selectedId ? {...r, [field]: value} : r);
  saveRows();
  renderStats();
  renderTable();
}
function render() {
  renderStats();
  renderFilters();
  renderTable();
  renderEditor();
}
bindEditor();
render();
</script>
</body>
</html>
"""


def write_classified_html(case: dict, dest: Path):
    payload = json.dumps(case, ensure_ascii=False, separators=(",", ":"))
    payload = payload.replace("<", "\\u003c")
    html = (
        HTML_TEMPLATE.replace("__TITLE__", f"{case['id']} 分类好文段 · 事实与观点")
        .replace("__H1__", f"{case['id']} 分类好文段 · 事实与观点")
        .replace(
            "__SUB__",
            f"{case['short_title']} · {case.get('platform') or ''} · 快照 {case.get('snapshot_at') or ''}",
        )
        .replace("__JSON__", payload)
    )
    dest.write_text(html, encoding="utf-8")


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
    backup = ROOT.parents[0] / "备份1-手机滚动截图太长识别会出错2张"
    images = [backup / "图1-评论区.jpg", backup / "图2-评论区.jpg"]
    pdf = case_dir / f"{case['year']}-{case['id']:02d}{case['short_title']}.pdf"
    original_docx = case_dir / f"{case['id']:02d}{case['short_title']}.docx"
    classified_docx = case_dir / "01分类好文段-事实与观点.docx"
    classified_md = case_dir / "01分类好文段-事实与观点.md"
    classified_html = case_dir / "01分类好文段-事实与观点.html"
    ocr_md = case_dir / "ocr-原文.md"

    write_pdf(images, pdf)
    write_original_docx(case, case_dir, images, original_docx)
    write_classified_docx(case, classified_docx)
    write_classified_md(case, classified_md)
    write_classified_html(case, classified_html)
    write_ocr_md(case, ocr_md)
    print("pdf", pdf)
    print("docx", original_docx)
    print("classified", classified_docx)
    print("md", classified_md)
    print("html", classified_html)


if __name__ == "__main__":
    main()
