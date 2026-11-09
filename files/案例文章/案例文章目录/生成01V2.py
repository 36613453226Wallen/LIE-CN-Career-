#!/usr/bin/env python3
"""Build 01V2 grouped fact/opinion files and a name-colored OCR DOCX.

Does not touch 01 分类好文段-事实与观点.* or 案例数据.json.
"""

from __future__ import annotations

import json
import sys
from collections import defaultdict
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, RGBColor
from PIL import Image

ROOT = Path(__file__).resolve().parent
CASE_DIR = ROOT / "01巴掌大的地方没有存在的必要"
LIB = ROOT.parents[1] / "指令库"
sys.path.insert(0, str(LIB))
from docx_style import (  # noqa: E402
    INK,
    NAVY,
    SLATE,
    add_bottom_border,
    compact,
    set_run_font,
)

ALIAS = {"地委最强洗地之族": "地委最强洗地"}

KIND_ORDER = ["事实", "观点", "事实+观点", "态度或追问", "存疑"]

NAME_PALETTE = [
    RGBColor(0x1F, 0x4E, 0x79),
    RGBColor(0xC0, 0x39, 0x2B),
    RGBColor(0x1E, 0x84, 0x4E),
    RGBColor(0x8E, 0x44, 0xAD),
    RGBColor(0xD3, 0x54, 0x00),
    RGBColor(0x16, 0xA0, 0x85),
    RGBColor(0xC2, 0x18, 0x5B),
    RGBColor(0x29, 0x80, 0xB9),
    RGBColor(0x7D, 0x66, 0x0D),
    RGBColor(0x6C, 0x34, 0x8D),
    RGBColor(0x1A, 0x52, 0x7A),
    RGBColor(0x0E, 0x66, 0x5A),
    RGBColor(0xE6, 0x7E, 0x22),
    RGBColor(0xB7, 0x95, 0x0B),
    RGBColor(0x2C, 0x3E, 0x50),
    RGBColor(0x4A, 0x67, 0x2B),
    RGBColor(0x7F, 0x8C, 0x8D),
    RGBColor(0x92, 0x2B, 0x21),
]

def canon(name: str) -> str:
    name = (name or "").strip()
    return ALIAS.get(name, name)


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


def case_images(case: dict) -> list[tuple[Path, str]]:
    out = []
    for item in case.get("source_images") or []:
        name = Path(item.get("file") or "").name
        local = CASE_DIR / name
        if not local.exists():
            continue
        caption = " ".join(part for part in (item.get("id"), item.get("role")) if part)
        out.append((local, caption or name))
    return out


def add_picture_keep_ratio(doc, path: Path, caption: str, max_width_in=2.4):
    """Phone screenshots: keep ratio, cap height so one image stays readable on a page."""
    im = Image.open(path)
    w, h = im.size
    im.close()
    width = max_width_in
    height = width * (h / w)
    if height > 8.8:
        height = 8.8
        width = height * (w / h)
    p = doc.add_paragraph()
    compact(p, after=4, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
    p.add_run().add_picture(str(path), width=Inches(width), height=Inches(height))
    cap = doc.add_paragraph()
    compact(cap, after=12, line=1.1, align=WD_ALIGN_PARAGRAPH.CENTER)
    add_text(cap, caption, size=9, color=SLATE, italic=True)


def heading(doc, text, size=16):
    p = doc.add_paragraph()
    compact(p, before=8, after=6, line=1.1)
    add_bottom_border(p, "1F4E79", "12")
    add_text(p, text, name="黑体", size=size, bold=True, color=NAVY)
    return p


def collect_names(utterances: list[dict]) -> list[str]:
    seen: list[str] = []
    for u in utterances:
        for raw in (u.get("speaker"), u.get("reply_to")):
            name = canon(raw)
            if name and name not in seen:
                seen.append(name)
    return seen


def name_colors(names: list[str]) -> dict[str, RGBColor]:
    return {n: NAME_PALETTE[i % len(NAME_PALETTE)] for i, n in enumerate(names)}


def build_groups(utterances: list[dict]) -> list[dict]:
    own: dict[str, list] = defaultdict(list)
    incoming: dict[str, list] = defaultdict(list)
    first_index: dict[str, int] = {}

    def mark(name: str, idx: int):
        if name and name not in first_index:
            first_index[name] = idx

    for i, u in enumerate(utterances):
        speaker = canon(u.get("speaker"))
        target = canon(u.get("reply_to"))
        mark(speaker, i)
        mark(target, i)
        if speaker:
            own[speaker].append(u)
        if target:
            incoming[target].append(u)

    heads = set(own) | set(incoming)

    def sort_key(name: str):
        has_replies = 0 if incoming.get(name) else 1
        return (has_replies, -len(incoming.get(name) or []), first_index.get(name, 10**6))

    groups = []
    for name in sorted(heads, key=sort_key):
        replies = incoming.get(name) or []
        repliers: dict[str, list] = defaultdict(list)
        for u in replies:
            repliers[canon(u.get("speaker"))].append(u)
        replier_order = sorted(repliers, key=lambda n: (first_index.get(n, 10**6), n))
        groups.append(
            {
                "name": name,
                "own": own.get(name) or [],
                "replies": replies,
                "repliers": [{"name": n, "items": repliers[n]} for n in replier_order],
                "has_replies": bool(replies),
            }
        )
    return groups


def clip(text: str, n=42) -> str:
    t = (text or "").replace("\n", " ")
    return t if len(t) <= n else t[:n] + "…"


def write_v2_md(case: dict, groups: list[dict], dest: Path):
    lines = [
        f"# 01V2-分类好文段-事实与观点",
        "",
        f"短标题：**{case['short_title']}**  ·  快照 {case['snapshot_at']}",
        "",
        "01 原版保留。V2 按「开头的人」分组：箭头表示回复。",
        "次级只列出**回复了这个人**的发言人；没有人回复他，就没有次级选项。",
        "",
    ]
    for g in groups:
        own_n = len(g["own"])
        reply_n = len(g["replies"])
        lines.append(f"## {g['name']}")
        lines.append("")
        lines.append(f"本人发言 {own_n} 条 · 被回复 {reply_n} 条")
        lines.append("")
        if g["own"]:
            lines.append("### 本人发言")
            lines.append("")
            for u in g["own"]:
                target = canon(u.get("reply_to"))
                reply_bit = f" （回复 {target}）" if target else " （顶层）"
                lines.append(f"- **{u['id']}**{reply_bit} · {u['kind']} · 置信{u['ocr_confidence']}")
                lines.append("")
                lines.append(f"  > {u['text']}")
                lines.append("")
        if g["has_replies"]:
            lines.append("### 回复他的人")
            lines.append("")
            for r in g["repliers"]:
                lines.append(f"#### {r['name']}（{len(r['items'])}）")
                lines.append("")
                for u in r["items"]:
                    lines.append(f"- **{u['id']}** · {u['kind']} · 置信{u['ocr_confidence']}")
                    lines.append("")
                    lines.append(f"  > {u['text']}")
                    lines.append("")
        else:
            lines.append("没有人回复他，故无次级选项。")
            lines.append("")
    dest.write_text("\n".join(lines), encoding="utf-8")


def write_v2_docx(case: dict, groups: list[dict], dest: Path):
    doc = Document()
    set_doc_margins(doc)
    heading(doc, "01V2-分类好文段 · 事实与观点", 18)
    p = doc.add_paragraph()
    compact(p, after=8, line=1.18)
    add_text(
        p,
        "01 原版保留。V2 按开头的人分组，箭头=回复。有人回复他才有次级「对侧回复人」。网页里点对侧某人，只看两人按时间轴的对话。文末附原评论截图。",
        size=10.5,
        color=INK,
    )
    for g in groups:
        heading(doc, g["name"], 14)
        p = doc.add_paragraph()
        compact(p, after=6, line=1.15)
        add_text(p, f"本人发言 {len(g['own'])} 条 · 被回复 {len(g['replies'])} 条", size=10, color=SLATE)
        if g["own"]:
            p = doc.add_paragraph()
            compact(p, before=4, after=4, line=1.1)
            add_text(p, "本人发言", name="黑体", size=12, bold=True, color=NAVY)
            for u in g["own"]:
                target = canon(u.get("reply_to"))
                bit = f"回复 {target}" if target else "顶层"
                q = doc.add_paragraph()
                compact(q, before=6, after=2, line=1.12)
                add_text(q, f"{u['id']}  ", name="黑体", size=11, bold=True, color=NAVY)
                add_text(q, f"{bit} · {u['kind']} · 置信{u['ocr_confidence']}", size=9, color=SLATE)
                body = doc.add_paragraph()
                compact(body, after=4, line=1.18)
                add_text(body, u["text"], size=10.5, color=INK)
        if g["has_replies"]:
            p = doc.add_paragraph()
            compact(p, before=8, after=4, line=1.1)
            add_text(p, "回复他的人", name="黑体", size=12, bold=True, color=NAVY)
            for r in g["repliers"]:
                h = doc.add_paragraph()
                compact(h, before=6, after=2, line=1.1)
                add_text(h, f"回复 · {r['name']}（{len(r['items'])}）", name="黑体", size=11, bold=True, color=NAVY)
                for u in r["items"]:
                    q = doc.add_paragraph()
                    compact(q, before=4, after=2, line=1.12)
                    add_text(q, f"{u['id']}  ", name="黑体", size=11, bold=True, color=NAVY)
                    add_text(q, f"{u['kind']} · 置信{u['ocr_confidence']}", size=9, color=SLATE)
                    body = doc.add_paragraph()
                    compact(body, after=4, line=1.18)
                    add_text(body, u["text"], size=10.5, color=INK)
        else:
            p = doc.add_paragraph()
            compact(p, after=8, line=1.15)
            add_text(p, "没有人回复他，故无次级选项。", size=10, color=SLATE)

    heading(doc, "评论原图", 14)
    p = doc.add_paragraph()
    compact(p, after=8, line=1.18)
    add_text(
        p,
        "文段之后附原评论截图，按手机竖图比例缩小，方便对照 OCR，不重排原图。",
        size=10.5,
        color=INK,
    )
    images = case_images(case)
    if not images:
        p = doc.add_paragraph()
        compact(p, after=8, line=1.15)
        add_text(p, "未见截图文件。", size=10, color=SLATE)
    for path, caption in images:
        add_picture_keep_ratio(doc, path, caption)

    doc.core_properties.title = "01V2-分类好文段-事实与观点"
    doc.save(dest)


def split_by_names(text: str, names: list[str]) -> list[tuple[str, str | None]]:
    """Split text into (chunk, matched_name_or_none), longest names first."""
    ordered = sorted([n for n in names if n], key=len, reverse=True)
    if not text:
        return []
    parts: list[tuple[str, str | None]] = []
    i = 0
    while i < len(text):
        hit = None
        for n in ordered:
            if text.startswith(n, i):
                hit = n
                break
        if hit:
            parts.append((hit, hit))
            i += len(hit)
        else:
            if parts and parts[-1][1] is None:
                parts[-1] = (parts[-1][0] + text[i], None)
            else:
                parts.append((text[i], None))
            i += 1
    return parts


def write_colored_ocr_docx(case: dict, names: list[str], colors: dict[str, RGBColor], dest: Path):
    doc = Document()
    set_doc_margins(doc)
    heading(doc, "01V2-OCR 人名分色", 18)
    p = doc.add_paragraph()
    compact(p, after=8, line=1.18)
    add_text(p, "同一份 OCR 按发言顺序。每个人名固定一种颜色。图例如下。", size=10.5, color=INK)

    heading(doc, "人名颜色图例", 13)
    for n in names:
        row = doc.add_paragraph()
        compact(row, after=2, line=1.15)
        add_text(row, "■ ", size=12, bold=True, color=colors[n])
        add_text(row, n, name="黑体", size=11, bold=True, color=colors[n])

    heading(doc, "OCR 正文", 13)
    for u in case["utterances"]:
        speaker = canon(u.get("speaker"))
        target = canon(u.get("reply_to"))
        head = doc.add_paragraph()
        compact(head, before=10, after=2, line=1.12)
        add_text(head, f"{u['id']}  ", name="黑体", size=11, bold=True, color=NAVY)
        add_text(head, "发言人 ", size=10, color=SLATE)
        add_text(head, speaker or "未见", name="黑体", size=11, bold=True, color=colors.get(speaker, INK))
        if target:
            add_text(head, "  回复 ", size=10, color=SLATE)
            add_text(head, target, name="黑体", size=11, bold=True, color=colors.get(target, INK))
        else:
            add_text(head, "  顶层", size=10, color=SLATE)
        add_text(head, f"  · {u['kind']} · 置信{u['ocr_confidence']} · {u.get('source') or ''}", size=9, color=SLATE)

        body = doc.add_paragraph()
        compact(body, after=4, line=1.2)
        for chunk, matched in split_by_names(u.get("text") or "", names):
            if matched:
                add_text(body, chunk, size=11, bold=True, color=colors[matched])
            else:
                add_text(body, chunk, size=11, color=INK)
        if u.get("note"):
            note = doc.add_paragraph()
            compact(note, after=6, line=1.15)
            add_text(note, "注：", size=9, bold=True, color=SLATE)
            add_text(note, u["note"], size=9, color=SLATE)

    doc.core_properties.title = "01V2-OCR人名分色"
    doc.save(dest)


HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>01V2-分类好文段-事实与观点</title>
<style>
  :root { --bg:#f7f7f5; --card:#fff; --ink:#1a1a1a; --muted:#667085; --line:#e6e6e1; --navy:#1f4e79; }
  * { box-sizing: border-box; }
  body { margin:0; font:15px/1.55 "Source Han Sans SC","Noto Sans SC","PingFang SC","Microsoft YaHei",sans-serif; background:var(--bg); color:var(--ink); }
  main { max-width:1080px; margin:0 auto; padding:32px 20px 64px; }
  h1 { font-size:28px; margin:0 0 6px; letter-spacing:-0.02em; }
  .sub { color:var(--muted); margin-bottom:16px; }
  .note { background:#eef5fb; border:1px solid #d5e4f2; border-radius:10px; padding:12px 14px; margin-bottom:18px; color:#1f4e79; }
  .toolbar, .subbar { display:flex; gap:8px; flex-wrap:wrap; align-items:center; margin-bottom:12px; }
  button.pill { border:1px solid var(--line); background:var(--card); border-radius:999px; padding:6px 12px; cursor:pointer; }
  button.pill[aria-pressed="true"] { background:#222; color:#fff; border-color:#222; }
  button.pill.secondary[aria-pressed="true"] { background:var(--navy); border-color:var(--navy); }
  .hint { color:var(--muted); font-size:12px; margin:0 0 14px; }
  details.group { background:var(--card); border:1px solid var(--line); border-radius:12px; padding:10px 14px; margin-bottom:12px; }
  details.group > summary { cursor:pointer; font-weight:650; font-size:16px; padding:6px 0; }
  details.reply { margin:8px 0 8px 16px; border-left:3px solid #d6dee8; padding-left:12px; }
  details.reply > summary { cursor:pointer; font-weight:600; color:var(--navy); }
  .meta { color:var(--muted); font-size:12px; margin-left:8px; font-weight:400; }
  .item { padding:8px 0; border-bottom:1px solid var(--line); }
  .item:last-child { border-bottom:0; }
  .item b { font-size:12px; color:var(--muted); }
  .item.talk { border-radius:10px; padding:10px 12px; margin:8px 0; border:1px solid var(--line); border-bottom:1px solid var(--line); }
  .item.talk.head { background:#f3f7fb; border-left:3px solid var(--navy); }
  .item.talk.other { background:#fff8f1; border-left:3px solid #c05621; }
  .who { display:inline-block; margin-right:8px; color:var(--ink); font-weight:650; }
  blockquote { margin:6px 0 0; padding:0; border:0; }
  .empty { color:var(--muted); font-size:13px; padding:6px 0; }
  footer { margin-top:28px; color:var(--muted); font-size:12px; }
</style>
</head>
<body>
<main>
  <h1>01V2-分类好文段 · 事实与观点</h1>
  <p class="sub">__SUB__</p>
  <div class="note">01 原版未改。V2 按<strong>开头的人</strong>分组。箭头=回复。对侧回复人只列出回复了他的人。点某个对侧回复人后，只显示这两人按时间轴的对话，不列本人全部回复。</div>
  <div class="toolbar" id="heads"></div>
  <p class="hint" id="subhint"></p>
  <div class="subbar" id="repliers"></div>
  <div id="board"></div>
  <footer>数据来自案例数据.json。01 的 HTML/Canvas/MD/DOCX 都还在。</footer>
</main>
<script id="payload" type="application/json">__JSON__</script>
<script>
const DATA = JSON.parse(document.getElementById("payload").textContent);
let head = "全部";
let replier = "全部回复";

function speakerOf(u) { return u.speaker_canon || u.speaker || ""; }
function targetOf(u) { return u.reply_to_canon || u.reply_to || ""; }
function visibleGroups() {
  return head === "全部" ? DATA.groups : DATA.groups.filter(g => g.name === head);
}
function pairTalk(a, b) {
  return (DATA.utterances || []).filter(u => {
    const s = speakerOf(u);
    const t = targetOf(u);
    return (s === a && t === b) || (s === b && t === a);
  }).sort((x, y) => x.id.localeCompare(y.id, undefined, {numeric: true}));
}
function renderHeads() {
  const labels = ["全部", ...DATA.groups.map(g => g.name)];
  document.getElementById("heads").innerHTML = labels.map(l => {
    const g = DATA.groups.find(x => x.name === l);
    const mark = g && g.has_replies ? ` · 被回复${g.replies.length}` : "";
    const name = l === "全部" ? "全部开头的人" : l + mark;
    return `<button class="pill" data-head="${l}" aria-pressed="${head===l}">${name}</button>`;
  }).join("");
  document.querySelectorAll("[data-head]").forEach(btn => {
    btn.onclick = () => { head = btn.dataset.head; replier = "全部回复"; render(); };
  });
}
function renderRepliers() {
  const g = DATA.groups.find(x => x.name === head);
  const bar = document.getElementById("repliers");
  const hint = document.getElementById("subhint");
  if (!g || !g.has_replies) {
    bar.innerHTML = "";
    hint.textContent = (!g || head === "全部")
      ? "先点一个开头的人。只有被回复过的人才会出现对侧回复人。"
      : `${head} 没有被回复，所以没有次级选项。`;
    return;
  }
  hint.textContent = `对侧回复人：只列出回复了「${head}」的人。点其中一人，只看这两人按时间轴的对话，不列本人全部回复。`;
  const labels = ["全部回复", ...g.repliers.map(r => r.name)];
  bar.innerHTML = labels.map(l =>
    `<button class="pill secondary" data-replier="${l}" aria-pressed="${replier===l}">${l}</button>`
  ).join("");
  document.querySelectorAll("[data-replier]").forEach(btn => {
    btn.onclick = () => { replier = btn.dataset.replier; render(); };
  });
}
function itemHtml(u, opts) {
  opts = opts || {};
  const speaker = speakerOf(u);
  const target = targetOf(u);
  const targetBit = target ? `回复 ${target}` : "顶层";
  const who = opts.showSpeaker ? `<span class="who">${speaker}</span>` : "";
  const side = opts.head && speaker === opts.head ? "talk head" : (opts.showSpeaker ? "talk other" : "");
  return `<div class="item ${side}"><b>${who}${u.id} · ${targetBit} · ${u.kind} · 置信${u.ocr_confidence}</b><blockquote>${u.text}</blockquote></div>`;
}
function renderPairBoard(g) {
  const items = pairTalk(head, replier);
  const body = items.map(u => itemHtml(u, {showSpeaker: true, head})).join("")
    || `<div class="empty">这两人没有互相回复的发言。</div>`;
  document.getElementById("board").innerHTML = `<details class="group" open>
    <summary>${head} × ${replier}<span class="meta">两人对话 ${items.length} 条 · 按时间轴</span></summary>
    <p class="meta">只显示 ${head} 与 ${replier} 互相回复的发言，不列本人全部回复。对侧回复人选项仍在上方。</p>
    ${body}
  </details>`;
}
function renderBoard() {
  if (head !== "全部" && replier !== "全部回复") {
    const g = DATA.groups.find(x => x.name === head);
    if (g) { renderPairBoard(g); return; }
  }
  const groups = visibleGroups();
  document.getElementById("board").innerHTML = groups.map(g => {
    const own = g.own.map(u => itemHtml(u)).join("") || `<div class="empty">这个人在截图里没有本人发言，只有别人回复他。</div>`;
    let repliesBlock = "";
    if (g.has_replies) {
      repliesBlock = `<p class="meta">对侧回复人</p>` + g.repliers.map(r => `
        <details class="reply" open>
          <summary>回复 · ${r.name}<span class="meta">${r.items.length}</span></summary>
          ${r.items.map(u => itemHtml(u)).join("")}
        </details>`).join("");
    } else {
      repliesBlock = `<div class="empty">没有人回复他，故无次级选项。</div>`;
    }
    const open = head === "全部" || head === g.name ? " open" : "";
    return `<details class="group"${open}>
      <summary>${g.name}<span class="meta">本人 ${g.own.length} · 被回复 ${g.replies.length}</span></summary>
      <p class="meta">本人发言</p>
      ${own}
      ${repliesBlock}
    </details>`;
  }).join("");
}
function render() {
  renderHeads();
  renderRepliers();
  renderBoard();
}
render();
</script>
</body>
</html>
"""


def annotate_utterances(utterances: list[dict]) -> list[dict]:
    out = []
    for u in utterances:
        item = dict(u)
        item["speaker_canon"] = canon(u.get("speaker"))
        item["reply_to_canon"] = canon(u.get("reply_to"))
        out.append(item)
    return out


def write_v2_html(case: dict, groups: list[dict], utterances: list[dict], dest: Path):
    payload = {
        "short_title": case["short_title"],
        "snapshot_at": case["snapshot_at"],
        "groups": groups,
        "utterances": utterances,
    }
    blob = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")
    html = HTML_TEMPLATE.replace(
        "__SUB__",
        f"{case['short_title']} · {case.get('platform') or ''} · 快照 {case.get('snapshot_at') or ''}",
    ).replace("__JSON__", blob)
    dest.write_text(html, encoding="utf-8")


def main():
    case = json.loads((CASE_DIR / "案例数据.json").read_text(encoding="utf-8"))
    utterances = annotate_utterances(case["utterances"])
    groups = build_groups(utterances)
    names = collect_names(case["utterances"])
    colors = name_colors(names)

    write_v2_md(case, groups, CASE_DIR / "01V2-分类好文段-事实与观点.md")
    write_v2_docx(case, groups, CASE_DIR / "01V2-分类好文段-事实与观点.docx")
    write_v2_html(case, groups, utterances, CASE_DIR / "01V2-分类好文段-事实与观点.html")
    write_colored_ocr_docx(case, names, colors, CASE_DIR / "01V2-OCR人名分色.docx")

    print("groups", len(groups))
    for g in groups:
        print(f"  {g['name']}: own={len(g['own'])} replies={len(g['replies'])} repliers={[r['name'] for r in g['repliers']]}")
    print("names", names)
    samples = [
        ("地委最强洗地", "青山绿水"),
        ("地委最强洗地", "颜乐师傅小张"),
        ("地委最强洗地", "止水之狩"),
        ("地委最强洗地", "涂涂老师"),
    ]
    for a, b in samples:
        ids = [
            u["id"]
            for u in utterances
            if {u.get("speaker_canon"), u.get("reply_to_canon")} == {a, b}
        ]
        print(f"pair {a} x {b}: {ids}")


if __name__ == "__main__":
    main()
