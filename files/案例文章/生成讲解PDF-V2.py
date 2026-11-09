#!/usr/bin/env python3
"""Render the two answer rounds as styled PDFs (V2-1 / V2-2).

Source of truth for wording: the chat answers of 2026-11-09. Also writes a DOCX
so the text stays editable.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT.parent / "指令库"))
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

BODY_FONT = "宋体"
HEAD_FONT = "微软雅黑"
MUTED = SLATE

V21_NAME = "V2-1-没人鸟你-公司与个人的找茬"
V22_NAME = "V2-2-地方财政结构-原因的原因"


def text_run(paragraph, text, *, name=BODY_FONT, size=10.5, bold=False, color=INK,
             italic=False):
    run = paragraph.add_run(text)
    set_run_font(run, name, size, bold, color, italic=italic)
    return run


def new_doc():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.85)
    section.bottom_margin = Inches(0.85)
    section.left_margin = Inches(0.95)
    section.right_margin = Inches(0.95)
    return doc


def add_page_footer(doc, label):
    footer = doc.sections[0].footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    text_run(p, f"案例 01 · 为什么一部分人向往广东？ · {label}　·　", size=8.5, color=MUTED)
    run = p.add_run()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), " PAGE ")
    run._r.addnext(fld)
    inner = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "17")
    rpr.append(sz)
    inner.append(rpr)
    t = OxmlElement("w:t")
    t.text = "1"
    inner.append(t)
    fld.append(inner)


def kicker(doc, text):
    p = doc.add_paragraph()
    compact(p, before=0, after=2, line=1.0)
    text_run(p, text, name=HEAD_FONT, size=9, bold=True, color=MUTED)


def doc_title(doc, title, sub=None):
    p = doc.add_paragraph()
    compact(p, before=2, after=4, line=1.12)
    text_run(p, title, name=HEAD_FONT, size=22, bold=True, color=NAVY)

    rule = doc.add_paragraph()
    compact(rule, before=0, after=8, line=1.0)
    add_bottom_border(rule, "1F4E79", "14")

    if sub:
        p = doc.add_paragraph()
        compact(p, before=0, after=10, line=1.35)
        text_run(p, sub, size=10.5, color=MUTED)


def h1(doc, text, n=None):
    p = doc.add_paragraph()
    compact(p, before=13, after=4, line=1.12)
    if n:
        text_run(p, f"{n}　", name=HEAD_FONT, size=12, bold=True, color=RGB(0xB4, 0x6A, 0x0F))
    text_run(p, text, name=HEAD_FONT, size=13, bold=True, color=NAVY)
    return p


def h2(doc, text):
    p = doc.add_paragraph()
    compact(p, before=8, after=2, line=1.12)
    text_run(p, text, name=HEAD_FONT, size=11, bold=True, color=INK)
    return p


def body(doc, text, *, size=10.5, after=4, color=INK, italic=False):
    p = doc.add_paragraph()
    compact(p, before=0, after=after, line=1.4)
    text_run(p, text, size=size, color=color, italic=italic)
    return p


def bullet(doc, text, *, head=None, size=10.5, indent=0.0):
    p = doc.add_paragraph()
    compact(p, before=0, after=3, line=1.38)
    p.paragraph_format.left_indent = Inches(0.22 + indent)
    p.paragraph_format.first_line_indent = Inches(-0.22)
    text_run(p, "▪  ", size=size, bold=True, color=NAVY)
    if head:
        text_run(p, head, size=size, bold=True, color=INK)
        text_run(p, "　", size=size)
    text_run(p, text, size=size, color=INK)
    return p


def quote(doc, text, *, cite=None):
    p = doc.add_paragraph()
    compact(p, before=6, after=4, line=1.4)
    p.paragraph_format.left_indent = Inches(0.3)
    p.paragraph_format.right_indent = Inches(0.15)
    add_bottom_border(p, "C9D6E2", "6")
    text_run(p, text, size=10, color=RGB(0x33, 0x3A, 0x42), italic=True)
    if cite:
        text_run(p, f"　—— {cite}", size=9, color=MUTED)


def RGB(r, g, b):
    from docx.shared import RGBColor
    return RGBColor(r, g, b)


def no_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant = OxmlElement("w:cantSplit")
    tr_pr.append(cant)


def note_box(doc, text, *, label="提醒"):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    row = table.rows[0]
    no_split(row)
    cell = row.cells[0]
    cell.width = Inches(6.35)
    shade_cell(cell, "FDF6E7")
    set_cell_margins(cell, top=120, bottom=120, left=160, right=160)
    p = cell.paragraphs[0]
    compact(p, before=0, after=0, line=1.42)
    text_run(p, f"{label}　", name=HEAD_FONT, size=9.5, bold=True, color=RGB(0xB4, 0x6A, 0x0F))
    text_run(p, text, size=10, color=INK)


def table(doc, headers, rows, widths, *, size=9.5):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    no_split(t.rows[0])
    for i, (h, w) in enumerate(zip(headers, widths)):
        cell = t.rows[0].cells[i]
        cell.width = Inches(w)
        shade_cell(cell, "1F4E79")
        set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
        p = cell.paragraphs[0]
        compact(p, before=0, after=0, line=1.2)
        text_run(p, h, name=HEAD_FONT, size=size, bold=True, color=RGB(0xFF, 0xFF, 0xFF))

    for ri, row in enumerate(rows):
        trow = t.add_row()
        no_split(trow)
        cells = trow.cells
        for ci, (val, w) in enumerate(zip(row, widths)):
            cell = cells[ci]
            cell.width = Inches(w)
            if ri % 2 == 1:
                shade_cell(cell, "F4F7FA")
            set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
            p = cell.paragraphs[0]
            compact(p, before=0, after=0, line=1.3)
            text_run(p, val, size=size, bold=(ci == 0), color=RGB(0x14, 0x33, 0x52) if ci == 0 else SLATE)
    return t


def section_break(doc):
    p = doc.add_paragraph()
    compact(p, before=10, after=2, line=1.0)
    add_bottom_border(p, "E3E8ED", "6")


# ------------------------------------------------------------------ V2-1

def build_v21(dest: Path):
    doc = new_doc()
    add_page_footer(doc, "V2-1")

    kicker(doc, "案例 01 · 讨论笔记 · V2-1")
    doc_title(
        doc,
        "「没人鸟你」到底指什么「找茬」",
        "针对公司（A）、针对个人（B），以及「几百万在珠三角得往后稍稍」（C）",
    )

    h1(doc, "先把「鸟」字说清", "一")
    body(doc, "「鸟」在这里是动词：搭理、盯上、找事。「没人鸟你」＝没人来管你，不是没人尊重你。")
    body(doc, "作者整段讲的是「政府行政水平」，所以「找茬」主要指行政与准行政的打扰，不是街坊邻里。")

    h1(doc, "A：针对公司的找茬", "二")
    body(doc, "按企业与政府打交道的阶段拆，四类：", after=6)

    h2(doc, "事前 —— 进门要盖章")
    bullet(doc, "办证、年检、资质、许可，标准不透明，得托人、得「表示」。")
    bullet(doc, "消防、环保、卫生、特种设备，各有各的前置条件，改一轮再改一轮。")

    h2(doc, "事中 —— 日常被打扰")
    bullet(doc, "运动式检查：逢年过节、重大活动、上级来视察，一轮一轮查，动不动停业整改。")
    bullet(doc, "多部门重复查同一件事：今天环保明天安监，台账要做三套。")
    bullet(doc, "摊派与「自愿」：协会费、赞助、评比达标、报刊订阅、结对帮扶。")
    bullet(doc, "被要求配合统计：产值、税收、就业数字要「好看」。")

    h2(doc, "事后 —— 罚与关")
    bullet(doc, "罚款指标化，以罚代管，选择性执法。")
    bullet(doc, "环保一刀切关停、限产、错峰生产。")
    bullet(doc, "新官不理旧账：招商时的补贴和承诺不落地。")
    bullet(doc, "政府或国企拖欠工程款。")

    h2(doc, "关系型")
    bullet(doc, "要「拜码头」、要站队；领导一换，旧账被翻。")
    bullet(doc, "越是「重点企业」越要频繁接待、陪同。")

    body(doc, "作者的原话是：广州几十人的小公司，合规经营，营业额上亿，到关门那天可能都没见过政府的人，只见过网格员。", after=3)
    body(doc, "他夸的是「执法的边界感」——不是不管，是不越界。")

    h1(doc, "B：针对个人的找茬", "三")
    h2(doc, "穷的人")
    bullet(doc, "低保、救助要自证穷，被公示、被反复核查。")
    bullet(doc, "城管对小摊小贩：驱赶、没收。")
    bullet(doc, "城中村、群租、出租屋的消防与治安清查，居住证盘查。")
    bullet(doc, "孩子上学要凑齐几证、算积分。")
    bullet(doc, "办事看脸色、要找人；被「清理低端人口」式的运动扫到。")

    h2(doc, "有点钱的人")
    bullet(doc, "被当成提款机：家乡招商、宗族修谱、各种「回报」募捐。")
    bullet(doc, "财富一露头，亲戚、地方、各种关系都来「鸟」你。")
    bullet(doc, "想做点事就得维持关系，不维持就被卡。")

    h2(doc, "很富的人")
    bullet(doc, "合规成本与监管关注都放大，历史问题容易被翻旧账。")
    bullet(doc, "跨省执法、涉案财产处置这类风险，钱越多越显眼。")
    bullet(doc, "反过来，越有钱越容易买到「保护」，代价是绑定。")

    h1(doc, "C：为什么几百万在这里「往后稍稍」", "四")
    body(doc, "你的读法基本对：几百万在有些地方是块肥肉，在珠三角排不上号。但不是因为「钱多到没人敢惹」，而是三件事把你稀释了：", after=6)
    bullet(doc, "人口基数：几千万外来务工者，你不是特殊的那个，穷也不显得穷。", head="①")
    bullet(doc, "财富密度：有钱人太多，几百万排不上号，出头也没人看。", head="②")
    bullet(doc, "汲取动机弱：地方财政不那么依赖从个体身上罚款、摊派、卖地来养活，所以没必要来「鸟」你。", head="③")
    body(doc, "所以「往后稍稍」＝既不出头也不垫底，排队都排不到你。这跟作者说的「没人鸟你才是最好的状态」是同一件事的两面。")

    section_break(doc)
    h1(doc, "一句提醒", "五")
    note_box(
        doc,
        "这句话有滤镜。同一个「没人鸟你」，对开公司的人是「不被骚扰」，"
        "对打工的人也可能是「出了事没人管」：欠薪、工伤、劳动监察，在没人搭理的地方照样没人搭理。"
        "作者是在推「躺平」叙事，文末还留了微信群号，读的时候把这层算进去。",
    )

    doc.core_properties.title = V21_NAME
    doc.save(dest)


# ------------------------------------------------------------------ V2-2

def build_v22(dest: Path):
    doc = new_doc()
    add_page_footer(doc, "V2-2")

    kicker(doc, "案例 01 · 讨论笔记 · V2-2")
    doc_title(
        doc,
        "地方财政为什么不靠罚款、摊派、卖地",
        "从税基结构，一路推到「没人鸟你」；以及这套解释的天花板",
    )

    h1(doc, "先修正一个前提", "一")
    body(doc, "不是「珠三角不卖地」——佛山、广州卖得很凶。真正不同的是钱从哪儿收：从企业的流转环节收，而不是从个人的存量、罚款、摊派里收。")

    h1(doc, "直接原因：税基结构不一样", "二")
    table(
        doc,
        ["", "珠三角（典型）", "内地不少县市"],
        [
            ("主税源", "增值税、企业所得税，跟着企业营业额走", "税基薄，靠转移支付与非税收入补"),
            ("非税占比", "低，罚没通常只占一两个百分点", "高，个别地方能到一成以上"),
            ("财政自给", "自给率高，是净上缴省", "自给率低，靠上级转移支付"),
            ("政府最优策略", "把蛋糕做大", "从存量里榨"),
        ],
        [1.15, 2.55, 2.55],
    )
    body(doc, "增值税是共享税（1994 年分税制中央 75 / 地方 25，2016 年营改增后调整为 50 / 50）。企业多开一天、多转一圈，地方就多分一笔。把厂子逼死等于把自己的税源掐死，这是最硬的一条约束。", size=9.5, after=4)

    h1(doc, "原因的原因，一层层往下", "三")

    h2(doc, "第一层：全球分工的位置")
    bullet(doc, "珠三角吃的是「前店后厂」——香港接单、金融、航运，珠三角出地、出人、出厂房。")
    bullet(doc, "做的是加工贸易，利润薄但周转快、规模大。")
    bullet(doc, "政府收入跟着营业额走，不跟着资产存量走。这是共生关系，不是掠夺关系。")

    h2(doc, "第二层：1994 分税制 + 地方竞争")
    bullet(doc, "分税制把地方推向「招商引资锦标赛」：GDP 和税收增长是考核指标。")
    bullet(doc, "资本能用脚投票——搬去东莞、搬去越南。所以地方政府必须自我约束：产权要稳、检查要少、承诺要认。")
    bullet(doc, "学界把这套机制叫「市场维护型联邦主义」，说白了就是养鸡取蛋，不能杀鸡。")

    h2(doc, "第三层：制度信誉的路径依赖")
    bullet(doc, "早期靠外资，外资对「随意罚款」最敏感。")
    bullet(doc, "为了留住第一批港资台资，地方必须建立可预期的规则。")
    bullet(doc, "这个信誉一旦累积起来，后来者跟着受益——不是每个官员都善良，而是破坏规则的代价太高。")

    h2(doc, "第四层：它是净上缴省")
    bullet(doc, "广东拿不到多少转移支付，得自己养自己，所以更在意本地企业活不活。")
    bullet(doc, "反过来，税基薄、人口流出、靠转移支付过日子的地方，缺口只能从非税收入补——罚款财政就是这么长出来的。")

    h1(doc, "基层为什么只见到「网格员」", "四")
    body(doc, "这点容易被忽略：东莞、佛山的「三来一补」是村组集体建厂房出租。村集体有租金、村民有分红，基层不用靠罚款养活自己。所以日常接触的是网格员这种「看一眼就走」的角色，而不是一排执法队。")

    h1(doc, "边界与反例", "五")
    body(doc, "这套解释有天花板：", after=6)
    bullet(doc, "2017 年中央环保督察，佛山、东莞大量小厂被一刀切关停——上级压力一来，「不打扰」立刻失效。")
    bullet(doc, "城中村改造、拆迁里的强制，照样有。")
    bullet(doc, "这几年经济下行，多地非税收入占比上升，广东也在升。")
    bullet(doc, "「没人鸟你」对打工者是双刃剑：欠薪、工伤、劳动监察，没人搭理的地方照样没人搭理。")

    note_box(
        doc,
        "准确说法是：珠三角的地方财政结构，让「骚扰企业」在边际上不划算，而不是「不会发生」。",
        label="收束",
    )

    doc.core_properties.title = V22_NAME
    doc.save(dest)


def to_pdf(docx_path: Path) -> Path:
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(docx_path.parent), str(docx_path)],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    pdf = docx_path.with_suffix(".pdf")
    if not pdf.exists():
        raise RuntimeError(f"PDF not produced: {pdf}")
    return pdf


def main():
    for name, builder in [(V21_NAME, build_v21), (V22_NAME, build_v22)]:
        docx_path = ROOT / f"{name}.docx"
        builder(docx_path)
        pdf = to_pdf(docx_path)
        print("wrote", docx_path.name, "->", pdf.name, pdf.stat().st_size, "bytes")


if __name__ == "__main__":
    main()
