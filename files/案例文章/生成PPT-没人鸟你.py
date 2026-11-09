#!/usr/bin/env python3
"""Build the 珠三角「没人鸟你」 explainer deck.

Output: files/案例文章/珠三角没人鸟你-地方财政结构解释.pptx
"""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parent
DEST = ROOT / "V1-珠三角没人鸟你-地方财政结构解释.pptx"

FONT = "Microsoft YaHei"

NAVY = RGBColor(0x1F, 0x4E, 0x79)
DEEP = RGBColor(0x14, 0x33, 0x52)
INK = RGBColor(0x1E, 0x22, 0x26)
SLATE = RGBColor(0x5C, 0x63, 0x6B)
MUTED = RGBColor(0x8A, 0x92, 0x9C)
LINE = RGBColor(0xDC, 0xE2, 0xE8)
PAPER = RGBColor(0xFB, 0xFA, 0xF7)
WASH = RGBColor(0xF1, 0xF5, 0xF9)
AMBER = RGBColor(0xB4, 0x6A, 0x0F)
AMBER_WASH = RGBColor(0xFD, 0xF6, 0xE7)
GREEN = RGBColor(0x1E, 0x6B, 0x52)
GREEN_WASH = RGBColor(0xEC, 0xF5, 0xF0)

W, H = Inches(13.333), Inches(7.5)
MARGIN = Inches(0.95)


def set_font(run, *, size=18, bold=False, color=INK, italic=False):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    rpr = run._r.get_or_add_rPr()
    ea = rpr.find(qn("a:ea"))
    if ea is None:
        ea = rpr.makeelement(qn("a:ea"), {})
        rpr.append(ea)
    ea.set("typeface", FONT)


def textbox(slide, x, y, w, h, *, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.paragraphs[0].alignment = align
    return tf


def para(tf, text, *, size=18, bold=False, color=INK, space_after=10, space_before=0,
         align=PP_ALIGN.LEFT, line=1.32, first=False, italic=False):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align
    p.space_after = Pt(space_after)
    p.space_before = Pt(space_before)
    p.line_spacing = line
    run = p.add_run()
    run.text = text
    set_font(run, size=size, bold=bold, color=color, italic=italic)
    return p


def rule(slide, x, y, w, color=LINE, h=Inches(0.014)):
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    ln.fill.solid()
    ln.fill.fore_color.rgb = color
    ln.line.fill.background()
    ln.shadow.inherit = False
    return ln


def block(slide, x, y, w, h, *, fill=WASH, line_color=None, radius=0.03):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.adjustments[0] = radius
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        shp.line.width = Pt(1)
    shp.shadow.inherit = False
    return shp


def blank(prs, *, bg=PAPER):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = bg
    return slide


def header(slide, kicker, title, *, sub=None):
    tf = textbox(slide, MARGIN, Inches(0.62), Inches(11.4), Inches(0.3))
    para(tf, kicker, size=12, bold=True, color=MUTED, space_after=0, first=True)
    tf = textbox(slide, MARGIN, Inches(0.95), Inches(11.4), Inches(0.7))
    para(tf, title, size=32, bold=True, color=DEEP, space_after=0, line=1.15, first=True)
    rule(slide, MARGIN, Inches(1.76), Inches(1.1), NAVY)
    if sub:
        tf = textbox(slide, MARGIN, Inches(1.95), Inches(11.2), Inches(0.4))
        para(tf, sub, size=15, color=SLATE, space_after=0, first=True)


def footer(slide, page, total, note="案例 01 · 为什么一部分人向往广东？ · 讨论笔记"):
    tf = textbox(slide, MARGIN, Inches(6.92), Inches(9.0), Inches(0.3))
    para(tf, note, size=10.5, color=MUTED, space_after=0, first=True)
    tf = textbox(slide, Inches(11.4), Inches(6.92), Inches(1.0), Inches(0.3))
    para(tf, f"{page} / {total}", size=10.5, color=MUTED, space_after=0, align=PP_ALIGN.RIGHT, first=True)


def bullets(slide, x, y, w, h, items, *, size=17, gap=13, bullet_color=NAVY, body_color=INK):
    tf = textbox(slide, x, y, w, h)
    first = True
    for head, body in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(gap)
        p.line_spacing = 1.34
        dot = p.add_run()
        dot.text = "▪  "
        set_font(dot, size=size, bold=True, color=bullet_color)
        if head:
            r = p.add_run()
            r.text = head + "　" if body else head
            set_font(r, size=size, bold=True, color=body_color)
        if body:
            r = p.add_run()
            r.text = body
            set_font(r, size=size, color=SLATE if head else body_color)
    return tf


# ---------------------------------------------------------------- slides

def slide_cover(prs, total):
    slide = blank(prs, bg=DEEP)
    band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(4.62), W, Inches(2.88))
    band.fill.solid()
    band.fill.fore_color.rgb = RGBColor(0x0E, 0x26, 0x3E)
    band.line.fill.background()
    band.shadow.inherit = False

    tf = textbox(slide, MARGIN, Inches(1.28), Inches(10.6), Inches(0.36))
    para(tf, "案例 01 讨论笔记", size=13, bold=True, color=RGBColor(0x8F, 0xB4, 0xD6), space_after=0, first=True)

    tf = textbox(slide, MARGIN, Inches(1.86), Inches(11.0), Inches(1.9))
    para(tf, "为什么珠三角", size=46, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), space_after=6, line=1.12, first=True)
    para(tf, "「没人鸟你」", size=46, bold=True, color=RGBColor(0xFF, 0xD9, 0x8A), space_after=0, line=1.12)

    tf = textbox(slide, MARGIN, Inches(4.98), Inches(9.6), Inches(1.3))
    para(tf, "地方财政结构，如何把「骚扰企业」变成一笔不划算的买卖",
         size=19, color=RGBColor(0xD6, 0xE2, 0xED), space_after=16, line=1.4, first=True)
    para(tf, "读解对象：知乎「为什么一部分人向往广东？」答主 欧卡老司机",
         size=13, color=RGBColor(0x8F, 0xB4, 0xD6), space_after=0, line=1.3)

    tf = textbox(slide, MARGIN, Inches(6.62), Inches(11.0), Inches(0.3))
    para(tf, "2026-11-09　·　数据与说法来源见末页",
         size=11, color=RGBColor(0x70, 0x8C, 0xA6), space_after=0, first=True)


def slide_thesis(prs, total, page):
    slide = blank(prs)
    header(slide, "一句话结论", "不是「不会发生」，而是「在边际上不划算」")

    block(slide, MARGIN, Inches(2.16), Inches(11.4), Inches(1.62), fill=NAVY, radius=0.04)
    tf = textbox(slide, MARGIN + Inches(0.5), Inches(2.44), Inches(10.4), Inches(1.1))
    para(tf, "珠三角的钱跟着企业的营业额走，不跟着个人的存量走。",
         size=25, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), space_after=8, line=1.3, first=True)
    para(tf, "把厂子逼死，等于把自己的税源掐死。", size=16, color=RGBColor(0xC9, 0xDB, 0xE9), space_after=0)

    bullets(slide, MARGIN, Inches(4.18), Inches(11.2), Inches(2.3), [
        ("先纠一个前提", "珠三角也卖地，佛山、广州卖得还很凶。差别不在「卖不卖」，在「主要从哪儿收」。"),
        ("于是有了两套行为", "税基厚的地方养鸡取蛋；税基薄的地方，缺口只能从非税、罚款、摊派里补。"),
        ("「没人鸟你」是结果", "不是官员更善良，而是破坏规则的代价太高。"),
    ], size=16.5)
    footer(slide, page, total)


def slide_correction(prs, total, page):
    slide = blank(prs)
    header(slide, "前提修正", "珠三角并不「不卖地」")

    block(slide, MARGIN, Inches(2.1), Inches(5.35), Inches(3.5), fill=AMBER_WASH, line_color=RGBColor(0xEF, 0xD9, 0xAE))
    tf = textbox(slide, MARGIN + Inches(0.42), Inches(2.4), Inches(4.5), Inches(3.0))
    para(tf, "常见的误读", size=13, bold=True, color=AMBER, space_after=12, first=True)
    para(tf, "「珠三角财政不靠卖地，所以不折腾人」", size=17, bold=True, color=INK, space_after=12, line=1.35)
    para(tf, "事实：土地财政在珠三角同样重要，广州、佛山的土地出让规模长期居前。", size=14.5, color=SLATE, space_after=0, line=1.45)

    block(slide, Inches(6.9), Inches(2.1), Inches(5.35), Inches(3.5), fill=GREEN_WASH, line_color=RGBColor(0xC4, 0xDD, 0xD0))
    tf = textbox(slide, Inches(7.32), Inches(2.4), Inches(4.5), Inches(3.0))
    para(tf, "真正成立的说法", size=13, bold=True, color=GREEN, space_after=12, first=True)
    para(tf, "「从企业流转环节收的钱，远大于从个人存量里榨的钱」", size=17, bold=True, color=INK, space_after=12, line=1.35)
    para(tf, "所以政府的最优策略是把蛋糕做大，而不是把存量刮一遍。", size=14.5, color=SLATE, space_after=0, line=1.45)

    tf = textbox(slide, MARGIN, Inches(5.95), Inches(11.3), Inches(0.6))
    para(tf, "后面所有推论，都建立在这一条修正过的前提上。", size=14, italic=True, color=MUTED, space_after=0, first=True)
    footer(slide, page, total)


def slide_taxbase(prs, total, page):
    slide = blank(prs)
    header(slide, "直接原因", "税基结构：钱从哪儿来")

    rows = [
        ("主税源", "增值税、企业所得税\n跟着企业营业额走", "税基薄，靠转移支付与非税收入补"),
        ("非税占比", "低，罚没通常只占一两个百分点", "高，个别地方能到一成以上"),
        ("财政自给", "自给率高，是净上缴省", "自给率低，靠上级转移支付"),
        ("政府最优策略", "把蛋糕做大", "从存量里榨"),
    ]
    left, top = MARGIN, Inches(2.16)
    col_w = [Inches(2.3), Inches(4.5), Inches(4.6)]
    row_h = Inches(0.88)
    head_h = Inches(0.5)

    heads = ["", "珠三角（典型）", "内地不少县市"]
    x = left
    for i, (hw, htext) in enumerate(zip(col_w, heads)):
        shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, top, hw, head_h)
        shp.fill.solid()
        shp.fill.fore_color.rgb = NAVY if i else PAPER
        shp.line.fill.background()
        shp.shadow.inherit = False
        tf = shp.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.16)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        r = p.add_run()
        r.text = htext
        set_font(r, size=13.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF) if i else INK)
        x += hw

    for ri, row in enumerate(rows):
        y = top + head_h + row_h * ri
        x = left
        for ci, (cw, cell) in enumerate(zip(col_w, row)):
            shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, cw, row_h)
            shp.fill.solid()
            shp.fill.fore_color.rgb = RGBColor(0xF4, 0xF7, 0xFA) if ci == 0 else PAPER
            shp.line.color.rgb = LINE
            shp.line.width = Pt(0.75)
            shp.shadow.inherit = False
            tf = shp.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = Inches(0.16)
            tf.margin_right = Inches(0.12)
            first = True
            for line in cell.split("\n"):
                p = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                p.space_after = Pt(1)
                p.line_spacing = 1.2
                r = p.add_run()
                r.text = line
                set_font(r, size=13 if ci == 0 else 12.5,
                         bold=(ci == 0), color=DEEP if ci == 0 else SLATE)
            x += cw

    tf = textbox(slide, MARGIN, Inches(6.22), Inches(11.3), Inches(0.5))
    para(tf, "增值税是共享税：1994 年分税制中央 75 / 地方 25，2016 年营改增后调整为 50 / 50。",
         size=12, color=MUTED, space_after=0, first=True)
    footer(slide, page, total)


def slide_vat(prs, total, page):
    slide = blank(prs)
    header(slide, "最硬的一条约束", "共享税把「逼死企业」变成自伤")

    steps = [
        ("企业多开一天", "营业额多转一圈"),
        ("地方多分一笔", "增值税 50% 留成"),
        ("税源更稳", "考核指标更好看"),
    ]
    x = MARGIN
    for i, (head, sub) in enumerate(steps):
        block(slide, x, Inches(2.24), Inches(3.42), Inches(1.62), fill=WASH)
        tf = textbox(slide, x + Inches(0.34), Inches(2.52), Inches(2.8), Inches(1.1))
        para(tf, head, size=18, bold=True, color=DEEP, space_after=6, first=True)
        para(tf, sub, size=13.5, color=SLATE, space_after=0, line=1.3)
        if i < 2:
            arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x + Inches(3.56), Inches(2.86), Inches(0.42), Inches(0.36))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = RGBColor(0xC3, 0xD2, 0xDF)
            arrow.line.fill.background()
            arrow.shadow.inherit = False
        x += Inches(3.99)

    block(slide, MARGIN, Inches(4.28), Inches(11.4), Inches(1.18), fill=AMBER_WASH, line_color=RGBColor(0xEF, 0xD9, 0xAE))
    tf = textbox(slide, MARGIN + Inches(0.42), Inches(4.52), Inches(10.5), Inches(0.8))
    para(tf, "反向推一遍：把厂子逼关门 → 营业额消失 → 地方留成归零 → 还得接手失业与维稳成本。",
         size=16.5, bold=True, color=INK, space_after=0, line=1.35, first=True)

    bullets(slide, MARGIN, Inches(5.72), Inches(11.2), Inches(1.0), [
        ("所以", "「少检查、认承诺、别乱罚」不是善意，是自利。"),
    ], size=15.5)
    footer(slide, page, total)


def slide_chain_overview(prs, total, page):
    slide = blank(prs)
    header(slide, "原因的原因", "四层，从位置一路推到行为")

    layers = [
        ("01", "全球分工的位置", "前店后厂：接单在香港，出地在珠三角", WASH, DEEP),
        ("02", "1994 分税制 + 地方竞争", "资本能用脚投票，于是必须自我约束", WASH, DEEP),
        ("03", "制度信誉的路径依赖", "外资对「随意罚款」最敏感，规则一旦立住就有溢价", WASH, DEEP),
        ("04", "净上缴省 vs 转移支付依赖", "一个自己养自己，一个缺口从非税里补", AMBER_WASH, AMBER),
    ]
    y = Inches(2.14)
    for code, head, body, fill, accent in layers:
        block(slide, MARGIN, y, Inches(11.4), Inches(0.98), fill=fill, radius=0.06)
        tf = textbox(slide, MARGIN + Inches(0.36), y + Inches(0.26), Inches(0.7), Inches(0.5))
        para(tf, code, size=20, bold=True, color=accent, space_after=0, first=True)
        tf = textbox(slide, MARGIN + Inches(1.12), y + Inches(0.16), Inches(3.6), Inches(0.7))
        para(tf, head, size=15.5, bold=True, color=INK, space_after=0, line=1.25, first=True)
        tf = textbox(slide, MARGIN + Inches(4.9), y + Inches(0.22), Inches(6.2), Inches(0.6))
        para(tf, body, size=13.5, color=SLATE, space_after=0, line=1.25, first=True)
        y += Inches(1.12)
    footer(slide, page, total)


def slide_two_col(prs, total, page, kicker, title, left_head, left_items, right_head, right_items, *, note=None):
    slide = blank(prs)
    header(slide, kicker, title)

    for idx, (head, items) in enumerate([(left_head, left_items), (right_head, right_items)]):
        x = MARGIN + Inches(5.86) * idx
        block(slide, x, Inches(2.1), Inches(5.4), Inches(4.3), fill=WASH if idx == 0 else GREEN_WASH, radius=0.04)
        tf = textbox(slide, x + Inches(0.42), Inches(2.36), Inches(4.6), Inches(0.4))
        para(tf, head, size=13, bold=True, color=NAVY if idx == 0 else GREEN, space_after=0, first=True)
        bullets(slide, x + Inches(0.42), Inches(2.84), Inches(4.6), Inches(3.3),
                items, size=14.5, gap=11)

    if note:
        tf = textbox(slide, MARGIN, Inches(6.5), Inches(11.3), Inches(0.4))
        para(tf, note, size=12, italic=True, color=MUTED, space_after=0, first=True)
    footer(slide, page, total)


def slide_tension(prs, total, page):
    slide = blank(prs)
    header(slide, "基层视角", "为什么日常只见到网格员")

    block(slide, MARGIN, Inches(2.14), Inches(11.4), Inches(1.5), fill=NAVY, radius=0.04)
    tf = textbox(slide, MARGIN + Inches(0.5), Inches(2.42), Inches(10.4), Inches(1.0))
    para(tf, "东莞、佛山的「三来一补」，厂房是村组集体建的。",
         size=22, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), space_after=8, line=1.3, first=True)
    para(tf, "村集体有租金，村民有分红 —— 基层不用靠罚款养活自己。", size=15, color=RGBColor(0xC9, 0xDB, 0xE9), space_after=0)

    bullets(slide, MARGIN, Inches(3.94), Inches(11.2), Inches(2.4), [
        ("结果", "日常接触的是「看一眼就走」的角色，而不是一排执法队。"),
        ("对照组", "税基薄、人口流出、靠转移支付的地方，缺口只能从非税里补，罚没就成了收入来源。"),
        ("换个说法", "同样一句「没人鸟你」，在这里是行政克制，在别处是无人负责。"),
    ], size=16)
    footer(slide, page, total)


def slide_limits(prs, total, page):
    slide = blank(prs)
    header(slide, "这套解释的天花板", "「不打扰」随时可以被上级叫停")

    items = [
        ("环保督察", "2017 年中央环保督察，佛山、东莞大量小厂被一刀切关停 —— 压力一来，克制立刻失效。"),
        ("拆迁与改造", "城中村改造里的强制，照样存在。"),
        ("下行周期", "近年多地非税收入占比上升，广东也在升。"),
        ("打工者一侧", "欠薪、工伤、劳动监察 —— 没人搭理的地方，出事照样没人搭理。"),
    ]
    y = Inches(2.16)
    for i, (head, body) in enumerate(items):
        block(slide, MARGIN, y, Inches(11.4), Inches(1.02), fill=AMBER_WASH if i % 2 == 0 else WASH, radius=0.06)
        tf = textbox(slide, MARGIN + Inches(0.4), y + Inches(0.2), Inches(2.3), Inches(0.6))
        para(tf, head, size=15.5, bold=True, color=AMBER if i % 2 == 0 else DEEP, space_after=0, first=True)
        tf = textbox(slide, MARGIN + Inches(2.8), y + Inches(0.22), Inches(8.2), Inches(0.62))
        para(tf, body, size=13.5, color=SLATE, space_after=0, line=1.3, first=True)
        y += Inches(1.14)
    footer(slide, page, total)


def slide_closing(prs, total, page):
    slide = blank(prs, bg=DEEP)
    tf = textbox(slide, MARGIN, Inches(1.62), Inches(11.0), Inches(0.4))
    para(tf, "收束", size=13, bold=True, color=RGBColor(0x8F, 0xB4, 0xD6), space_after=0, first=True)

    tf = textbox(slide, MARGIN, Inches(2.16), Inches(11.0), Inches(1.7))
    para(tf, "一句话可以记住的版本", size=26, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF), space_after=18, first=True)
    para(tf, "珠三角的地方财政结构，让「骚扰企业」在边际上不划算；\n不是「不会发生」。",
         size=22, color=RGBColor(0xFF, 0xD9, 0x8A), space_after=0, line=1.5)

    tf = textbox(slide, MARGIN, Inches(4.5), Inches(11.0), Inches(1.5))
    para(tf, "原文说的是行政边界感，不是道德优越。", size=15, color=RGBColor(0xD6, 0xE2, 0xED), space_after=8, line=1.4, first=True)
    para(tf, "它同时在推「躺平」叙事，文末留了微信群号 —— 读的时候把这层算进去。",
         size=15, color=RGBColor(0x8F, 0xB4, 0xD6), space_after=0, line=1.4)

    rule(slide, MARGIN, Inches(6.02), Inches(11.4), RGBColor(0x33, 0x4E, 0x68))
    tf = textbox(slide, MARGIN, Inches(6.24), Inches(11.0), Inches(0.7))
    para(tf, "材料：知乎「为什么一部分人向往广东？」· 答主欧卡老司机 · 编辑于 2026-11-03",
         size=11.5, color=RGBColor(0x70, 0x8C, 0xA6), space_after=4, first=True)
    para(tf, "评论区 79 条回复见 files/案例文章/备份2/　·　回答页 PDF 见 备份3/",
         size=11.5, color=RGBColor(0x70, 0x8C, 0xA6), space_after=0)


def build():
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H

    total = 11
    slide_cover(prs, total)
    slide_thesis(prs, total, 2)
    slide_correction(prs, total, 3)
    slide_taxbase(prs, total, 4)
    slide_vat(prs, total, 5)
    slide_chain_overview(prs, total, 6)
    slide_two_col(
        prs, total, 7,
        "第一层 · 第二层", "位置决定收入，收入决定行为",
        "全球分工：前店后厂",
        [
            ("", "香港接单、金融、航运；珠三角出地、出人、出厂房。"),
            ("", "加工贸易利润薄，但周转快、规模大。"),
            ("", "收入跟着营业额走，不跟着资产存量走 —— 共生，不是掠夺。"),
        ],
        "1994 分税制 + 地方竞争",
        [
            ("", "GDP 与税收增长进入考核，招商引资成了锦标赛。"),
            ("", "资本能用脚投票：搬去东莞，或者搬去越南。"),
            ("", "自我约束不是觉悟，是竞争压力。"),
        ],
        note="学界把这套机制叫作「市场维护型联邦主义」，说白了就是养鸡取蛋、不能杀鸡。",
    )
    slide_two_col(
        prs, total, 8,
        "第三层 · 第四层", "信誉会累积，缺口会反噬",
        "制度信誉的路径依赖",
        [
            ("", "早期靠外资，而外资对「随意罚款」最敏感。"),
            ("", "为了留住第一批港资台资，规则必须可预期。"),
            ("", "信誉一旦累积，后来者跟着受益 —— 不必假设每个官员都善良。"),
        ],
        "净上缴省 vs 转移支付依赖",
        [
            ("", "广东拿不到多少转移支付，得自己养自己。"),
            ("", "税基薄、人口流出的地方，缺口只能从非税收入补。"),
            ("", "「罚款财政」就是这么长出来的。"),
        ],
        note="这两层解释了为什么同样的制度文本，落到不同地方会长出完全不同的执法风格。",
    )
    slide_tension(prs, total, 9)
    slide_limits(prs, total, 10)
    slide_closing(prs, total, 11)

    prs.save(DEST)
    print("saved", DEST, "slides", len(prs.slides._sldIdLst))


if __name__ == "__main__":
    build()
