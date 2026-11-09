"""Shared DOCX typography for instruction catalog and case-article files."""

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor


NAVY = RGBColor(0x1F, 0x4E, 0x79)
SLATE = RGBColor(0x5F, 0x63, 0x68)
INK = RGBColor(0x21, 0x25, 0x29)
AGENT_BLUE = RGBColor(0x1B, 0x5E, 0x8A)
YAO_BROWN = RGBColor(0x8A, 0x4B, 0x12)
ASK_GREEN = RGBColor(0x2E, 0x6B, 0x4A)
FACT = RGBColor(0x1B, 0x5E, 0x8A)
OPINION = RGBColor(0x8A, 0x4B, 0x12)
MIXED = RGBColor(0x6B, 0x4A, 0x1B)
UNSURE = RGBColor(0x7A, 0x5C, 0x00)
RULE = "D6DEE8"


def set_run_font(run, name="宋体", size=11, bold=False, color=None, italic=False):
    run.font.name = name
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color


def compact(paragraph, *, before=0, after=6, line=1.15, align=None, first_line=None):
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.line_spacing = line
    fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    if align is not None:
        paragraph.alignment = align
    if first_line is not None:
        fmt.first_line_indent = Pt(first_line)


def add_bottom_border(paragraph, color="1F4E79", sz="12"):
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color)
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def shade_cell(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_margins(cell, **sides):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for edge, twips in sides.items():
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:w"), str(twips))
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)
    tc_pr.append(tc_mar)
