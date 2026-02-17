# -*- coding: utf-8 -*-
"""
Slick Minimal template v2 - with Zilla Slab / Source Sans 3 proxies,
increased text sizes, adaptive layouts, bold lead-ins, rounded cards.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
from lxml import etree

W = Inches(10)
H = Inches(5.625)

GREEN = RGBColor(0x36, 0x87, 0x27)
GREEN_LIGHT = RGBColor(0xF7, 0xF4, 0xE4)
GREEN_MID = RGBColor(0x1D, 0xE4, 0xCA)
BLUE = RGBColor(0x38, 0x80, 0xF3)
PURPLE = RGBColor(0x5B, 0x2C, 0x8F)
COBALT = RGBColor(0x04, 0x54, 0x7C)
GOLD = RGBColor(0xD4, 0xA8, 0x43)
DARK = RGBColor(0x40, 0x3F, 0x3E)
MID = RGBColor(0x55, 0x55, 0x55)
LIGHT = RGBColor(0xE5, 0xE5, 0xE5)
OFF_WHITE = RGBColor(0xF9, 0xF7, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLORS = [GREEN, BLUE, PURPLE, COBALT, GOLD]

TITLE_FONT = "Zilla Slab"
BODY_FONT = "Source Sans 3"

LM = Inches(0.9)
ACC_W = Inches(0.25)
CW = Inches(8.6)


# -- helpers --

def _add_rect(slide, x, y, w, h, fill_color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.fill.solid(); shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background(); return shape

def _add_rounded_rect(slide, x, y, w, h, fill_color):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shape.fill.solid(); shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    pg = shape._element.find('.//' + qn('a:prstGeom'))
    if pg is not None:
        av = pg.find(qn('a:avLst'))
        if av is None: av = etree.SubElement(pg, qn('a:avLst'))
        for g in av.findall(qn('a:gd')): av.remove(g)
        g = etree.SubElement(av, qn('a:gd')); g.set('name','adj'); g.set('fmla','val 5000')
    return shape

def _add_oval(slide, x, y, w, h, fill_color):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
    shape.fill.solid(); shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background(); return shape

def _add_text_box(slide, x, y, w, h, text, font_name=None, font_size=12,
                  color=None, bold=False, italic=False, align=PP_ALIGN.LEFT,
                  valign=MSO_ANCHOR.TOP, line_spacing=None):
    font_name = font_name or BODY_FONT; color = color or DARK
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = text; p.font.name = font_name
    p.font.size = Pt(font_size); p.font.color.rgb = color
    p.font.bold = bold; p.font.italic = italic; p.alignment = align
    if line_spacing: p.line_spacing = Pt(line_spacing)
    bp = tf._txBody.find(qn('a:bodyPr'))
    if bp is not None:
        bp.set('anchor', {MSO_ANCHOR.TOP:'t', MSO_ANCHOR.MIDDLE:'ctr', MSO_ANCHOR.BOTTOM:'b'}.get(valign,'t'))
    return txBox

def _accent(slide):
    _add_rect(slide, Inches(0), Inches(0), ACC_W, H, GREEN)

def _slide_title(slide, title, y=None, size=28):
    """Title at top with green rule. Moved UP for more content space."""
    _add_text_box(slide, LM, Inches(0.15), CW, Inches(0.65), title,
                  TITLE_FONT, size, DARK, bold=True, valign=MSO_ANCHOR.BOTTOM)
    _add_rect(slide, LM, Inches(0.85), Inches(2.5), Inches(0.04), GREEN)

# Content starts at 1.05" (below title + rule + breathing room)
CONTENT_TOP = Inches(1.05)

def _find_split(text):
    """Find where to split text into bold lead-in + regular continuation."""
    for delim in [' — ', ' – ']:
        idx = text.find(delim)
        if idx > 0: return text[:idx + len(delim)], text[idx + len(delim):]
    idx = text.find(': ')
    if idx > 0: return text[:idx + 1], text[idx + 2:]
    words = text.split(); char_count = 0
    for i, w in enumerate(words):
        char_count += len(w) + 1
        if i >= 3:
            ci = text.find(',', char_count - 1)
            if ci > 0 and ci < len(text) * 0.55: return text[:ci + 1], text[ci + 2:]
    if len(words) > 10:
        n = min(8, len(words) // 2)
        pos = sum(len(words[i]) + 1 for i in range(n))
        return text[:pos].rstrip(), text[pos:]
    return text, ""

def _add_split_text(slide, x, y, w, h, text, font_name, font_size, color, line_spacing=None):
    """Text with bold lead-in clause."""
    tb = slide.shapes.add_textbox(x, y, w, h); tf = tb.text_frame; tf.word_wrap = True
    bold_part, rest_part = _find_split(text)
    p = tf.paragraphs[0]
    r1 = p.add_run(); r1.text = bold_part + (" " if rest_part else "")
    r1.font.name = font_name; r1.font.size = Pt(font_size)
    r1.font.color.rgb = color; r1.font.bold = True
    if rest_part:
        r2 = p.add_run(); r2.text = rest_part
        r2.font.name = font_name; r2.font.size = Pt(font_size)
        r2.font.color.rgb = color; r2.font.bold = False
    p.alignment = PP_ALIGN.LEFT
    if line_spacing: p.line_spacing = Pt(line_spacing)
    bp = tf._txBody.find(qn('a:bodyPr'))
    if bp is not None: bp.set('anchor', 'ctr')
    return tb

def _adaptive_layout(n, content_top, slide_h, bottom_margin=Inches(0.35)):
    """Calculate card positions - fills available space, biased slightly up."""
    avail = slide_h - content_top - bottom_margin
    gap = Inches(0.12)
    rowH = int((avail - (n - 1) * gap) / n)
    max_row = Inches(1.15)
    if rowH > max_row: rowH = max_row
    total = n * rowH + (n - 1) * gap
    remaining = avail - total
    startY = Emu(content_top + int(remaining * 0.3))
    fs = 16 if n <= 3 else 15
    ls = 20 if n <= 3 else 19
    return startY, rowH, gap, fs, ls


# -- slide builders --

def build_title(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_rect(slide, Inches(0), Inches(0), ACC_W, H, GREEN)
    _add_text_box(slide, LM, Inches(1.0), CW, Inches(1.5), c.get("title", "Title"),
                  TITLE_FONT, 40, DARK, bold=True, valign=MSO_ANCHOR.BOTTOM)
    _add_rect(slide, LM, Inches(2.65), Inches(2.5), Inches(0.04), GREEN)
    if c.get("subtitle"):
        _add_text_box(slide, LM, Inches(2.85), CW, Inches(0.6), c["subtitle"],
                      BODY_FONT, 16, MID)
    if c.get("author"):
        _add_text_box(slide, LM, Inches(4.1), CW, Inches(0.35), c["author"],
                      BODY_FONT, 12, MID, bold=True)
    if c.get("date"):
        _add_text_box(slide, LM, Inches(4.45), CW, Inches(0.3), c["date"],
                      BODY_FONT, 11, MID)


def build_in_brief(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide)
    _slide_title(slide, c.get("title", "In Brief"))
    bullets = c.get("bullets", [])
    n = len(bullets)
    startY, rowH, gap, fs, ls = _adaptive_layout(n, CONTENT_TOP, H)
    
    for i, b in enumerate(bullets):
        y = Emu(startY + i * (rowH + gap))
        col = COLORS[i % len(COLORS)]
        _add_rounded_rect(slide, LM, y, CW, rowH, OFF_WHITE)
        _add_rect(slide, LM, y, Inches(0.10), rowH, col)
        cs = Inches(0.40); cx = Emu(LM + Inches(0.22)); cy = Emu(y + (rowH - cs) // 2)
        _add_oval(slide, cx, cy, cs, cs, col)
        _add_text_box(slide, cx, cy, cs, cs, str(i+1), TITLE_FONT, 14, WHITE,
                      bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        _add_split_text(slide, Emu(LM + Inches(0.8)), y, Emu(CW - Inches(0.95)), rowH,
                        b, BODY_FONT, fs, DARK, line_spacing=ls)


def build_section_divider(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; fill = bg.fill; fill.solid(); fill.fore_color.rgb = GREEN
    if c.get("sectionNumber"):
        _add_text_box(slide, Inches(0.8), Inches(1.2), Inches(2), Inches(0.8),
                      f"0{c['sectionNumber']}", TITLE_FONT, 48, GREEN_MID, bold=True,
                      valign=MSO_ANCHOR.BOTTOM)
    _add_text_box(slide, Inches(0.8), Inches(2.1), Inches(8), Inches(1.0),
                  c.get("title", "Section"), TITLE_FONT, 36, WHITE, bold=True,
                  valign=MSO_ANCHOR.MIDDLE)
    _add_rect(slide, Inches(0.8), Inches(3.2), Inches(2.0), Inches(0.04), GREEN_MID)
    if c.get("subtitle"):
        _add_text_box(slide, Inches(0.8), Inches(3.4), Inches(8), Inches(0.5),
                      c["subtitle"], BODY_FONT, 15, GREEN_LIGHT)


def build_stat_callout(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Key Metric"), size=22)
    _add_text_box(slide, LM, Inches(1.3), CW, Inches(1.8), c.get("stat", "—"),
                  TITLE_FONT, 80, GREEN, bold=True, align=PP_ALIGN.CENTER,
                  valign=MSO_ANCHOR.MIDDLE)
    if c.get("headline"):
        _add_text_box(slide, Inches(1.5), Inches(3.1), Inches(7), Inches(0.6),
                      c["headline"], BODY_FONT, 17, DARK, bold=True, align=PP_ALIGN.CENTER)
    if c.get("detail"):
        _add_text_box(slide, Inches(1.5), Inches(3.75), Inches(7), Inches(0.8),
                      c["detail"], BODY_FONT, 13, MID, align=PP_ALIGN.CENTER, line_spacing=17)
    if c.get("source"):
        _add_text_box(slide, Inches(0.5), Inches(4.85), Inches(9), Inches(0.3),
                      c["source"], BODY_FONT, 9, MID, italic=True, align=PP_ALIGN.CENTER)


def build_quote(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "In Their Words"), size=22)
    _add_text_box(slide, Emu(LM - Inches(0.1)), Inches(1.2), Inches(0.8), Inches(0.8),
                  "\u201C", "Georgia", 64, GREEN_LIGHT, bold=True)
    _add_text_box(slide, Emu(LM + Inches(0.5)), Inches(1.5), Inches(7.5), Inches(2.0),
                  c.get("quote", ""), BODY_FONT, 18, DARK, italic=True,
                  valign=MSO_ANCHOR.MIDDLE, line_spacing=25)
    _add_rect(slide, Emu(LM + Inches(0.5)), Inches(3.7), Inches(1.5), Inches(0.04), GREEN)
    if c.get("attribution"):
        _add_text_box(slide, Emu(LM + Inches(0.5)), Inches(3.85), Inches(7), Inches(0.35),
                      c["attribution"], BODY_FONT, 13, MID)
    if c.get("context"):
        _add_text_box(slide, Emu(LM + Inches(0.5)), Inches(4.2), Inches(7), Inches(0.35),
                      c["context"], BODY_FONT, 11, MID, italic=True)


def build_comparison(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Comparison"), size=24)
    cardY = CONTENT_TOP; cardH = Inches(3.6)
    # Left card
    _add_rounded_rect(slide, LM, cardY, Inches(4.0), cardH, OFF_WHITE)
    _add_rect(slide, LM, cardY, Inches(0.08), cardH, COBALT)
    _add_text_box(slide, Emu(LM + Inches(0.2)), Emu(cardY + Inches(0.15)),
                  Inches(3.6), Inches(0.4), c.get("leftLabel", "Before"),
                  TITLE_FONT, 15, COBALT, bold=True)
    for i, item in enumerate(c.get("leftItems", [])):
        _add_text_box(slide, Emu(LM + Inches(0.2)), Emu(cardY + Inches(0.65) + i * Inches(0.6)),
                      Inches(3.6), Inches(0.55), item, BODY_FONT, 13, DARK,
                      valign=MSO_ANCHOR.MIDDLE)
    # Divider
    _add_rect(slide, Inches(5.05), Emu(cardY + Inches(0.2)), Inches(0.03), Emu(cardH - Inches(0.4)), LIGHT)
    # Right card
    _add_rounded_rect(slide, Inches(5.25), cardY, Inches(4.25), cardH, OFF_WHITE)
    _add_rect(slide, Inches(5.25), cardY, Inches(0.08), cardH, GREEN)
    _add_text_box(slide, Inches(5.5), Emu(cardY + Inches(0.15)),
                  Inches(3.8), Inches(0.4), c.get("rightLabel", "After"),
                  TITLE_FONT, 15, GREEN, bold=True)
    for i, item in enumerate(c.get("rightItems", [])):
        _add_text_box(slide, Inches(5.5), Emu(cardY + Inches(0.65) + i * Inches(0.6)),
                      Inches(3.85), Inches(0.55), item, BODY_FONT, 13, DARK,
                      valign=MSO_ANCHOR.MIDDLE)


def build_text_graph(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Title"), size=24)
    texts = c.get("text", []); texts = [texts] if not isinstance(texts, list) else texts
    for i, t in enumerate(texts):
        _add_text_box(slide, LM, Emu(CONTENT_TOP + i * Inches(1.2)),
                      Inches(4.0), Inches(1.1), t, BODY_FONT, 13, DARK, line_spacing=18)
    _add_rect(slide, Inches(5.1), CONTENT_TOP, Inches(0.03), Inches(3.5), LIGHT)
    chart_data_raw = c.get("chartData", [{"name": "S1", "labels": ["A","B","C"], "values": [25,45,30]}])
    from pptx.chart.data import CategoryChartData
    chart_data = CategoryChartData()
    if chart_data_raw:
        cd = chart_data_raw[0]
        chart_data.categories = cd.get("labels", ["A","B","C"])
        chart_data.add_series(cd.get("name","Series 1"), cd.get("values",[25,45,30]))
    ct = {"line": XL_CHART_TYPE.LINE, "pie": XL_CHART_TYPE.PIE}.get(c.get("chartType","bar"), XL_CHART_TYPE.COLUMN_CLUSTERED)
    slide.shapes.add_chart(ct, Inches(5.3), CONTENT_TOP, Inches(4.2), Inches(3.8), chart_data)
    if c.get("note"):
        _add_text_box(slide, Inches(5.3), Inches(5.0), Inches(4.2), Inches(0.3),
                      c["note"], BODY_FONT, 8, MID, italic=True)


def build_process_flow(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Process"), size=24)
    steps = c.get("steps", []); count = min(len(steps), 5)
    if count == 0: return
    total_w = 8.6; arrow_w = 0.3
    step_w = (total_w - (count - 1) * arrow_w) / count
    startX = 0.9; stepY = 1.05; stepH = 3.9  # taller cards, start higher
    for i, step in enumerate(steps[:count]):
        x = Inches(startX + i * (step_w + arrow_w)); y = Inches(stepY)
        sw = Inches(step_w); sh = Inches(stepH)
        _add_rounded_rect(slide, x, y, sw, sh, OFF_WHITE)
        _add_rect(slide, x, y, Inches(0.06), sh, COLORS[i % len(COLORS)])
        _add_text_box(slide, Emu(x + Inches(0.15)), Emu(y + Inches(0.15)),
                      Inches(0.35), Inches(0.35), str(i + 1),
                      TITLE_FONT, 16, COLORS[i % len(COLORS)], bold=True)
        _add_text_box(slide, Emu(x + Inches(0.15)), Emu(y + Inches(0.55)),
                      Emu(sw - Inches(0.3)), Inches(0.65),
                      step.get("title", ""), BODY_FONT, 13, DARK, bold=True)
        if step.get("detail"):
            _add_text_box(slide, Emu(x + Inches(0.15)), Emu(y + Inches(1.25)),
                          Emu(sw - Inches(0.3)), Inches(2.3),
                          step["detail"], BODY_FONT, 11, MID, line_spacing=15)
        if i < count - 1:
            _add_text_box(slide, Emu(x + sw + Inches(0.02)), Emu(y + Inches(1.5)),
                          Emu(Inches(arrow_w) - Inches(0.04)), Inches(0.4),
                          "\u2192", BODY_FONT, 16, LIGHT, align=PP_ALIGN.CENTER,
                          valign=MSO_ANCHOR.MIDDLE)


def build_matrix(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Framework"), size=24)
    quads = c.get("quadrants", [{},{},{},{}])
    accents = [GREEN, BLUE, PURPLE, COBALT]
    qW = Inches(4.05); qH = Inches(1.85); gap = Inches(0.2)
    sX = LM; sY = CONTENT_TOP
    for i, q in enumerate(quads[:4]):
        col = i % 2; row = i // 2
        x = Emu(sX + col * (qW + gap)); y = Emu(sY + row * (qH + gap))
        _add_rounded_rect(slide, x, y, qW, qH, OFF_WHITE)
        _add_rect(slide, x, y, Inches(0.06), qH, accents[i])
        _add_text_box(slide, Emu(x + Inches(0.2)), Emu(y + Inches(0.12)),
                      Emu(qW - Inches(0.4)), Inches(0.35),
                      q.get("label", ""), TITLE_FONT, 12, accents[i], bold=True)
        _add_text_box(slide, Emu(x + Inches(0.2)), Emu(y + Inches(0.5)),
                      Emu(qW - Inches(0.4)), Emu(qH - Inches(0.65)),
                      q.get("detail", ""), BODY_FONT, 11, DARK, line_spacing=15)


def build_methods(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Approach"), size=24)
    fields = c.get("fields", []); n = len(fields)
    avail = H - CONTENT_TOP - Inches(0.3)
    rowH = min(int(avail / n), Inches(0.82))
    gap = Inches(0.06) if n > 4 else Inches(0.1)
    total = n * rowH + (n - 1) * gap
    startY = Emu(CONTENT_TOP + int((avail - total) * 0.3))
    for i, f in enumerate(fields):
        y = Emu(startY + i * (rowH + gap))
        col = COLORS[i % len(COLORS)]
        _add_rounded_rect(slide, LM, y, CW, rowH, OFF_WHITE)
        _add_rect(slide, LM, y, Inches(0.10), rowH, col)
        _add_text_box(slide, Emu(LM + Inches(0.22)), y, Inches(1.8), rowH,
                      f.get("label", ""), BODY_FONT, 13, col, bold=True,
                      valign=MSO_ANCHOR.MIDDLE)
        _add_text_box(slide, Emu(LM + Inches(2.1)), y, Inches(6.3), rowH,
                      f.get("value", ""), BODY_FONT, 13, DARK,
                      valign=MSO_ANCHOR.MIDDLE)


def build_hypotheses(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Hypotheses"), size=24)
    hyps = c.get("hypotheses", []); n = len(hyps)
    avail = H - CONTENT_TOP - Inches(0.3)
    rowH = min(int(avail / n), Inches(0.72))
    gap = Inches(0.08)
    total = n * rowH + (n - 1) * gap
    startY = Emu(CONTENT_TOP + int((avail - total) * 0.3))
    for i, h in enumerate(hyps):
        y = Emu(startY + i * (rowH + gap))
        col = COLORS[i % len(COLORS)]
        _add_rounded_rect(slide, LM, y, CW, rowH, OFF_WHITE)
        _add_rect(slide, LM, y, Inches(0.10), rowH, col)
        cs = Inches(0.36); cx = Emu(LM + Inches(0.18)); cy = Emu(y + (rowH - cs) // 2)
        _add_oval(slide, cx, cy, cs, cs, col)
        _add_text_box(slide, cx, cy, cs, cs, f"H{i+1}", TITLE_FONT, 11, WHITE,
                      bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        _add_text_box(slide, Emu(LM + Inches(0.72)), y, Inches(5.95), rowH,
                      h.get("text", ""), BODY_FONT, 12, DARK,
                      valign=MSO_ANCHOR.MIDDLE)
        if h.get("status"):
            sc = GREEN if h["status"] == "Confirmed" else (COBALT if h["status"] == "Rejected" else MID)
            _add_text_box(slide, Inches(7.8), y, Inches(1.2), rowH,
                          h["status"], BODY_FONT, 10, sc, bold=True,
                          align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)


def build_wsn_dense(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Key Finding"), size=24)
    cols = [("What", GREEN, c.get("what", {})),
            ("So What", BLUE, c.get("soWhat", {})),
            ("Now What", PURPLE, c.get("nowWhat", {}))]
    colW = Inches(2.75); gap = Inches(0.2); cardH = Inches(3.8)
    startX = LM; startY = CONTENT_TOP
    for i, (label, color, data) in enumerate(cols):
        x = Emu(startX + i * (colW + gap))
        _add_rounded_rect(slide, x, startY, colW, cardH, OFF_WHITE)
        _add_rect(slide, x, startY, Inches(0.06), cardH, color)
        _add_text_box(slide, Emu(x + Inches(0.2)), Emu(startY + Inches(0.15)),
                      Emu(colW - Inches(0.4)), Inches(0.35),
                      label, TITLE_FONT, 14, color, bold=True)
        _add_text_box(slide, Emu(x + Inches(0.2)), Emu(startY + Inches(0.55)),
                      Emu(colW - Inches(0.4)), Inches(1.1),
                      data.get("headline", ""), BODY_FONT, 12, DARK, bold=True,
                      line_spacing=16)
        if data.get("detail"):
            _add_text_box(slide, Emu(x + Inches(0.2)), Emu(startY + Inches(1.7)),
                          Emu(colW - Inches(0.4)), Inches(1.9),
                          data["detail"], BODY_FONT, 11, MID, line_spacing=14)


def build_wsn_reveal(prs, c):
    """Builds 3 progressive slides: What -> So What -> Now What.
    Enhanced with step indicators, connecting arrows, and visual progression."""
    WSN_COLORS = [GREEN, BLUE, PURPLE]
    WSN_LABELS = ["What We Found", "So What", "Now What"]

    def _draw_step_indicator(slide, x, y, num, label, color, active=True):
        """Circle with number + label below."""
        cs = Inches(0.42)
        fill = color if active else LIGHT
        _add_oval(slide, x, y, cs, cs, fill)
        tc = WHITE if active else MID
        _add_text_box(slide, x, y, cs, cs, str(num), TITLE_FONT, 15, tc,
                      bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        lw = Inches(1.2)
        _add_text_box(slide, Emu(x - (lw - cs) // 2), Emu(y + cs + Inches(0.04)),
                      lw, Inches(0.25), label, BODY_FONT, 9, color if active else MID,
                      bold=active, align=PP_ALIGN.CENTER)

    def _draw_connector(slide, x1, x2, y):
        """Dotted line between step indicators."""
        _add_rect(slide, Emu(x1), y, Emu(x2 - x1), Inches(0.025), LIGHT)

    def _draw_step_bar(slide, active_count):
        """Draw the 1-2-3 step indicator bar at the top."""
        bar_y = Emu(CONTENT_TOP + Inches(0.05))
        spacing = Inches(2.5)
        start_x = Emu(LM + Inches(1.5))
        cs = Inches(0.42)
        for i in range(3):
            sx = Emu(start_x + i * spacing)
            active = i < active_count
            _draw_step_indicator(slide, sx, bar_y, i + 1, WSN_LABELS[i],
                                 WSN_COLORS[i], active)
            if i < 2:
                conn_x1 = Emu(sx + cs + Inches(0.08))
                conn_x2 = Emu(start_x + (i + 1) * spacing - Inches(0.08))
                _draw_connector(slide, conn_x1, conn_x2, Emu(bar_y + cs // 2))

    def _draw_zone(slide, x, w, color, data, emphasis=False):
        """Draw a content zone card."""
        card_y = Emu(CONTENT_TOP + Inches(0.95))
        card_h = Inches(3.7)
        _add_rounded_rect(slide, x, card_y, w, card_h, OFF_WHITE)
        _add_rect(slide, x, card_y, Inches(0.10), card_h, color)
        fs_head = 14 if emphasis else 13
        fs_detail = 12 if emphasis else 11
        _add_text_box(slide, Emu(x + Inches(0.25)), Emu(card_y + Inches(0.2)),
                      Emu(w - Inches(0.5)), Inches(0.85),
                      data.get("headline", ""), BODY_FONT, fs_head, DARK, bold=True,
                      line_spacing=18 if emphasis else 17)
        if data.get("detail"):
            _add_text_box(slide, Emu(x + Inches(0.25)), Emu(card_y + Inches(1.15)),
                          Emu(w - Inches(0.5)), Inches(2.2),
                          data["detail"], BODY_FONT, fs_detail, MID, line_spacing=16)

    # Slide 1: What only - single wide card
    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(s1); _slide_title(s1, c.get("title", "Key Finding"), size=24)
    _draw_step_bar(s1, 1)
    _draw_zone(s1, LM, Inches(5.5), GREEN, c.get("what", {}), emphasis=True)

    # Slide 2: What + So What - two cards side by side
    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(s2); _slide_title(s2, c.get("title", "Key Finding"), size=24)
    _draw_step_bar(s2, 2)
    _draw_zone(s2, LM, Inches(4.05), GREEN, c.get("what", {}))
    _draw_zone(s2, Inches(5.2), Inches(4.3), BLUE, c.get("soWhat", {}), emphasis=True)

    # Slide 3: All three - condensed top two + full-width Now What
    s3 = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(s3); _slide_title(s3, c.get("title", "Key Finding"), size=24)
    _draw_step_bar(s3, 3)
    # Condensed What + So What
    cond_y = Emu(CONTENT_TOP + Inches(0.95)); cond_h = Inches(1.75)
    _add_rounded_rect(s3, LM, cond_y, Inches(4.05), cond_h, OFF_WHITE)
    _add_rect(s3, LM, cond_y, Inches(0.08), cond_h, GREEN)
    _add_text_box(s3, Emu(LM + Inches(0.2)), Emu(cond_y + Inches(0.12)),
                  Inches(3.6), Inches(0.5),
                  c.get("what", {}).get("headline", ""), BODY_FONT, 11, DARK, bold=True)
    _add_text_box(s3, Emu(LM + Inches(0.2)), Emu(cond_y + Inches(0.65)),
                  Inches(3.6), Inches(0.9),
                  c.get("what", {}).get("detail", ""), BODY_FONT, 10, MID)

    _add_rounded_rect(s3, Inches(5.2), cond_y, Inches(4.3), cond_h, OFF_WHITE)
    _add_rect(s3, Inches(5.2), cond_y, Inches(0.08), cond_h, BLUE)
    _add_text_box(s3, Inches(5.45), Emu(cond_y + Inches(0.12)),
                  Inches(3.8), Inches(0.5),
                  c.get("soWhat", {}).get("headline", ""), BODY_FONT, 11, DARK, bold=True)
    _add_text_box(s3, Inches(5.45), Emu(cond_y + Inches(0.65)),
                  Inches(3.8), Inches(0.9),
                  c.get("soWhat", {}).get("detail", ""), BODY_FONT, 10, MID)

    # Now What - full width, emphasized
    nwY = Emu(cond_y + cond_h + Inches(0.15)); nwH = Inches(1.65)
    _add_rounded_rect(s3, LM, nwY, CW, nwH, OFF_WHITE)
    _add_rect(s3, LM, nwY, Inches(0.10), nwH, PURPLE)
    nw = c.get("nowWhat", {})
    _add_text_box(s3, Emu(LM + Inches(0.25)), Emu(nwY + Inches(0.15)),
                  Emu(CW - Inches(0.5)), Inches(0.6),
                  nw.get("headline", ""), BODY_FONT, 16, DARK, bold=True)
    if nw.get("detail"):
        _add_text_box(s3, Emu(LM + Inches(0.25)), Emu(nwY + Inches(0.8)),
                      Emu(CW - Inches(0.5)), Inches(0.65),
                      nw["detail"], BODY_FONT, 12, MID)

def build_findings_recs(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Findings & Recommendations"), size=24)
    items = c.get("items", []); n = min(len(items), 5)
    avail = H - CONTENT_TOP - Inches(0.3)
    rowH = min(int((avail - (n - 1) * Inches(0.12)) / n), Inches(0.85))
    gap = Inches(0.12)
    total = n * rowH + (n - 1) * gap
    startY = Emu(CONTENT_TOP + int((avail - total) * 0.3))
    for i, item in enumerate(items[:5]):
        y = Emu(startY + i * (rowH + gap))
        col = COLORS[i % len(COLORS)]
        _add_rounded_rect(slide, LM, y, Inches(3.9), rowH, OFF_WHITE)
        _add_rect(slide, LM, y, Inches(0.10), rowH, col)
        _add_text_box(slide, Emu(LM + Inches(0.2)), y, Inches(3.5), rowH,
                      item.get("finding", ""), BODY_FONT, 12, DARK,
                      valign=MSO_ANCHOR.MIDDLE)
        # Arrow in colored circle
        acs = Inches(0.30); acx = Emu(Inches(4.97)); acy = Emu(y + (rowH - acs) // 2)
        _add_oval(slide, acx, acy, acs, acs, col)
        _add_text_box(slide, acx, acy, acs, acs,
                      "\u2192", BODY_FONT, 13, WHITE, align=PP_ALIGN.CENTER,
                      valign=MSO_ANCHOR.MIDDLE)
        _add_rounded_rect(slide, Inches(5.4), y, Inches(4.1), rowH, OFF_WHITE)
        _add_rect(slide, Inches(5.4), y, Inches(0.10), rowH, col)
        _add_text_box(slide, Inches(5.6), y, Inches(3.75), rowH,
                      item.get("recommendation", ""), BODY_FONT, 12, DARK,
                      valign=MSO_ANCHOR.MIDDLE)


def build_findings_recs_dense(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Complete Findings"), size=22)
    items = c.get("items", []); n = min(len(items), 8)
    avail = H - CONTENT_TOP - Inches(0.2)
    rowH = min(int((avail - (n - 1) * Inches(0.06)) / n), Inches(0.48))
    gap = Inches(0.06)
    total = n * rowH + (n - 1) * gap
    startY = Emu(CONTENT_TOP + int((avail - total) * 0.3))
    for i, item in enumerate(items[:8]):
        y = Emu(startY + i * (rowH + gap))
        bg = OFF_WHITE if i % 2 == 0 else WHITE
        _add_rect(slide, LM, y, Inches(4.1), rowH, bg)
        _add_rect(slide, LM, y, Inches(0.04), rowH, GREEN)
        _add_text_box(slide, Emu(LM + Inches(0.15)), y, Inches(3.85), rowH,
                      item.get("finding", ""), BODY_FONT, 10, DARK,
                      valign=MSO_ANCHOR.MIDDLE)
        _add_rect(slide, Inches(5.15), y, Inches(4.35), rowH, bg)
        _add_rect(slide, Inches(5.15), y, Inches(0.04), rowH, BLUE)
        _add_text_box(slide, Inches(5.3), y, Inches(4.1), rowH,
                      item.get("recommendation", ""), BODY_FONT, 10, DARK,
                      valign=MSO_ANCHOR.MIDDLE)


def build_open_questions(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Open Questions"), size=26)
    questions = c.get("questions", [])
    cardW = Inches(4.1); cardH = Inches(1.7); gX = Inches(0.4); gY = Inches(0.2)
    gridY = CONTENT_TOP
    for i, question in enumerate(questions[:4]):
        col = i % 2; row = i // 2
        x = Emu(LM + col * (cardW + gX)); y = Emu(gridY + row * (cardH + gY))
        _add_rounded_rect(slide, x, y, cardW, cardH, OFF_WHITE)
        _add_rect(slide, x, y, Inches(0.06), cardH, COLORS[i % len(COLORS)])
        _add_oval(slide, Emu(x + Inches(0.15)), Emu(y + Inches(0.12)),
                  Inches(0.35), Inches(0.35), COLORS[i % len(COLORS)])
        _add_text_box(slide, Emu(x + Inches(0.15)), Emu(y + Inches(0.12)),
                      Inches(0.35), Inches(0.35), str(i + 1),
                      TITLE_FONT, 14, WHITE, bold=True, align=PP_ALIGN.CENTER,
                      valign=MSO_ANCHOR.MIDDLE)
        _add_text_box(slide, Emu(x + Inches(0.2)), Emu(y + Inches(0.55)),
                      Emu(cardW - Inches(0.4)), Inches(1.0),
                      question, BODY_FONT, 13, DARK, line_spacing=17)


def build_agenda(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _accent(slide); _slide_title(slide, c.get("title", "Agenda"), size=26)
    items = c.get("items", []); n = len(items)
    avail = H - CONTENT_TOP - Inches(0.3)
    rowH = min(int(avail / n), Inches(0.72))
    gap = Inches(0.08)
    total = n * rowH + (n - 1) * gap
    startY = Emu(CONTENT_TOP + int((avail - total) * 0.3))
    for i, item in enumerate(items):
        y = Emu(startY + i * (rowH + gap))
        col = COLORS[i % len(COLORS)]
        _add_rounded_rect(slide, LM, y, CW, rowH, OFF_WHITE)
        _add_rect(slide, LM, y, Inches(0.10), rowH, col)
        cs = Inches(0.40); cx = Emu(LM + Inches(0.22)); cy = Emu(y + (rowH - cs) // 2)
        _add_oval(slide, cx, cy, cs, cs, col)
        _add_text_box(slide, cx, cy, cs, cs, str(i + 1), TITLE_FONT, 14, WHITE,
                      bold=True, align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        title_text = item if isinstance(item, str) else item.get("title", "")
        _add_text_box(slide, Emu(LM + Inches(0.8)), y, Inches(5.5), rowH,
                      title_text, BODY_FONT, 15, DARK, bold=True,
                      valign=MSO_ANCHOR.MIDDLE)
        if isinstance(item, dict) and item.get("detail"):
            _add_text_box(slide, Inches(7.5), y, Inches(1.8), rowH,
                          item["detail"], BODY_FONT, 12, MID,
                          align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.MIDDLE)


def build_progressive_reveal(prs, c):
    takeaways = c.get("takeaways", [])
    for n in range(min(len(takeaways), 5)):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _accent(slide); _slide_title(slide, c.get("title", "Building the Picture"))
        cur = takeaways[n]
        _add_rounded_rect(slide, LM, CONTENT_TOP, CW, Inches(2.4), OFF_WHITE)
        _add_rect(slide, LM, CONTENT_TOP, Inches(0.06), Inches(2.4), GREEN)
        _add_text_box(slide, Emu(LM + Inches(0.2)), Emu(CONTENT_TOP + Inches(0.1)),
                      Emu(CW - Inches(0.4)), Inches(0.6),
                      cur.get("headline", ""), BODY_FONT, 16, DARK, bold=True)
        if cur.get("detail"):
            _add_text_box(slide, Emu(LM + Inches(0.2)), Emu(CONTENT_TOP + Inches(0.75)),
                          Emu(CW - Inches(0.4)), Inches(1.4),
                          cur["detail"], BODY_FONT, 12, MID, line_spacing=16)
        _add_rect(slide, Inches(0), Inches(3.65), W, Inches(0.04), GREEN)
        _add_text_box(slide, LM, Inches(3.75), Inches(3), Inches(0.25),
                      "Running Takeaways", TITLE_FONT, 10, GREEN, bold=True)
        for j in range(n + 1):
            ty = Emu(Inches(4.05) + j * Inches(0.35))
            active = j == n
            _add_rect(slide, LM, Emu(ty + Inches(0.05)), Inches(0.12), Inches(0.12), GREEN)
            _add_text_box(slide, Emu(LM + Inches(0.25)), ty,
                          Emu(CW - Inches(0.25)), Inches(0.32),
                          takeaways[j].get("summary", takeaways[j].get("headline", "")),
                          BODY_FONT, 10, DARK if active else MID, bold=active,
                          valign=MSO_ANCHOR.MIDDLE)


def build_closer(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background; fill = bg.fill; fill.solid(); fill.fore_color.rgb = GREEN
    _add_text_box(slide, Inches(0.5), Inches(1.4), Inches(9), Inches(1.2),
                  c.get("title", "Thank You"), TITLE_FONT, 44, WHITE, bold=True,
                  align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.BOTTOM)
    _add_rect(slide, Inches(3.75), Inches(2.75), Inches(2.5), Inches(0.04), GREEN_MID)
    if c.get("subtitle"):
        _add_text_box(slide, Inches(0.5), Inches(2.95), Inches(9), Inches(0.5),
                      c["subtitle"], BODY_FONT, 16, WHITE, align=PP_ALIGN.CENTER)
    if c.get("contact"):
        _add_text_box(slide, Inches(0.5), Inches(3.8), Inches(9), Inches(0.4),
                      c["contact"], BODY_FONT, 12, GREEN_LIGHT, align=PP_ALIGN.CENTER)


BUILDERS = {
    "title": build_title, "in_brief": build_in_brief,
    "section_divider": build_section_divider, "stat_callout": build_stat_callout,
    "quote": build_quote, "comparison": build_comparison,
    "text_graph": build_text_graph, "process_flow": build_process_flow,
    "matrix": build_matrix, "methods": build_methods,
    "hypotheses": build_hypotheses, "wsn_dense": build_wsn_dense,
    "wsn_reveal": build_wsn_reveal, "findings_recs": build_findings_recs,
    "findings_recs_dense": build_findings_recs_dense,
    "open_questions": build_open_questions, "agenda": build_agenda,
    "progressive_reveal": build_progressive_reveal, "closer": build_closer,
}

def build_deck(slide_configs, output_path):
    prs = Presentation(); prs.slide_width = W; prs.slide_height = H
    for slide_type, data in slide_configs:
        builder = BUILDERS.get(slide_type)
        if builder: builder(prs, data)
    prs.save(output_path); return output_path
