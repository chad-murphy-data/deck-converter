"""
Colorful template — python-pptx port.
Signature: colored header bars, multi-color card system with section theming.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn

W = Inches(10)
H = Inches(5.625)

# Colors
GREEN = RGBColor(0x36, 0x87, 0x27)
GREEN_LIGHT = RGBColor(0xE2, 0xF0, 0xD9)
GREEN_MID = RGBColor(0x1D, 0xE4, 0xCA)
BLUE = RGBColor(0x38, 0x80, 0xF3)
BLUE_LIGHT = RGBColor(0xE0, 0xEA, 0xFC)
PURPLE = RGBColor(0x5B, 0x2C, 0x8F)
PURPLE_LIGHT = RGBColor(0xED, 0xE4, 0xF5)
ORANGE = RGBColor(0x04, 0x54, 0x7C)
ORANGE_LIGHT = RGBColor(0xE0, 0xEE, 0xF4)
GOLD = RGBColor(0xD4, 0xA8, 0x43)
GOLD_LIGHT = RGBColor(0xF5, 0xEC, 0xD4)
DARK = RGBColor(0x40, 0x3F, 0x3E)
MID = RGBColor(0x55, 0x55, 0x55)
LIGHT = RGBColor(0xE8, 0xE8, 0xE8)
OFF_WHITE = RGBColor(0xF9, 0xF7, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

TITLE_FONT = "Calibri"
BODY_FONT = "Calibri"

COLORS = [GREEN, BLUE, PURPLE, ORANGE, GOLD]
LIGHTS = [GREEN_LIGHT, BLUE_LIGHT, PURPLE_LIGHT, ORANGE_LIGHT, GOLD_LIGHT]


def _add_rect(slide, x, y, w, h, fill_color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape


def _add_oval(slide, x, y, w, h, fill_color):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape


def _add_text_box(slide, x, y, w, h, text, font_name=BODY_FONT, font_size=12,
                  color=DARK, bold=False, italic=False, align=PP_ALIGN.LEFT,
                  valign=MSO_ANCHOR.TOP, line_spacing=None):
    txBox = slide.shapes.add_textbox(x, y, w, h)
    tf = txBox.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = text
    p.font.name = font_name
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.italic = italic
    p.alignment = align

    if line_spacing:
        p.line_spacing = Pt(line_spacing)

    bodyPr = tf._txBody.find(qn('a:bodyPr'))
    if bodyPr is not None:
        anchor_map = {MSO_ANCHOR.TOP: 't', MSO_ANCHOR.MIDDLE: 'ctr', MSO_ANCHOR.BOTTOM: 'b'}
        bodyPr.set('anchor', anchor_map.get(valign, 't'))

    return txBox


def _header(slide, title, color=GREEN):
    _add_rect(slide, Inches(0), Inches(0), W, Inches(1.0), color)
    _add_text_box(slide, Inches(0.5), Inches(0.15), Inches(8), Inches(0.7),
                  title, TITLE_FONT, 24, WHITE, bold=True, valign=MSO_ANCHOR.MIDDLE)


# ============================================================
# BUILDERS
# ============================================================

def build_title(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_rect(slide, Inches(0), Inches(0), Inches(0.12), H, GREEN)
    _add_text_box(slide, Inches(0.5), Inches(0.8), Inches(9), Inches(2.0),
                  c.get("title", "Title"), TITLE_FONT, 42, GREEN, bold=True,
                  valign=MSO_ANCHOR.BOTTOM)
    # Multi-color bar
    uw = Inches(1.5)
    bar_colors = [GREEN, BLUE, PURPLE, ORANGE]
    for i, col in enumerate(bar_colors):
        _add_rect(slide, Emu(Inches(0.5) + i * uw), Inches(2.92),
                  Emu(uw - Inches(0.08)), Inches(0.05), col)
    if c.get("subtitle"):
        _add_text_box(slide, Inches(0.5), Inches(3.1), Inches(9), Inches(0.5),
                      c["subtitle"], BODY_FONT, 16, MID)
    meta = []
    if c.get("author"):
        meta.append(c["author"])
    if c.get("date"):
        meta.append(c["date"])
    if meta:
        _add_text_box(slide, Inches(0.5), Inches(4.2), Inches(5), Inches(0.7),
                      "\n".join(meta), BODY_FONT, 12, MID)


def build_in_brief(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, c.get("title", "In Brief"))
    bullets = c.get("bullets", [])
    startY = Inches(1.2)
    rowH = Inches(0.95)
    gap = Inches(0.08)

    for i, b in enumerate(bullets):
        y = Emu(startY + i * (rowH + gap))
        col = COLORS[i % len(COLORS)]
        lt = LIGHTS[i % len(LIGHTS)]
        _add_rect(slide, Inches(0.5), y, Inches(9), rowH, lt)
        _add_rect(slide, Inches(0.5), y, Inches(0.12), rowH, col)
        _add_oval(slide, Inches(0.8), Emu(y + (rowH - Inches(0.42)) // 2),
                  Inches(0.42), Inches(0.42), col)
        _add_text_box(slide, Inches(0.8), Emu(y + (rowH - Inches(0.42)) // 2),
                      Inches(0.42), Inches(0.42), str(i + 1),
                      TITLE_FONT, 13, WHITE, bold=True, align=PP_ALIGN.CENTER,
                      valign=MSO_ANCHOR.MIDDLE)
        _add_text_box(slide, Inches(1.5), Emu(y + Inches(0.05)),
                      Inches(7.8), Emu(rowH - Inches(0.1)),
                      b, BODY_FONT, 13, DARK, valign=MSO_ANCHOR.MIDDLE)


def build_section_divider(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = GREEN

    if c.get("sectionNumber"):
        _add_text_box(slide, Inches(0.6), Inches(1.2), Inches(2), Inches(0.8),
                      f"0{c['sectionNumber']}", TITLE_FONT, 48, GREEN_MID, bold=True,
                      valign=MSO_ANCHOR.BOTTOM)
    _add_text_box(slide, Inches(0.6), Inches(2.0), Inches(8), Inches(1.2),
                  c.get("title", "Section"), TITLE_FONT, 36, WHITE, bold=True,
                  valign=MSO_ANCHOR.MIDDLE)
    if c.get("subtitle"):
        _add_text_box(slide, Inches(0.6), Inches(3.3), Inches(8), Inches(0.6),
                      c["subtitle"], BODY_FONT, 14, WHITE)


def build_stat_callout(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, c.get("title", "Key Metric"))
    _add_text_box(slide, Inches(0.5), Inches(1.3), Inches(9), Inches(1.8),
                  c.get("stat", "—"), TITLE_FONT, 72, GREEN, bold=True,
                  align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
    if c.get("headline"):
        _add_text_box(slide, Inches(1.0), Inches(3.1), Inches(8), Inches(0.7),
                      c["headline"], BODY_FONT, 18, DARK, bold=True, align=PP_ALIGN.CENTER)
    if c.get("detail"):
        _add_text_box(slide, Inches(1.5), Inches(3.8), Inches(7), Inches(0.8),
                      c["detail"], BODY_FONT, 12, MID, align=PP_ALIGN.CENTER)
    if c.get("source"):
        _add_text_box(slide, Inches(0.5), Inches(4.8), Inches(9), Inches(0.3),
                      c["source"], BODY_FONT, 9, MID, italic=True, align=PP_ALIGN.CENTER)


def build_quote(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, c.get("title", "In Their Words"))
    _add_text_box(slide, Inches(0.8), Inches(1.2), Inches(1), Inches(1),
                  "\u201C", "Georgia", 72, GREEN_LIGHT, bold=True)
    _add_text_box(slide, Inches(1.5), Inches(1.6), Inches(7.0), Inches(2.0),
                  c.get("quote", ""), BODY_FONT, 18, DARK, italic=True,
                  valign=MSO_ANCHOR.MIDDLE, line_spacing=25)
    _add_rect(slide, Inches(1.5), Inches(3.8), Inches(1.5), Inches(0.05), GREEN)
    if c.get("attribution"):
        _add_text_box(slide, Inches(1.5), Inches(3.95), Inches(7), Inches(0.5),
                      c["attribution"], BODY_FONT, 12, MID)
    if c.get("context"):
        _add_text_box(slide, Inches(1.5), Inches(4.3), Inches(7), Inches(0.4),
                      c["context"], BODY_FONT, 10, MID, italic=True)


def build_comparison(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, c.get("title", "Comparison"))

    # Left
    _add_rect(slide, Inches(0.5), Inches(1.2), Inches(4.2), Inches(3.6), ORANGE_LIGHT)
    _add_rect(slide, Inches(0.5), Inches(1.2), Inches(4.2), Inches(0.1), ORANGE)
    _add_text_box(slide, Inches(0.7), Inches(1.4), Inches(3.8), Inches(0.4),
                  c.get("leftLabel", "Before"), TITLE_FONT, 16, ORANGE, bold=True)
    for i, item in enumerate(c.get("leftItems", [])):
        _add_text_box(slide, Inches(0.7), Emu(Inches(1.95) + i * Inches(0.55)),
                      Inches(3.8), Inches(0.5), item, BODY_FONT, 11, DARK)

    # Divider with "vs" circle
    _add_rect(slide, Inches(4.85), Inches(1.4), Inches(0.14), Inches(3.2), GREEN)
    _add_oval(slide, Inches(4.72), Inches(2.7), Inches(0.4), Inches(0.4), GREEN)
    _add_text_box(slide, Inches(4.72), Inches(2.7), Inches(0.4), Inches(0.4),
                  "vs", BODY_FONT, 10, WHITE, bold=True, align=PP_ALIGN.CENTER,
                  valign=MSO_ANCHOR.MIDDLE)

    # Right
    _add_rect(slide, Inches(5.3), Inches(1.2), Inches(4.2), Inches(3.6), GREEN_LIGHT)
    _add_rect(slide, Inches(5.3), Inches(1.2), Inches(4.2), Inches(0.1), GREEN)
    _add_text_box(slide, Inches(5.5), Inches(1.4), Inches(3.8), Inches(0.4),
                  c.get("rightLabel", "After"), TITLE_FONT, 16, GREEN, bold=True)
    for i, item in enumerate(c.get("rightItems", [])):
        _add_text_box(slide, Inches(5.5), Emu(Inches(1.95) + i * Inches(0.55)),
                      Inches(3.8), Inches(0.5), item, BODY_FONT, 11, DARK)


def build_text_graph(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, c.get("title", "Title"))
    texts = c.get("text", [])
    if not isinstance(texts, list):
        texts = [texts]
    for i, t in enumerate(texts):
        _add_text_box(slide, Inches(0.5), Emu(Inches(1.2) + i * Inches(1.15)),
                      Inches(4.2), Inches(1.1), t, BODY_FONT, 12, DARK)
    _add_rect(slide, Inches(4.9), Inches(1.2), Inches(0.06), Inches(3.6), GREEN)

    from pptx.chart.data import CategoryChartData
    chart_data_raw = c.get("chartData", [{"name": "S1", "labels": ["A", "B", "C"], "values": [25, 45, 30]}])
    chart_data = CategoryChartData()
    if chart_data_raw:
        cd = chart_data_raw[0]
        chart_data.categories = cd.get("labels", ["A", "B", "C"])
        chart_data.add_series(cd.get("name", "Series 1"), cd.get("values", [25, 45, 30]))

    ct = XL_CHART_TYPE.COLUMN_CLUSTERED
    if c.get("chartType") == "line":
        ct = XL_CHART_TYPE.LINE
    elif c.get("chartType") == "pie":
        ct = XL_CHART_TYPE.PIE

    slide.shapes.add_chart(ct, Inches(5.15), Inches(1.0), Inches(4.5), Inches(4.0), chart_data)


def build_process_flow(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, c.get("title", "Process"))
    steps = c.get("steps", [])
    count = min(len(steps), 5)
    if count == 0:
        return

    total_w = 8.5
    arrow_w = 0.35
    step_w = (total_w - (count - 1) * arrow_w) / count
    startX = 0.75
    stepY = 1.4
    stepH = 3.4

    for i, step in enumerate(steps[:count]):
        x = Inches(startX + i * (step_w + arrow_w))
        y = Inches(stepY)
        sw = Inches(step_w)
        sh = Inches(stepH)
        col = COLORS[i % len(COLORS)]
        lt = LIGHTS[i % len(LIGHTS)]

        _add_rect(slide, x, y, sw, sh, lt)
        _add_rect(slide, x, y, sw, Inches(0.1), col)
        _add_oval(slide, Emu(x + Inches(0.1)), Emu(y + Inches(0.2)),
                  Inches(0.4), Inches(0.4), col)
        _add_text_box(slide, Emu(x + Inches(0.1)), Emu(y + Inches(0.2)),
                      Inches(0.4), Inches(0.4), str(i + 1),
                      TITLE_FONT, 13, WHITE, bold=True, align=PP_ALIGN.CENTER,
                      valign=MSO_ANCHOR.MIDDLE)
        _add_text_box(slide, Emu(x + Inches(0.1)), Emu(y + Inches(0.7)),
                      Emu(sw - Inches(0.2)), Inches(0.6),
                      step.get("title", ""), BODY_FONT, 11, DARK, bold=True)
        if step.get("detail"):
            _add_text_box(slide, Emu(x + Inches(0.1)), Emu(y + Inches(1.35)),
                          Emu(sw - Inches(0.2)), Inches(1.8),
                          step["detail"], BODY_FONT, 9, MID)
        if i < count - 1:
            _add_text_box(slide, Emu(x + sw + Inches(0.05)), Emu(y + Inches(1.2)),
                          Emu(Inches(arrow_w) - Inches(0.1)), Inches(0.5),
                          "\u2192", BODY_FONT, 18, MID, align=PP_ALIGN.CENTER,
                          valign=MSO_ANCHOR.MIDDLE)


def build_matrix(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, c.get("title", "Framework"))
    quads = c.get("quadrants", [{}, {}, {}, {}])
    qW = Inches(4.15)
    qH = Inches(1.65)
    gap = Inches(0.2)
    sX = Inches(0.85)
    sY = Inches(1.5)

    for i, q in enumerate(quads[:4]):
        col_idx = i % 2
        row = i // 2
        x = Emu(sX + col_idx * (qW + gap))
        y = Emu(sY + row * (qH + gap))
        col = COLORS[i]
        lt = LIGHTS[i]

        _add_rect(slide, x, y, qW, qH, lt)
        _add_rect(slide, x, y, qW, Inches(0.08), col)
        _add_text_box(slide, Emu(x + Inches(0.15)), Emu(y + Inches(0.12)),
                      Emu(qW - Inches(0.3)), Inches(0.3),
                      q.get("label", ""), TITLE_FONT, 12, col, bold=True)
        _add_text_box(slide, Emu(x + Inches(0.15)), Emu(y + Inches(0.45)),
                      Emu(qW - Inches(0.3)), Emu(qH - Inches(0.6)),
                      q.get("detail", ""), BODY_FONT, 10, DARK)


def build_methods(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, c.get("title", "Approach"))
    for i, f in enumerate(c.get("fields", [])):
        y = Emu(Inches(1.3) + i * Inches(0.8))
        col = COLORS[i % len(COLORS)]
        _add_rect(slide, Inches(0.5), y, Inches(0.08), Inches(0.7), col)
        _add_text_box(slide, Inches(0.75), y, Inches(2.0), Inches(0.7),
                      f.get("label", ""), BODY_FONT, 12, col, bold=True,
                      valign=MSO_ANCHOR.MIDDLE)
        _add_text_box(slide, Inches(2.8), y, Inches(6.7), Inches(0.7),
                      f.get("value", ""), BODY_FONT, 12, DARK,
                      valign=MSO_ANCHOR.MIDDLE)


def build_hypotheses(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, c.get("title", "Hypotheses"))
    for i, h in enumerate(c.get("hypotheses", [])):
        y = Emu(Inches(1.2) + i * Inches(0.75))
        col = COLORS[i % len(COLORS)]
        lt = LIGHTS[i % len(LIGHTS)]
        _add_rect(slide, Inches(0.5), y, Inches(8.2), Inches(0.65), lt)
        _add_rect(slide, Inches(0.5), y, Inches(0.1), Inches(0.65), col)
        _add_text_box(slide, Inches(0.75), y, Inches(0.5), Inches(0.65),
                      f"H{i + 1}", TITLE_FONT, 14, col, bold=True,
                      align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)
        _add_text_box(slide, Inches(1.3), y, Inches(6.2), Inches(0.65),
                      h.get("text", ""), BODY_FONT, 12, DARK,
                      valign=MSO_ANCHOR.MIDDLE)
        if h.get("status"):
            sc = GREEN if h["status"] == "Confirmed" else (ORANGE if h["status"] == "Rejected" else MID)
            _add_rect(slide, Inches(7.8), Emu(y + (Inches(0.65) - Inches(0.3)) // 2),
                      Inches(0.9), Inches(0.3), sc)
            _add_text_box(slide, Inches(7.8), Emu(y + (Inches(0.65) - Inches(0.3)) // 2),
                          Inches(0.9), Inches(0.3),
                          h["status"], BODY_FONT, 8, WHITE, bold=True,
                          align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.MIDDLE)


def build_wsn_dense(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, c.get("title", "Key Finding"))
    cols = [
        ("What", GREEN, GREEN_LIGHT, c.get("what", {})),
        ("So What", BLUE, BLUE_LIGHT, c.get("soWhat", {})),
        ("Now What", PURPLE, PURPLE_LIGHT, c.get("nowWhat", {})),
    ]
    colW = Inches(2.85)
    gap = Inches(0.2)
    startX = Inches(0.5)
    startY = Inches(1.2)
    cardH = Inches(3.8)

    for i, (label, color, light, data) in enumerate(cols):
        x = Emu(startX + i * (colW + gap))
        _add_rect(slide, x, startY, colW, cardH, light)
        _add_rect(slide, x, startY, colW, Inches(0.1), color)
        _add_text_box(slide, Emu(x + Inches(0.15)), Emu(startY + Inches(0.15)),
                      Emu(colW - Inches(0.3)), Inches(0.35),
                      label, TITLE_FONT, 14, color, bold=True)
        _add_text_box(slide, Emu(x + Inches(0.15)), Emu(startY + Inches(0.55)),
                      Emu(colW - Inches(0.3)), Inches(1.2),
                      data.get("headline", ""), BODY_FONT, 12, DARK, bold=True)
        if data.get("detail"):
            _add_text_box(slide, Emu(x + Inches(0.15)), Emu(startY + Inches(1.8)),
                          Emu(colW - Inches(0.3)), Inches(1.7),
                          data["detail"], BODY_FONT, 10, MID)


def build_wsn_reveal(prs, c):
    def _zone(slide, x, label, color, light, data, condensed=False):
        h = Inches(2.0) if condensed else Inches(3.4)
        y = Inches(1.2)
        w = Inches(4.15)
        _add_rect(slide, x, y, w, h, light)
        _add_rect(slide, x, y, w, Inches(0.1), color)
        fs = 10 if condensed else 11
        _add_text_box(slide, Emu(x + Inches(0.12)), Emu(y + Inches(0.15)),
                      Inches(2), Inches(0.3), label, TITLE_FONT, fs, color, bold=True)
        _add_text_box(slide, Emu(x + Inches(0.12)), Emu(y + Inches(0.5)),
                      Emu(w - Inches(0.24)), Inches(0.7) if condensed else Inches(0.9),
                      data.get("headline", ""), BODY_FONT, 10 if condensed else 13,
                      DARK, bold=True)
        if data.get("detail"):
            _add_text_box(slide, Emu(x + Inches(0.12)),
                          Emu(y + Inches(1.2) if condensed else y + Inches(1.45)),
                          Emu(w - Inches(0.24)), Inches(0.55) if condensed else Inches(1.6),
                          data["detail"], BODY_FONT, 8.5 if condensed else 11, MID)

    leftX = Inches(0.5)
    rightX = Inches(5.2)

    def _hdr(slide):
        _add_rect(slide, Inches(0), Inches(0), W, Inches(0.15), GREEN)
        _add_text_box(slide, Inches(0.5), Inches(0.3), Inches(9), Inches(0.65),
                      c.get("title", "Key Finding"), TITLE_FONT, 28, DARK, bold=True)

    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    _hdr(s1)
    _zone(s1, leftX, "What We Found", GREEN, GREEN_LIGHT, c.get("what", {}))

    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    _hdr(s2)
    _zone(s2, leftX, "What We Found", GREEN, GREEN_LIGHT, c.get("what", {}))
    _zone(s2, rightX, "So What", BLUE, BLUE_LIGHT, c.get("soWhat", {}))

    s3 = prs.slides.add_slide(prs.slide_layouts[6])
    _hdr(s3)
    _zone(s3, leftX, "What We Found", GREEN, GREEN_LIGHT, c.get("what", {}), True)
    _zone(s3, rightX, "So What", BLUE, BLUE_LIGHT, c.get("soWhat", {}), True)

    nwY = Inches(3.4)
    nwH = Inches(1.85)
    _add_rect(s3, Inches(0.5), nwY, Inches(9.0), nwH, PURPLE_LIGHT)
    _add_rect(s3, Inches(0.5), nwY, Inches(9.0), Inches(0.1), PURPLE)
    _add_text_box(s3, Inches(0.65), Emu(nwY + Inches(0.15)),
                  Inches(2), Inches(0.3), "Now What", TITLE_FONT, 12, PURPLE, bold=True)
    nw = c.get("nowWhat", {})
    _add_text_box(s3, Inches(0.65), Emu(nwY + Inches(0.5)),
                  Inches(8.7), Inches(0.65),
                  nw.get("headline", ""), BODY_FONT, 18, DARK, bold=True)
    if nw.get("detail"):
        _add_text_box(s3, Inches(0.65), Emu(nwY + Inches(1.15)),
                      Inches(8.7), Inches(0.55), nw["detail"], BODY_FONT, 11, MID)


def build_findings_recs(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, c.get("title", "Findings & Recommendations"))
    items = c.get("items", [])
    startY = Inches(1.15)
    rowH = Inches(0.82)
    gap = Inches(0.08)

    for i, item in enumerate(items[:5]):
        y = Emu(startY + i * (rowH + gap))
        col = COLORS[i % len(COLORS)]
        lt = LIGHTS[i % len(LIGHTS)]

        _add_rect(slide, Inches(0.5), y, Inches(3.9), rowH, lt)
        _add_rect(slide, Inches(0.5), y, Inches(0.1), rowH, col)
        _add_text_box(slide, Inches(0.75), y, Inches(3.5), rowH,
                      item.get("finding", ""), BODY_FONT, 10.5, DARK,
                      valign=MSO_ANCHOR.MIDDLE)

        _add_oval(slide, Inches(4.62), Emu(y + (rowH - Inches(0.42)) // 2),
                  Inches(0.42), Inches(0.42), col)
        _add_text_box(slide, Inches(4.62), Emu(y + (rowH - Inches(0.42)) // 2),
                      Inches(0.42), Inches(0.42), "\u2192",
                      BODY_FONT, 14, WHITE, bold=True, align=PP_ALIGN.CENTER,
                      valign=MSO_ANCHOR.MIDDLE)

        _add_rect(slide, Inches(5.2), y, Inches(4.3), rowH, lt)
        _add_rect(slide, Inches(5.2), y, Inches(0.1), rowH, BLUE)
        _add_text_box(slide, Inches(5.45), y, Inches(3.9), rowH,
                      item.get("recommendation", ""), BODY_FONT, 10.5, DARK,
                      valign=MSO_ANCHOR.MIDDLE)


def build_findings_recs_dense(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, c.get("title", "Complete Findings"))
    items = c.get("items", [])
    startY = Inches(1.15)
    rowH = Inches(0.47)
    gap = Inches(0.06)

    for i, item in enumerate(items[:8]):
        y = Emu(startY + i * (rowH + gap))
        col = COLORS[i % len(COLORS)]
        bg = OFF_WHITE if i % 2 == 0 else WHITE

        _add_rect(slide, Inches(0.5), y, Inches(4.2), rowH, bg)
        _add_oval(slide, Inches(0.2), Emu(y + (rowH - Inches(0.3)) // 2),
                  Inches(0.3), Inches(0.3), col)
        _add_text_box(slide, Inches(0.2), Emu(y + (rowH - Inches(0.3)) // 2),
                      Inches(0.3), Inches(0.3), str(i + 1),
                      TITLE_FONT, 9, WHITE, bold=True, align=PP_ALIGN.CENTER,
                      valign=MSO_ANCHOR.MIDDLE)
        _add_text_box(slide, Inches(0.55), y, Inches(4.1), rowH,
                      item.get("finding", ""), BODY_FONT, 9, DARK,
                      valign=MSO_ANCHOR.MIDDLE)

        _add_rect(slide, Inches(4.85), y, Inches(4.65), rowH, bg)
        _add_rect(slide, Inches(4.85), y, Inches(0.06), rowH, BLUE)
        _add_text_box(slide, Inches(5.0), y, Inches(4.4), rowH,
                      item.get("recommendation", ""), BODY_FONT, 9, DARK,
                      valign=MSO_ANCHOR.MIDDLE)


def build_open_questions(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, c.get("title", "Open Questions"))
    questions = c.get("questions", [])
    cardW = Inches(4.2)
    cardH = Inches(1.7)
    gX = Inches(0.3)
    gY = Inches(0.2)
    gridX = Inches(0.5)
    gridY = Inches(1.2)

    for i, question in enumerate(questions[:4]):
        col_idx = i % 2
        row = i // 2
        x = Emu(gridX + col_idx * (cardW + gX))
        y = Emu(gridY + row * (cardH + gY))
        col = COLORS[i % len(COLORS)]
        lt = LIGHTS[i % len(LIGHTS)]

        _add_rect(slide, x, y, cardW, cardH, lt)
        _add_rect(slide, x, y, cardW, Inches(0.08), col)
        _add_text_box(slide, Emu(x + Inches(0.15)), Emu(y + Inches(0.15)),
                      Inches(0.5), Inches(0.5), str(i + 1),
                      TITLE_FONT, 26, col, bold=True)
        _add_text_box(slide, Emu(x + Inches(0.15)), Emu(y + Inches(0.7)),
                      Emu(cardW - Inches(0.3)), Inches(0.85),
                      question, BODY_FONT, 12, DARK)


def build_agenda(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _header(slide, c.get("title", "Agenda"))
    items = c.get("items", [])
    for i, item in enumerate(items):
        y = Emu(Inches(1.2) + i * Inches(0.75))
        col = COLORS[i % len(COLORS)]
        lt = LIGHTS[i % len(LIGHTS)]

        _add_rect(slide, Inches(0.5), y, Inches(9), Inches(0.65), lt)
        _add_rect(slide, Inches(0.5), y, Inches(0.1), Inches(0.65), col)
        _add_oval(slide, Inches(0.8), Emu(y + (Inches(0.65) - Inches(0.4)) // 2),
                  Inches(0.4), Inches(0.4), col)
        _add_text_box(slide, Inches(0.8), Emu(y + (Inches(0.65) - Inches(0.4)) // 2),
                      Inches(0.4), Inches(0.4), str(i + 1),
                      TITLE_FONT, 13, WHITE, bold=True, align=PP_ALIGN.CENTER,
                      valign=MSO_ANCHOR.MIDDLE)

        title_text = item if isinstance(item, str) else item.get("title", "")
        _add_text_box(slide, Inches(1.45), y, Inches(5.5), Inches(0.65),
                      title_text, BODY_FONT, 14, DARK, bold=True,
                      valign=MSO_ANCHOR.MIDDLE)
        if isinstance(item, dict) and item.get("detail"):
            _add_text_box(slide, Inches(7.0), y, Inches(2.3), Inches(0.65),
                          item["detail"], BODY_FONT, 10, MID,
                          align=PP_ALIGN.RIGHT, valign=MSO_ANCHOR.MIDDLE)


def build_progressive_reveal(prs, c):
    takeaways = c.get("takeaways", [])
    for n in range(min(len(takeaways), 5)):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _add_rect(slide, Inches(0), Inches(0), W, Inches(0.1), GREEN)
        _add_text_box(slide, Inches(0.5), Inches(0.25), Inches(9), Inches(0.6),
                      c.get("title", "Building the Picture"), TITLE_FONT, 24, DARK, bold=True)

        cur = takeaways[n]
        col = COLORS[n % len(COLORS)]
        lt = LIGHTS[n % len(LIGHTS)]

        _add_rect(slide, Inches(0.5), Inches(1.1), Inches(9), Inches(2.4), lt)
        _add_rect(slide, Inches(0.5), Inches(1.1), Inches(9), Inches(0.1), col)
        _add_text_box(slide, Inches(0.7), Inches(1.3), Inches(8.6), Inches(0.7),
                      cur.get("headline", ""), BODY_FONT, 16, DARK, bold=True)
        if cur.get("detail"):
            _add_text_box(slide, Inches(0.7), Inches(2.05), Inches(8.6), Inches(1.2),
                          cur["detail"], BODY_FONT, 11, MID)

        _add_rect(slide, Inches(0), Inches(3.7), W, Inches(0.08), GREEN)
        _add_text_box(slide, Inches(0.5), Inches(3.85), Inches(3), Inches(0.25),
                      "Running Takeaways", TITLE_FONT, 9, GREEN, bold=True)

        for j in range(n + 1):
            ty = Emu(Inches(4.15) + j * Inches(0.32))
            active = j == n
            jcol = COLORS[j % len(COLORS)]
            _add_oval(slide, Inches(0.5), Emu(ty + Inches(0.02)),
                      Inches(0.18), Inches(0.18), jcol)
            _add_text_box(slide, Inches(0.8), ty, Inches(8.7), Inches(0.3),
                          takeaways[j].get("summary", takeaways[j].get("headline", "")),
                          BODY_FONT, 9, DARK if active else MID, bold=active,
                          valign=MSO_ANCHOR.MIDDLE)


def build_closer(prs, c):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = GREEN

    _add_text_box(slide, Inches(0.5), Inches(1.4), Inches(9), Inches(1.2),
                  c.get("title", "Thank You"), TITLE_FONT, 44, WHITE, bold=True,
                  align=PP_ALIGN.CENTER, valign=MSO_ANCHOR.BOTTOM)
    _add_rect(slide, Inches(3.75), Inches(2.75), Inches(2.5), Inches(0.05), GREEN_MID)
    if c.get("subtitle"):
        _add_text_box(slide, Inches(0.5), Inches(2.95), Inches(9), Inches(0.5),
                      c["subtitle"], BODY_FONT, 16, WHITE, align=PP_ALIGN.CENTER)
    if c.get("contact"):
        _add_text_box(slide, Inches(0.5), Inches(3.8), Inches(9), Inches(0.4),
                      c["contact"], BODY_FONT, 12, GREEN_LIGHT, align=PP_ALIGN.CENTER)


# ============================================================
# DISPATCH
# ============================================================
BUILDERS = {
    "title": build_title,
    "in_brief": build_in_brief,
    "section_divider": build_section_divider,
    "stat_callout": build_stat_callout,
    "quote": build_quote,
    "comparison": build_comparison,
    "text_graph": build_text_graph,
    "process_flow": build_process_flow,
    "matrix": build_matrix,
    "methods": build_methods,
    "hypotheses": build_hypotheses,
    "wsn_dense": build_wsn_dense,
    "wsn_reveal": build_wsn_reveal,
    "findings_recs": build_findings_recs,
    "findings_recs_dense": build_findings_recs_dense,
    "open_questions": build_open_questions,
    "agenda": build_agenda,
    "progressive_reveal": build_progressive_reveal,
    "closer": build_closer,
}


def build_deck(slide_configs, output_path):
    """Build a complete deck from a list of (slide_type, data_dict) tuples."""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)

    for slide_type, data in slide_configs:
        if slide_type == "skip":
            continue
        builder = BUILDERS.get(slide_type)
        if builder:
            builder(prs, data)

    prs.save(output_path)
    return output_path
