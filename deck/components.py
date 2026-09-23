"""
Composants réutilisables : masters de slides, en-têtes, cartes KPI,
timelines, tableaux et graphiques natifs.
"""

from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION

from theme import *

STATE = {"n": 1, "total": 0}  # la couverture compte pour la page 01


# --------------------------------------------------------------------------
# MASTERS
# --------------------------------------------------------------------------
def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def light_slide(prs, title, kicker=None, subtitle=None, footer=True):
    """Slide de contenu sur fond clair."""
    s = blank(prs)
    rect(s, 0, 0, SW, SH, PAPER)
    # filet d'accent en haut à gauche
    rect(s, 0, 0, Inches(0.09), SH, ACCENT)

    if kicker:
        text(s, MARGIN, Inches(0.62), CONTENT_W, Inches(0.24), kicker.upper(),
             size=10.5, color=ACCENT, bold=True, font=FONT_SEMI, spacing=2.2)
    text(s, MARGIN, Inches(0.95), CONTENT_W, Inches(0.55), title,
         size=29, color=GRAPHITE, bold=True, font=FONT_SEMI, line=1.05)
    rect(s, MARGIN, Inches(1.62), Inches(1.15), Pt(3), ACCENT)
    if subtitle:
        text(s, MARGIN, Inches(1.8), Inches(9.9), Inches(0.42), subtitle,
             size=12.5, color=SLATE, line=1.3)
    if footer:
        _footer(s, dark=False)
    return s


def dark_slide(prs, title, kicker=None, subtitle=None, footer=True, bg=None):
    s = blank(prs)
    rect(s, 0, 0, SW, SH, INK)
    if bg:
        picture_bg(s, bg, prs)
        scrim(s, prs, INK, 62000)
    rect(s, 0, 0, Inches(0.09), SH, ACCENT)
    if kicker:
        text(s, MARGIN, Inches(0.62), CONTENT_W, Inches(0.24), kicker.upper(),
             size=10.5, color=TEAL, bold=True, font=FONT_SEMI, spacing=2.2)
    text(s, MARGIN, Inches(0.95), CONTENT_W, Inches(0.55), title,
         size=29, color=WHITE, bold=True, font=FONT_SEMI, line=1.05)
    rect(s, MARGIN, Inches(1.62), Inches(1.15), Pt(3), TEAL)
    if subtitle:
        text(s, MARGIN, Inches(1.8), Inches(9.9), Inches(0.42), subtitle,
             size=12.5, color=SLATE_DK, line=1.3)
    if footer:
        _footer(s, dark=True)
    return s


def _footer(slide, dark=False):
    STATE["n"] += 1
    c = SLATE_DK if dark else SLATE
    line_c = RGBColor(0x1E, 0x33, 0x4A) if dark else LINE_LIGHT
    rect(slide, MARGIN, Inches(6.72), CONTENT_W, Pt(0.75), line_c)
    text(slide, MARGIN, FOOT_Y, Inches(7), Inches(0.24),
         "SHEP  ·  Dossier stratégique confidentiel  ·  2026",
         size=8.5, color=c, spacing=1.2)
    text(slide, SW - MARGIN - Inches(1.2), FOOT_Y, Inches(1.2), Inches(0.24),
         f"{STATE['n']:02d}", size=9, color=c, bold=True,
         font=FONT_SEMI, align=PP_ALIGN.RIGHT)


# --------------------------------------------------------------------------
# COMPOSANTS
# --------------------------------------------------------------------------
def kpi_card(slide, x, y, w, h, value, label, note=None, accent=ACCENT,
             dark=False, compact=False):
    """Carte KPI. `compact` pour les hauteurs < 1.6 pouce."""
    bg = INK_SOFT if dark else WHITE
    c = card(slide, x, y, w, h, bg, corner=0.09, shadow=not dark,
             border=None if dark else LINE_LIGHT)
    pad = Inches(0.26)
    if compact:
        rect(slide, x + pad, y + Inches(0.24), Pt(3.5), h - Inches(0.48), accent)
        tx = x + pad + Inches(0.22)
        tw = w - pad - Inches(0.48)
        text(slide, tx, y + Inches(0.21), tw, Inches(0.42), value,
             size=24, color=WHITE if dark else GRAPHITE, bold=True,
             font=FONT_SEMI, line=1.0)
        text(slide, tx, y + Inches(0.63), tw, Inches(0.26), label,
             size=11, color=WHITE if dark else GRAPHITE, bold=True,
             font=FONT_SEMI)
        if note:
            text(slide, tx, y + Inches(0.88), tw, Inches(0.3), note,
                 size=9.5, color=SLATE_DK if dark else SLATE, line=1.2)
        return c
    rect(slide, x + Inches(0.28), y + Inches(0.34), Inches(0.42), Pt(3), accent)
    text(slide, x + Inches(0.28), y + Inches(0.6), w - Inches(0.5), Inches(0.62),
         value, size=34, color=WHITE if dark else GRAPHITE, bold=True,
         font=FONT_SEMI, line=1.0)
    text(slide, x + Inches(0.28), y + Inches(1.2), w - Inches(0.5), Inches(0.3),
         label, size=11.5, color=WHITE if dark else GRAPHITE,
         bold=True, font=FONT_SEMI)
    if note:
        text(slide, x + Inches(0.28), y + Inches(1.52), w - Inches(0.5), Inches(0.6),
             note, size=9.5, color=SLATE_DK if dark else SLATE, line=1.25)
    return c


def feature_card(slide, x, y, w, h, num, title, body, accent=ACCENT, dark=False):
    bg = INK_SOFT if dark else WHITE
    card(slide, x, y, w, h, bg, corner=0.07, shadow=not dark,
         border=None if dark else LINE_LIGHT)
    badge = rect(slide, x + Inches(0.3), y + Inches(0.32), Inches(0.44),
                 Inches(0.44), accent, MSO_SHAPE.ROUNDED_RECTANGLE, 0.28)
    text(slide, x + Inches(0.3), y + Inches(0.4), Inches(0.44), Inches(0.3),
         num, size=12, color=WHITE, bold=True, font=FONT_SEMI,
         align=PP_ALIGN.CENTER)
    text(slide, x + Inches(0.3), y + Inches(0.95), w - Inches(0.6), Inches(0.36),
         title, size=14, color=WHITE if dark else GRAPHITE, bold=True,
         font=FONT_SEMI, line=1.1)
    text(slide, x + Inches(0.3), y + Inches(1.42), w - Inches(0.6), h - Inches(1.7),
         body, size=11, color=SLATE_DK if dark else SLATE, line=1.38)


def pill(slide, x, y, w, h, label, fill=ACCENT, color=WHITE, size=10):
    s = rect(slide, x, y, w, h, fill, MSO_SHAPE.ROUNDED_RECTANGLE, 0.5)
    text(slide, x, y + (h - Inches(0.18)) / 2, w, Inches(0.2), label,
         size=size, color=color, bold=True, font=FONT_SEMI,
         align=PP_ALIGN.CENTER)
    return s


def timeline(slide, x, y, w, phases, dark=False):
    """phases = [(période, titre, [jalons]), ...]"""
    n = len(phases)
    gap = Inches(0.26)
    cw = (w - gap * (n - 1)) / n
    axis_y = y + Inches(0.52)
    rect(slide, x, axis_y, w, Pt(2),
         RGBColor(0x22, 0x3A, 0x53) if dark else LINE_LIGHT)
    cols = [ACCENT, TEAL, AMBER, CORAL, ACCENT_DK]
    for i, (period, title, items) in enumerate(phases):
        cx = x + i * (cw + gap)
        acc = cols[i % len(cols)]
        text(slide, cx, y, cw, Inches(0.26), period.upper(), size=9.5,
             color=acc, bold=True, font=FONT_SEMI, spacing=1.6)
        dot = rect(slide, cx, axis_y - Inches(0.07), Inches(0.19), Inches(0.19),
                   acc, MSO_SHAPE.OVAL)
        rect(slide, cx + Inches(0.055), axis_y - Inches(0.015),
             Inches(0.08), Inches(0.08),
             INK if dark else PAPER, MSO_SHAPE.OVAL)
        text(slide, cx, axis_y + Inches(0.3), cw - Inches(0.1), Inches(0.4),
             title, size=13, color=WHITE if dark else GRAPHITE, bold=True,
             font=FONT_SEMI, line=1.15)
        bullets(slide, cx, axis_y + Inches(0.92), cw - Inches(0.15), Inches(2.0),
                items, size=10.5,
                color=SLATE_DK if dark else SLATE, bullet_color=acc, gap=6)


def comparison_table(slide, x, y, w, headers, rows, col_w=None, dark=False,
                     highlight_col=1):
    """Tableau dessiné à la main (contrôle total du style)."""
    n = len(headers)
    col_w = col_w or [w / n] * n
    rh = Inches(0.46)
    hh = Inches(0.52)
    # en-tête
    cx = x
    rect(slide, x, y, w, hh, NAVY if not dark else INK_SOFT,
         MSO_SHAPE.ROUNDED_RECTANGLE, 0.12)
    rect(slide, x, y + hh - Inches(0.12), w, Inches(0.12),
         NAVY if not dark else INK_SOFT)
    for i, hd in enumerate(headers):
        al = PP_ALIGN.LEFT if i == 0 else PP_ALIGN.CENTER
        text(slide, cx + Inches(0.22), y + Inches(0.17),
             col_w[i] - Inches(0.3), Inches(0.26), hd,
             size=10.5, color=WHITE, bold=True, font=FONT_SEMI, align=al,
             spacing=0.8)
        cx += col_w[i]
    # lignes
    for r, row in enumerate(rows):
        ry = y + hh + r * rh
        if r % 2 == 0:
            rect(slide, x, ry, w, rh,
                 RGBColor(0xEF, 0xF3, 0xF9) if not dark else RGBColor(0x0D, 0x1C, 0x2E))
        # colonne mise en avant
        cx = x + sum(col_w[:highlight_col])
        rect(slide, cx, ry, col_w[highlight_col], rh,
             RGBColor(0xE6, 0xEF, 0xFF) if not dark else RGBColor(0x11, 0x28, 0x45))
        cx = x
        for i, cell in enumerate(row):
            al = PP_ALIGN.LEFT if i == 0 else PP_ALIGN.CENTER
            is_head = (i == 0)
            text(slide, cx + Inches(0.22), ry + Inches(0.145),
                 col_w[i] - Inches(0.3), Inches(0.3), cell,
                 size=10.5,
                 color=(GRAPHITE if not dark else WHITE) if is_head or i == highlight_col
                 else (SLATE if not dark else SLATE_DK),
                 bold=is_head or i == highlight_col,
                 font=FONT_SEMI if (is_head or i == highlight_col) else FONT,
                 align=al)
            cx += col_w[i]
        rect(slide, x, ry + rh - Pt(0.5), w, Pt(0.5),
             LINE_LIGHT if not dark else RGBColor(0x1C, 0x2F, 0x45))


# --------------------------------------------------------------------------
# GRAPHIQUES NATIFS
# --------------------------------------------------------------------------
def _style_chart(chart, dark=False, legend=True, font_size=10):
    txt = SLATE_DK if dark else SLATE
    chart.font.size = Pt(font_size)
    chart.font.name = FONT
    chart.font.color.rgb = txt
    if legend:
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.TOP
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(font_size)
        chart.legend.font.color.rgb = txt
    else:
        chart.has_legend = False


def bar_chart(slide, x, y, w, h, categories, series, dark=False,
              colors=None, labels=True, number_format='#,##0'):
    data = CategoryChartData()
    data.categories = categories
    for name, vals in series:
        data.add_series(name, vals)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                                int(x), int(y), int(w), int(h), data)
    chart = gf.chart
    _style_chart(chart, dark, legend=len(series) > 1)
    colors = colors or [ACCENT, TEAL, AMBER]
    plot = chart.plots[0]
    plot.gap_width = 55
    plot.overlap = -12 if len(series) > 1 else 0
    for i, ser in enumerate(chart.series):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = colors[i % len(colors)]
        ser.format.line.fill.background()
    if labels:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.number_format = number_format
        dl.number_format_is_linked = False
        dl.font.size = Pt(9.5)
        dl.font.bold = True
        dl.font.color.rgb = WHITE if dark else GRAPHITE
        dl.position = XL_LABEL_POSITION.OUTSIDE_END
    va = chart.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = (
        RGBColor(0x1E, 0x33, 0x4A) if dark else LINE_LIGHT)
    va.major_gridlines.format.line.width = Pt(0.75)
    va.format.line.fill.background()
    va.has_title = False
    va.tick_labels.font.size = Pt(9)
    ca = chart.category_axis
    ca.format.line.color.rgb = RGBColor(0x1E, 0x33, 0x4A) if dark else LINE_LIGHT
    ca.has_major_gridlines = False
    ca.tick_labels.font.size = Pt(9.5)
    return chart


def line_chart(slide, x, y, w, h, categories, series, dark=False, colors=None):
    data = CategoryChartData()
    data.categories = categories
    for name, vals in series:
        data.add_series(name, vals)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS,
                                int(x), int(y), int(w), int(h), data)
    chart = gf.chart
    _style_chart(chart, dark, legend=len(series) > 1)
    colors = colors or [ACCENT, TEAL, AMBER]
    for i, ser in enumerate(chart.series):
        ser.format.line.color.rgb = colors[i % len(colors)]
        ser.format.line.width = Pt(2.75)
        ser.smooth = False
        m = ser.marker
        m.style = 8  # circle
        m.size = 7
        m.format.fill.solid()
        m.format.fill.fore_color.rgb = colors[i % len(colors)]
        m.format.line.color.rgb = WHITE if not dark else INK
        m.format.line.width = Pt(1.25)
    va = chart.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = (
        RGBColor(0x1E, 0x33, 0x4A) if dark else LINE_LIGHT)
    va.major_gridlines.format.line.width = Pt(0.75)
    va.format.line.fill.background()
    ca = chart.category_axis
    ca.format.line.color.rgb = RGBColor(0x1E, 0x33, 0x4A) if dark else LINE_LIGHT
    ca.has_major_gridlines = False
    return chart


def donut_chart(slide, x, y, w, h, categories, values, colors=None, dark=False):
    data = CategoryChartData()
    data.categories = categories
    data.add_series("Répartition", values)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT,
                                int(x), int(y), int(w), int(h), data)
    chart = gf.chart
    _style_chart(chart, dark, legend=True, font_size=10)
    chart.legend.position = XL_LEGEND_POSITION.RIGHT
    colors = colors or [ACCENT, TEAL, AMBER, CORAL, NAVY]
    pts = chart.plots[0]
    pts.has_data_labels = True
    dl = pts.data_labels
    dl.number_format = '0"%"'
    dl.number_format_is_linked = False
    dl.font.size = Pt(9.5)
    dl.font.bold = True
    dl.font.color.rgb = WHITE
    for i, pt in enumerate(chart.series[0].points):
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = colors[i % len(colors)]
        pt.format.line.color.rgb = PAPER if not dark else INK
        pt.format.line.width = Pt(2)
    return chart


def stacked_bar(slide, x, y, w, h, categories, series, dark=False, colors=None):
    data = CategoryChartData()
    data.categories = categories
    for name, vals in series:
        data.add_series(name, vals)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED,
                                int(x), int(y), int(w), int(h), data)
    chart = gf.chart
    _style_chart(chart, dark, legend=True)
    colors = colors or [ACCENT, TEAL, AMBER, CORAL]
    chart.plots[0].gap_width = 60
    for i, ser in enumerate(chart.series):
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = colors[i % len(colors)]
        ser.format.line.fill.background()
    va = chart.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = (
        RGBColor(0x1E, 0x33, 0x4A) if dark else LINE_LIGHT)
    va.format.line.fill.background()
    chart.category_axis.format.line.color.rgb = (
        RGBColor(0x1E, 0x33, 0x4A) if dark else LINE_LIGHT)
    return chart
