"""
Zambia Economic Recovery — Follow-Up Assignment
Storytelling with Data: Big Idea + 3-Minute Story Structure
Author: Boldwin Mweemba
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE

# ── Colour palette ────────────────────────────────────────────────────────────
DARK_BG    = RGBColor(0x0D, 0x1B, 0x2A)
ACCENT     = RGBColor(0xE8, 0x8C, 0x00)   # amber / copper
ACCENT2    = RGBColor(0x2E, 0xCC, 0x71)   # emerald green
LIGHT_TEXT = RGBColor(0xF5, 0xF5, 0xF5)
MID_TEXT   = RGBColor(0xB0, 0xBE, 0xC5)
RED_WARN   = RGBColor(0xE7, 0x4C, 0x3C)
DIVIDER    = RGBColor(0x1E, 0x3A, 0x5F)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

# ── Data ──────────────────────────────────────────────────────────────────────
YEARS_ALL   = ['2015','2016','2017','2018','2019','2020','2021','2022','2023','2024']

EXT_DEBT    = [12.3, 14.0, 17.5, 22.4, 29.1, 31.5, 29.8, 26.4, 22.7, 18.9]   # $ billions

GDP_GROWTH  = [2.9, 3.8, 3.5, 4.0, 1.4, -2.8, 4.6, 4.7, 4.7, 5.3]            # % annual

CURR_ACCT   = [-3.6, -4.1, -1.9, -1.7, -1.5, 2.9, 11.9, 4.7, 3.1, 2.0]       # % of GDP

FDI         = [3.2, 2.5, 3.5, 2.8, 2.4, 1.8, 3.7, 5.1, 6.8, 9.32]            # % of GDP

# ── Text helpers ──────────────────────────────────────────────────────────────

def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_rect(slide, left, top, width, height, color):
    s = slide.shapes.add_shape(1, left, top, width, height)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    return s


def add_textbox(slide, left, top, width, height, text, size, bold=False,
                color=None, align=PP_ALIGN.LEFT, italic=False):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color if color else LIGHT_TEXT
    r.font.name = "Calibri"
    return tb


def add_multiline(slide, left, top, width, height, lines, size,
                  color=None, bold=False, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for item in lines:
        text = item[0] if isinstance(item, (list, tuple)) else item
        b    = item[1] if isinstance(item, (list, tuple)) and len(item) > 1 else bold
        c    = item[2] if isinstance(item, (list, tuple)) and len(item) > 2 else color
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.bold = b
        r.font.color.rgb = c if c else (color if color else LIGHT_TEXT)
        r.font.name = "Calibri"
    return tb


def add_bar(slide, left, top, width, color=None):
    ln = slide.shapes.add_shape(1, left, top, width, Pt(1.5))
    ln.fill.solid()
    ln.fill.fore_color.rgb = color if color else ACCENT
    ln.line.fill.background()


def stat_card(slide, left, top, width, height, metric, value, sub, val_color=None):
    add_rect(slide, left, top, width, height, DIVIDER)
    add_textbox(slide, left + Inches(0.1), top + Inches(0.12),
                width - Inches(0.2), Inches(0.4),
                metric, 9, color=MID_TEXT, align=PP_ALIGN.CENTER)
    add_textbox(slide, left + Inches(0.05), top + Inches(0.48),
                width - Inches(0.1), Inches(0.55),
                value, 22, bold=True,
                color=val_color if val_color else ACCENT, align=PP_ALIGN.CENTER)
    add_textbox(slide, left + Inches(0.1), top + Inches(1.0),
                width - Inches(0.2), Inches(0.4),
                sub, 8, color=MID_TEXT, align=PP_ALIGN.CENTER)


# ── Chart helpers ─────────────────────────────────────────────────────────────

def style_chart(chart, series_colors=None):
    """Apply series colors and axis font styling to a chart."""
    chart.has_legend = False

    # Category axis
    try:
        ca = chart.category_axis
        ca.tick_labels.font.size = Pt(8)
        ca.tick_labels.font.name = "Calibri"
        ca.tick_labels.font.bold = False
    except Exception:
        pass

    # Value axis
    try:
        va = chart.value_axis
        va.tick_labels.font.size = Pt(8)
        va.tick_labels.font.name = "Calibri"
        va.tick_labels.font.bold = False
    except Exception:
        pass

    # Series colors
    if series_colors:
        for i, color in enumerate(series_colors):
            if i < len(chart.series):
                s = chart.series[i]
                s.format.fill.solid()
                s.format.fill.fore_color.rgb = color
                s.format.line.color.rgb = color


def add_column_chart(slide, left, top, width, height, categories, values,
                     series_name="", series_color=None, point_colors=None):
    cd = ChartData()
    cd.categories = categories
    cd.add_series(series_name, values)
    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, left, top, width, height, cd)
    chart = graphic_frame.chart
    style_chart(chart, series_colors=[series_color or ACCENT])

    # Per-point colors (e.g. red for negative GDP)
    if point_colors:
        series = chart.series[0]
        for idx, col in point_colors.items():
            pt = series.points[idx]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = col

    chart.plots[0].gap_width = 80
    return chart


def add_line_chart(slide, left, top, width, height, categories, values,
                   series_name="", series_color=None):
    cd = ChartData()
    cd.categories = categories
    cd.add_series(series_name, values)
    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE, left, top, width, height, cd)
    chart = graphic_frame.chart
    style_chart(chart, series_colors=[series_color or ACCENT2])

    series = chart.series[0]
    series.smooth = True
    series.format.line.width = Pt(2.0)
    series.format.line.color.rgb = series_color or ACCENT2
    series.marker.format.fill.solid()
    series.marker.format.fill.fore_color.rgb = series_color or ACCENT2
    series.marker.format.line.color.rgb = series_color or ACCENT2

    return chart


# ── Presentation setup ────────────────────────────────────────────────────────

prs = Presentation()
prs.slide_width  = SLIDE_W
prs.slide_height = SLIDE_H
blank = prs.slide_layouts[6]


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — Title / Big Idea
# ══════════════════════════════════════════════════════════════════════════════
s1 = prs.slides.add_slide(blank)
set_bg(s1, DARK_BG)
add_rect(s1, Inches(0), Inches(0), Inches(0.12), SLIDE_H, ACCENT)
add_rect(s1, Inches(0.12), Inches(0), SLIDE_W - Inches(0.12), Inches(0.06), DIVIDER)

add_textbox(s1, Inches(0.5), Inches(0.35), Inches(8), Inches(0.4),
            "ZAMBIA  ·  MACROECONOMIC ANALYSIS  ·  2015–2025", 9, color=ACCENT, bold=True)

add_multiline(s1, Inches(0.5), Inches(1.0), Inches(12.4), Inches(2.2),
              [("Zambia's Debt Restructuring", True, LIGHT_TEXT),
               ("Converted a Commodity Windfall", True, LIGHT_TEXT),
               ("Into a Durable Recovery", True, ACCENT)],
              size=38, align=PP_ALIGN.LEFT)

add_bar(s1, Inches(0.5), Inches(3.2), Inches(9))
add_textbox(s1, Inches(0.5), Inches(3.35), Inches(1.8), Inches(0.35),
            "BIG IDEA", 9, bold=True, color=ACCENT)
add_textbox(s1, Inches(0.5), Inches(3.7), Inches(11.8), Inches(1.5),
            "Zambia's G20 Common Framework debt restructuring converted a copper-driven "
            "fiscal windfall into a durable institutional framework — slashing PPG debt "
            "service from 12.5 % to 3 % of exports, sustaining 5 %+ GDP growth, and "
            "attracting record FDI of 9.32 % of GDP in 2024 — proving that commodity "
            "luck alone cannot anchor a recovery without credible institutional reform.",
            13, color=LIGHT_TEXT, italic=True)

add_textbox(s1, Inches(0.5), Inches(6.9), Inches(6), Inches(0.4),
            "Boldwin Mweemba  ·  Data Analyst  ·  2026", 9, color=MID_TEXT)
add_textbox(s1, Inches(12.5), Inches(6.9), Inches(0.7), Inches(0.4),
            "01", 9, color=MID_TEXT, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — The Problem  (+ External Debt bar chart)
# ══════════════════════════════════════════════════════════════════════════════
s2 = prs.slides.add_slide(blank)
set_bg(s2, DARK_BG)
add_rect(s2, Inches(0), Inches(0), Inches(0.12), SLIDE_H, RED_WARN)
add_rect(s2, Inches(0.12), Inches(0), SLIDE_W - Inches(0.12), Inches(0.06), DIVIDER)

add_textbox(s2, Inches(0.5), Inches(0.25), Inches(8), Inches(0.4),
            "3-MINUTE STORY  ·  PART 1 OF 4  ·  THE PROBLEM", 9, color=RED_WARN, bold=True)
add_multiline(s2, Inches(0.5), Inches(0.7), Inches(12), Inches(0.9),
              [("The Debt Spiral: How Zambia Reached Default", True, LIGHT_TEXT)], size=26)
add_bar(s2, Inches(0.5), Inches(1.65), Inches(5.5), color=RED_WARN)

add_multiline(s2, Inches(0.5), Inches(1.8), Inches(6.0), Inches(2.8),
              [("Between 2015 and 2019, Zambia's external debt tripled — from $12.3 B to "
                "$29.1 B — as the government borrowed against future copper revenues. "
                "When COVID-19 struck in 2020, revenues collapsed. In November 2020 "
                "Zambia became the first African sovereign pandemic-era default, missing "
                "a $42.5 M Eurobond coupon payment.", False, LIGHT_TEXT),
               ("", False, LIGHT_TEXT),
               ("  •  Copper-dependent export base: no buffer against price shocks", False, MID_TEXT),
               ("  •  Debt service consumed 12.5 % of all export earnings", False, MID_TEXT)],
              size=11)

# PPG stat card
add_rect(s2, Inches(0.5), Inches(5.2), Inches(5.8), Inches(0.95), DIVIDER)
add_textbox(s2, Inches(0.65), Inches(5.3), Inches(5.5), Inches(0.35),
            "PPG DEBT SERVICE AT PEAK", 9, color=MID_TEXT)
add_textbox(s2, Inches(0.65), Inches(5.65), Inches(2.0), Inches(0.4),
            "12.5%", 20, bold=True, color=RED_WARN)
add_textbox(s2, Inches(2.8), Inches(5.72), Inches(3.2), Inches(0.32),
            "of total export earnings", 10, color=MID_TEXT)

# External Debt bar chart
add_textbox(s2, Inches(6.6), Inches(1.3), Inches(6.5), Inches(0.35),
            "EXTERNAL DEBT  ($ billions)  2015 – 2024", 9, bold=True, color=MID_TEXT)

debt_colors = {
    4: RED_WARN,   # 2019 peak
    5: RED_WARN,   # 2020 default year
}
add_column_chart(s2, Inches(6.6), Inches(1.65), Inches(6.5), Inches(4.8),
                 YEARS_ALL, EXT_DEBT,
                 series_color=ACCENT, point_colors=debt_colors)

add_textbox(s2, Inches(12.5), Inches(6.9), Inches(0.7), Inches(0.4),
            "02", 9, color=MID_TEXT, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — The Inflection  (+ Current Account line chart)
# ══════════════════════════════════════════════════════════════════════════════
s3 = prs.slides.add_slide(blank)
set_bg(s3, DARK_BG)
add_rect(s3, Inches(0), Inches(0), Inches(0.12), SLIDE_H, ACCENT)
add_rect(s3, Inches(0.12), Inches(0), SLIDE_W - Inches(0.12), Inches(0.06), DIVIDER)

add_textbox(s3, Inches(0.5), Inches(0.25), Inches(9), Inches(0.4),
            "3-MINUTE STORY  ·  PART 2 OF 4  ·  THE INFLECTION", 9, color=ACCENT, bold=True)
add_multiline(s3, Inches(0.5), Inches(0.7), Inches(12), Inches(0.9),
              [("Copper Created the Breathing Room", True, LIGHT_TEXT)], size=26)
add_bar(s3, Inches(0.5), Inches(1.65), Inches(5.5))

add_multiline(s3, Inches(0.5), Inches(1.8), Inches(6.0), Inches(3.0),
              [("Before the restructuring deal was signed, a copper price supercycle "
                "transformed Zambia's external position. The current account swung from "
                "a -3.6 % deficit in 2015 to a +11.9 % surplus in 2021 — purely on the "
                "back of commodity revenue.", False, LIGHT_TEXT),
               ("", False, LIGHT_TEXT),
               ("Simultaneously, the IMF approved a $1.3 B Extended Credit Facility (ECF) "
                "programme, providing a credible fiscal anchor.", False, LIGHT_TEXT),
               ("", False, LIGHT_TEXT),
               ("KEY INSIGHT: The copper windfall was necessary but not sufficient. "
                "Without institutional reform, history suggested it would be spent — not saved.",
                True, ACCENT)],
              size=11)

# Stat cards
for i, (m, v, s, vc) in enumerate([
    ("Current Account\n2015", "-3.6%", "% of GDP  (deficit)", RED_WARN),
    ("Current Account\n2021", "+11.9%", "% of GDP  (surplus)", ACCENT2),
    ("IMF ECF\nApproved", "$1.3 B", "2022 facility", ACCENT),
]):
    stat_card(s3,
              Inches(0.5) + i * Inches(2.0),
              Inches(5.2), Inches(1.85), Inches(1.35),
              m, v, s, vc)

# Current Account line chart
add_textbox(s3, Inches(6.6), Inches(1.3), Inches(6.5), Inches(0.35),
            "CURRENT ACCOUNT BALANCE  (% of GDP)  2015 – 2024", 9, bold=True, color=MID_TEXT)
add_line_chart(s3, Inches(6.6), Inches(1.65), Inches(6.5), Inches(4.8),
               YEARS_ALL, CURR_ACCT,
               series_color=ACCENT2)

add_textbox(s3, Inches(12.5), Inches(6.9), Inches(0.7), Inches(0.4),
            "03", 9, color=MID_TEXT, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — The Structural Fix  (+ GDP Growth chart)
# ══════════════════════════════════════════════════════════════════════════════
s4 = prs.slides.add_slide(blank)
set_bg(s4, DARK_BG)
add_rect(s4, Inches(0), Inches(0), Inches(0.12), SLIDE_H, ACCENT2)
add_rect(s4, Inches(0.12), Inches(0), SLIDE_W - Inches(0.12), Inches(0.06), DIVIDER)

add_textbox(s4, Inches(0.5), Inches(0.25), Inches(9), Inches(0.4),
            "3-MINUTE STORY  ·  PART 3 OF 4  ·  THE STRUCTURAL FIX", 9, color=ACCENT2, bold=True)
add_multiline(s4, Inches(0.5), Inches(0.7), Inches(12), Inches(0.9),
              [("Debt Restructuring Locked In the Gains", True, LIGHT_TEXT)], size=26)
add_bar(s4, Inches(0.5), Inches(1.65), Inches(5.5), color=ACCENT2)

add_multiline(s4, Inches(0.5), Inches(1.8), Inches(6.0), Inches(3.5),
              [("In November 2023, Zambia finalised the G20 Common Framework restructuring "
                "deal — the first successful completion under this framework.", False, LIGHT_TEXT),
               ("", False, LIGHT_TEXT),
               ("  •  PPG debt service: 12.5 % → 3 % of exports", False, ACCENT2),
               ("  •  GDP growth held at 5.2–5.4 % as copper prices cooled", False, ACCENT2),
               ("  •  FDI hit a decade-high of 9.32 % of GDP in 2024", False, ACCENT2),
               ("", False, LIGHT_TEXT),
               ("Key distinction: copper created fiscal space; restructuring institutionalised it.",
                True, LIGHT_TEXT)],
              size=11)

# Before/after table (condensed)
add_textbox(s4, Inches(0.5), Inches(5.2), Inches(5.9), Inches(0.3),
            "BEFORE  →  AFTER RESTRUCTURING", 9, bold=True, color=MID_TEXT)
for i, (metric, before, after) in enumerate([
    ("PPG Debt Service / Exports", "12.5 %", "3.0 %"),
    ("FDI Inflows (% GDP)", "1.8 %  (2020)", "9.32 %  (2024)"),
]):
    y = Inches(5.55) + i * Inches(0.48)
    add_rect(s4, Inches(0.5), y, Inches(5.85), Inches(0.42), DIVIDER)
    add_textbox(s4, Inches(0.6), y + Inches(0.05), Inches(2.2), Inches(0.3),
                metric, 8, color=MID_TEXT)
    add_textbox(s4, Inches(2.85), y + Inches(0.05), Inches(1.2), Inches(0.3),
                before, 10, bold=True, color=RED_WARN)
    add_textbox(s4, Inches(4.1), y + Inches(0.05), Inches(2.1), Inches(0.3),
                "→  " + after, 10, bold=True, color=ACCENT2)

# GDP Growth column chart
add_textbox(s4, Inches(6.6), Inches(1.3), Inches(6.5), Inches(0.35),
            "GDP GROWTH  (% annual)  2015 – 2024", 9, bold=True, color=MID_TEXT)
add_textbox(s4, Inches(6.6), Inches(1.62), Inches(5.5), Inches(0.28),
            "Red = contraction  ·  Green = growth post-restructuring", 8, color=MID_TEXT)

gdp_point_colors = {
    4: RGBColor(0xFF, 0xA5, 0x00),   # 2019 (slowdown) — orange
    5: RED_WARN,                      # 2020 (COVID contraction) — red
    6: ACCENT2,                       # 2021 — green
    7: ACCENT2,                       # 2022 — green
    8: ACCENT2,                       # 2023 — green
    9: ACCENT2,                       # 2024 — green
}
add_column_chart(s4, Inches(6.6), Inches(1.9), Inches(6.5), Inches(4.55),
                 YEARS_ALL, GDP_GROWTH,
                 series_color=ACCENT, point_colors=gdp_point_colors)

add_textbox(s4, Inches(12.5), Inches(6.9), Inches(0.7), Inches(0.4),
            "04", 9, color=MID_TEXT, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — The Evidence: Charts  (FDI + Debt Service visual)
# ══════════════════════════════════════════════════════════════════════════════
s5 = prs.slides.add_slide(blank)
set_bg(s5, DARK_BG)
add_rect(s5, Inches(0), Inches(0), Inches(0.12), SLIDE_H, ACCENT)
add_rect(s5, Inches(0.12), Inches(0), SLIDE_W - Inches(0.12), Inches(0.06), DIVIDER)

add_textbox(s5, Inches(0.5), Inches(0.25), Inches(9), Inches(0.4),
            "3-MINUTE STORY  ·  PART 4 OF 4  ·  THE EVIDENCE", 9, color=ACCENT, bold=True)
add_multiline(s5, Inches(0.5), Inches(0.7), Inches(12), Inches(0.9),
              [("FDI & Debt Service: The Recovery in Numbers", True, LIGHT_TEXT)], size=26)
add_bar(s5, Inches(0.5), Inches(1.65), Inches(12.5))

# FDI chart (left)
add_textbox(s5, Inches(0.5), Inches(1.85), Inches(6.2), Inches(0.32),
            "FDI NET INFLOWS  (% of GDP)  2015 – 2024", 9, bold=True, color=MID_TEXT)
fdi_point_colors = {
    5: RED_WARN,     # 2020 — decade low
    9: ACCENT2,      # 2024 — decade high
}
add_column_chart(s5, Inches(0.5), Inches(2.18), Inches(6.2), Inches(4.55),
                 YEARS_ALL, FDI,
                 series_color=ACCENT, point_colors=fdi_point_colors)

# Debt service comparison (right) — bar chart with 2 categories
add_textbox(s5, Inches(7.1), Inches(1.85), Inches(6.0), Inches(0.32),
            "PPG DEBT SERVICE  (% of exports)  BEFORE vs AFTER DEAL", 9, bold=True, color=MID_TEXT)

cd2 = ChartData()
cd2.categories = ['Pre-restructuring\n(2020)', 'Post-restructuring\n(2024)']
cd2.add_series("Debt Service % Exports", [12.5, 3.0])
gf2 = s5.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(7.1), Inches(2.18), Inches(6.0), Inches(4.55), cd2)
chart2 = gf2.chart
style_chart(chart2, series_colors=[ACCENT])
chart2.series[0].points[0].format.fill.solid()
chart2.series[0].points[0].format.fill.fore_color.rgb = RED_WARN
chart2.series[0].points[1].format.fill.solid()
chart2.series[0].points[1].format.fill.fore_color.rgb = ACCENT2
chart2.plots[0].gap_width = 100

add_textbox(s5, Inches(0.5), Inches(6.82), Inches(12), Inches(0.4),
            "Sources: World Bank WDI · IMF WEO · Bank of Zambia · Ministry of Finance — Report on the Economy 2025",
            8, color=MID_TEXT)
add_textbox(s5, Inches(12.5), Inches(6.9), Inches(0.7), Inches(0.4),
            "05", 9, color=MID_TEXT, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — KPI Summary Cards
# ══════════════════════════════════════════════════════════════════════════════
s6 = prs.slides.add_slide(blank)
set_bg(s6, DARK_BG)
add_rect(s6, Inches(0), Inches(0), Inches(0.12), SLIDE_H, ACCENT)
add_rect(s6, Inches(0.12), Inches(0), SLIDE_W - Inches(0.12), Inches(0.06), DIVIDER)

add_textbox(s6, Inches(0.5), Inches(0.25), Inches(9), Inches(0.4),
            "KEY METRICS  ·  THE NUMBERS AT A GLANCE", 9, color=ACCENT, bold=True)
add_multiline(s6, Inches(0.5), Inches(0.7), Inches(12), Inches(0.9),
              [("The Numbers Confirm the Verdict", True, LIGHT_TEXT)], size=26)
add_bar(s6, Inches(0.5), Inches(1.65), Inches(12.5))

kpis = [
    ("External Debt Peak", "$29.1 B", "2019  (tripled in 4 years)", RED_WARN),
    ("Current Account Swing", "+15.5 pp", "-3.6 % → +11.9 % of GDP", ACCENT2),
    ("IMF ECF Approved", "$1.3 B", "2022 facility", ACCENT),
    ("PPG Debt Service (post-deal)", "3.0 %", "down from 12.5 % of exports", ACCENT2),
    ("GDP Growth (2024)", "5.3 %", "sustained as Cu prices cooled", ACCENT2),
    ("FDI Inflows (2024)", "9.32 %", "decade high  ·  % of GDP", ACCENT),
]
for i, (m, v, s, vc) in enumerate(kpis):
    col = i % 3
    row = i // 3
    x = Inches(0.5) + col * (Inches(4.1) + Inches(0.12))
    y = Inches(2.05) + row * (Inches(1.55) + Inches(0.3))
    stat_card(s6, x, y, Inches(4.1), Inches(1.55), m, v, s, vc)

add_textbox(s6, Inches(0.5), Inches(6.55), Inches(12), Inches(0.4),
            "Sources: World Bank WDI · IMF WEO · Bank of Zambia · Ministry of Finance — Report on the Economy 2025",
            8, color=MID_TEXT)
add_textbox(s6, Inches(12.5), Inches(6.9), Inches(0.7), Inches(0.4),
            "06", 9, color=MID_TEXT, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — The Verdict + Call to Action
# ══════════════════════════════════════════════════════════════════════════════
s7 = prs.slides.add_slide(blank)
set_bg(s7, DARK_BG)
add_rect(s7, Inches(0), Inches(0), Inches(0.12), SLIDE_H, ACCENT2)
add_rect(s7, Inches(0.12), Inches(0), SLIDE_W - Inches(0.12), Inches(0.06), DIVIDER)

add_textbox(s7, Inches(0.5), Inches(0.25), Inches(9), Inches(0.4),
            "CONCLUSION  ·  THE VERDICT", 9, color=ACCENT2, bold=True)
add_multiline(s7, Inches(0.5), Inches(0.7), Inches(12.5), Inches(1.0),
              [("Both Were Necessary — Neither Was Sufficient Alone", True, LIGHT_TEXT)],
              size=26)
add_bar(s7, Inches(0.5), Inches(1.75), Inches(12.3), color=ACCENT2)

for side, head, body, col in [
    (0, "COPPER PROVIDED THE WINDFALL",
     "The global copper price supercycle of 2021 handed Zambia a fiscal windfall: "
     "the current account swung by +15.5 pp and exports surged. This created the "
     "fiscal space that made restructuring politically viable and economically credible.\n\n"
     "Without the copper boom, there would have been nothing to restructure toward — "
     "and no creditors willing to take a haircut on a country with no revenue trajectory.",
     ACCENT),
    (1, "RESTRUCTURING MADE IT DURABLE",
     "The G20 Common Framework deal converted a temporary commodity windfall into a "
     "permanent reduction in debt service obligations. PPG debt service fell from "
     "12.5 % to 3 % of exports — freeing fiscal resources for productive investment.\n\n"
     "GDP growth of 5.2–5.4 % persisted even as copper prices moderated, and FDI "
     "inflows hit a decade-high of 9.32 % of GDP — confirming that investors saw "
     "institutional, not just commodity, credibility.",
     ACCENT2),
]:
    x = Inches(0.5) + side * Inches(6.5)
    add_textbox(s7, x, Inches(2.0), Inches(6.0), Inches(0.4), head, 11, bold=True, color=col)
    add_textbox(s7, x, Inches(2.5), Inches(6.15), Inches(3.0), body, 11, color=LIGHT_TEXT)

add_rect(s7, Inches(6.55), Inches(1.9), Inches(0.04), Inches(3.7), DIVIDER)

add_rect(s7, Inches(0.5), Inches(5.75), Inches(12.3), Inches(1.05), DIVIDER)
add_textbox(s7, Inches(0.65), Inches(5.82), Inches(1.2), Inches(0.3),
            "BIG IDEA", 8, bold=True, color=ACCENT)
add_textbox(s7, Inches(0.65), Inches(6.12), Inches(12.0), Inches(0.5),
            "Zambia's recovery proves that commodity luck and institutional reform are complements, "
            "not substitutes — copper created the window; the G20 deal made it a door.",
            13, bold=True, color=LIGHT_TEXT, italic=True)

add_textbox(s7, Inches(12.5), Inches(6.9), Inches(0.7), Inches(0.4),
            "07", 9, color=MID_TEXT, align=PP_ALIGN.RIGHT)


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — Appendix: 3-Minute Story & Big Idea
# ══════════════════════════════════════════════════════════════════════════════
s8 = prs.slides.add_slide(blank)
set_bg(s8, DARK_BG)
add_rect(s8, Inches(0), Inches(0), Inches(0.12), SLIDE_H, MID_TEXT)
add_rect(s8, Inches(0.12), Inches(0), SLIDE_W - Inches(0.12), Inches(0.06), DIVIDER)

add_textbox(s8, Inches(0.5), Inches(0.25), Inches(9), Inches(0.4),
            "APPENDIX  ·  3-MINUTE STORY & BIG IDEA", 9, color=MID_TEXT, bold=True)
add_multiline(s8, Inches(0.5), Inches(0.65), Inches(12), Inches(0.7),
              [("Storytelling with Data — Written Narrative", True, LIGHT_TEXT)], size=20)
add_bar(s8, Inches(0.5), Inches(1.35), Inches(12.3), color=MID_TEXT)

add_textbox(s8, Inches(0.5), Inches(1.5), Inches(6.2), Inches(5.2),
            "Between 2015 and 2019, Zambia borrowed heavily against future copper revenues, "
            "tripling its external debt from $12.3 B to $29.1 B. When COVID-19 struck in 2020, "
            "revenues collapsed and Zambia became the first African sovereign pandemic-era default. "
            "The question we needed to answer: was the recovery that followed real — and sustainable?\n\n"
            "In 2021, a global copper price supercycle delivered a fiscal windfall. The current "
            "account swung 15.5 percentage points — from -3.6 % to +11.9 % of GDP — before the "
            "restructuring deal had even been negotiated. Copper created the breathing room.\n\n"
            "But breathing room is not the same as structural reform. In November 2023, Zambia "
            "finalised the G20 Common Framework deal, becoming the first country to complete the "
            "process. PPG debt service fell from 12.5 % to 3 % of exports; GDP growth held at "
            "5.2–5.4 % even as copper prices cooled; FDI reached a decade-high 9.32 % of GDP.\n\n"
            "The verdict: both were necessary. Commodity luck created the window. Institutional "
            "reform made it a door. Zambia's challenge now is to diversify beyond copper — into "
            "green energy, agriculture, and manufacturing — before the next downcycle.",
            10.5, color=LIGHT_TEXT)

add_rect(s8, Inches(7.0), Inches(1.5), Inches(5.9), Inches(2.4), DIVIDER)
add_textbox(s8, Inches(7.15), Inches(1.6), Inches(3), Inches(0.32),
            "BIG IDEA", 9, bold=True, color=ACCENT)
add_textbox(s8, Inches(7.15), Inches(2.0), Inches(5.6), Inches(1.75),
            "Zambia's G20 Common Framework debt restructuring converted a copper-driven "
            "fiscal windfall into a durable institutional framework — slashing PPG debt "
            "service from 12.5 % to 3 % of exports, sustaining 5 %+ GDP growth, and "
            "attracting record FDI of 9.32 % of GDP in 2024 — proving that commodity "
            "luck alone cannot anchor a recovery without credible institutional reform.",
            11, color=LIGHT_TEXT, italic=True)

for i, txt in enumerate(["✓  Unique point of view", "✓  Conveys what's at stake",
                          "✓  Complete sentence"]):
    add_textbox(s8, Inches(7.15), Inches(4.05) + i * Inches(0.38),
                Inches(5.6), Inches(0.35), txt, 10, color=ACCENT2, bold=True)

add_textbox(s8, Inches(12.5), Inches(6.9), Inches(0.7), Inches(0.4),
            "08", 9, color=MID_TEXT, align=PP_ALIGN.RIGHT)


# ── Save ──────────────────────────────────────────────────────────────────────
out_path = r"c:\Users\bmweemba\Documents\Boldwin-Mweemba-Zambia Economic recovery\Zambia_Economic_Recovery_Presentation.pptx"
prs.save(out_path)
print(f"Saved: {out_path}")
