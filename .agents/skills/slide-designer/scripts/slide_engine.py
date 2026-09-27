import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

# ==============================================================================
# Modern Design System (Palette & Fonts)
# ==============================================================================
FONT_HEADING = 'Leelawadee UI'
FONT_BODY = 'Leelawadee UI'

# 60-30-10 Professional Color Harmony
CLR_CANVAS = RGBColor(248, 250, 252)        # #F8FAFC Ultra-light slate canvas
CLR_WHITE = RGBColor(255, 255, 255)         # #FFFFFF Crisp White
CLR_TITLE = RGBColor(15, 23, 42)            # #0F172A Dark Slate 900
CLR_BODY = RGBColor(51, 65, 85)             # #334155 Slate 700
CLR_MUTED = RGBColor(100, 116, 139)         # #64748B Muted Slate 500
CLR_BORDER = RGBColor(226, 232, 240)        # #E2E8F0 Subtle Border Slate 200
CLR_BORDER_STRONG = RGBColor(203, 213, 225) # #CBD5E1

# Vivid Brand & Intent Accents
CLR_EMERALD = RGBColor(16, 149, 106)        # #10956A LINE / Food / Success
CLR_EMERALD_LIGHT = RGBColor(236, 253, 245) # #ECFDF5 Tinted BG
CLR_BLUE = RGBColor(37, 99, 235)            # #2563EB Royal Tech Blue
CLR_BLUE_LIGHT = RGBColor(239, 246, 255)    # #EFF6FF Tinted BG
CLR_ORANGE = RGBColor(234, 88, 12)          # #EA580C KMUTNB Orange / Alert
CLR_ORANGE_LIGHT = RGBColor(255, 247, 237)  # #FFF7ED Tinted BG
CLR_PURPLE = RGBColor(124, 58, 237)         # #7C3AED Innovation Violet
CLR_PURPLE_LIGHT = RGBColor(245, 243, 255)  # #F5F3FF Tinted BG
CLR_RED = RGBColor(220, 38, 38)             # #DC2626 Problem Red
CLR_AMBER = RGBColor(217, 119, 6)           # #D97706 Warning Amber

def add_header(slide, action_title, category_tag, slide_num, total_slides=15, accent_color=CLR_BLUE):
    """
    Creates an executive-grade header with:
    1. A subtle colored category pill/badge at top
    2. A high-contrast, punchy Action Headline
    3. Integrated metadata and subtle divider
    """
    # Header container
    header_box = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.35), Inches(11.733), Inches(0.95)
    )
    header_box.fill.solid()
    header_box.fill.fore_color.rgb = CLR_WHITE
    header_box.line.color.rgb = CLR_BORDER
    header_box.line.width = Pt(1)

    # Accent color indicator bar on left
    stripe = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.35), Inches(0.08), Inches(0.95)
    )
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = accent_color
    stripe.line.fill.background()

    tf = header_box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.3)
    tf.margin_top = Inches(0.06)
    tf.margin_bottom = Inches(0.06)

    # Line 1: Category Tag Pill
    p_tag = tf.paragraphs[0]
    p_tag.text = category_tag.upper()
    p_tag.font.name = FONT_HEADING
    p_tag.font.size = Pt(9.5)
    p_tag.font.bold = True
    p_tag.font.color.rgb = accent_color

    # Line 2: Action Title
    p_title = tf.add_paragraph()
    p_title.text = action_title
    p_title.font.name = FONT_HEADING
    p_title.font.size = Pt(18.5)
    p_title.font.bold = True
    p_title.font.color.rgb = CLR_TITLE

def add_footer(slide, slide_num, total_slides=15):
    """
    Adds a minimalist, clean academic footer with page numbers.
    """
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(7.0), Inches(11.733), Inches(0.015)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = CLR_BORDER
    line.line.fill.background()

    tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.04), Inches(11.733), Inches(0.35))
    tf = tx_box.text_frame
    tf.margin_top = Inches(0)
    tf.margin_bottom = Inches(0)
    tf.margin_left = Inches(0)
    tf.margin_right = Inches(0)

    p = tf.paragraphs[0]
    p.text = "วิชา 080203914 ผู้ประกอบการนวัตกรรม (Innovative Technopreneurs) | คณะ FITM มจพ. ปราจีนบุรี"
    p.font.name = FONT_BODY
    p.font.size = Pt(9.5)
    p.font.color.rgb = CLR_MUTED

    p_num = tf.add_paragraph()
    p_num.alignment = PP_ALIGN.RIGHT
    p_num.text = f"{slide_num:02d} / {total_slides:02d}"
    p_num.font.name = FONT_BODY
    p_num.font.size = Pt(9.5)
    p_num.font.bold = True
    p_num.font.color.rgb = CLR_TITLE

def add_bento_card(slide, left, top, width, height, title, items, badge_text=None, accent_color=CLR_BLUE, bg_color=CLR_WHITE):
    """
    Draws a sleek Bento Card container with:
    - Subtle 1pt border and clean background
    - Optional badge chip on top right or header bar
    - Bold title
    - Scannable, formatted bullet items with bold prefix support
    """
    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = CLR_BORDER
    card.line.width = Pt(1)

    # Accent line at top
    accent_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.04))
    accent_bar.fill.solid()
    accent_bar.fill.fore_color.rgb = accent_color
    accent_bar.line.fill.background()

    # Header section inside card
    title_height = Inches(0.55)
    title_box = slide.shapes.add_textbox(left + Inches(0.18), top + Inches(0.08), width - Inches(0.36), title_height)
    ttf = title_box.text_frame
    ttf.word_wrap = True
    ttf.margin_left = ttf.margin_top = ttf.margin_bottom = ttf.margin_right = 0

    tp = ttf.paragraphs[0]
    tp.text = title
    tp.font.name = FONT_HEADING
    tp.font.size = Pt(13.5)
    tp.font.bold = True
    tp.font.color.rgb = CLR_TITLE

    if badge_text:
        # Mini pill badge
        badge_box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + width - Inches(1.3), top + Inches(0.12), Inches(1.15), Inches(0.28))
        badge_box.fill.solid()
        badge_box.fill.fore_color.rgb = CLR_BLUE_LIGHT if accent_color == CLR_BLUE else CLR_EMERALD_LIGHT
        badge_box.line.fill.background()
        btf = badge_box.text_frame
        btf.vertical_anchor = MSO_ANCHOR.MIDDLE
        bp = btf.paragraphs[0]
        bp.alignment = PP_ALIGN.CENTER
        bp.text = badge_text
        bp.font.name = FONT_HEADING
        bp.font.size = Pt(8.5)
        bp.font.bold = True
        bp.font.color.rgb = accent_color

    # Content body
    content_top = top + title_height + Inches(0.08)
    content_height = height - title_height - Inches(0.18)
    content_box = slide.shapes.add_textbox(left + Inches(0.18), content_top, width - Inches(0.36), content_height)
    ctf = content_box.text_frame
    ctf.word_wrap = True
    ctf.margin_left = ctf.margin_top = ctf.margin_bottom = ctf.margin_right = 0

    for i, itm in enumerate(items):
        p = ctf.paragraphs[0] if i == 0 else ctf.add_paragraph()
        p.space_before = Pt(4)
        if isinstance(itm, tuple):
            prefix, text = itm
            r1 = p.add_run()
            r1.text = prefix + " "
            r1.font.name = FONT_BODY
            r1.font.bold = True
            r1.font.size = Pt(11)
            r1.font.color.rgb = CLR_TITLE

            r2 = p.add_run()
            r2.text = text
            r2.font.name = FONT_BODY
            r2.font.size = Pt(11)
            r2.font.color.rgb = CLR_BODY
        else:
            p.text = itm
            p.font.name = FONT_BODY
            p.font.size = Pt(11)
            p.font.color.rgb = CLR_BODY

def add_metric_hero_card(slide, left, top, width, height, value, label, subtext="", accent_color=CLR_EMERALD):
    """
    Renders a high-impact metric KPI card with huge number and crisp label.
    """
    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = CLR_WHITE
    card.line.color.rgb = CLR_BORDER
    card.line.width = Pt(1)

    top_stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.05))
    top_stripe.fill.solid()
    top_stripe.fill.fore_color.rgb = accent_color
    top_stripe.line.fill.background()

    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.15)
    tf.margin_right = Inches(0.15)

    p_val = tf.paragraphs[0]
    p_val.alignment = PP_ALIGN.CENTER
    p_val.text = value
    p_val.font.name = FONT_HEADING
    p_val.font.size = Pt(30)
    p_val.font.bold = True
    p_val.font.color.rgb = accent_color

    p_lbl = tf.add_paragraph()
    p_lbl.alignment = PP_ALIGN.CENTER
    p_lbl.text = label
    p_lbl.font.name = FONT_HEADING
    p_lbl.font.size = Pt(11)
    p_lbl.font.bold = True
    p_lbl.font.color.rgb = CLR_TITLE

    if subtext:
        p_sub = tf.add_paragraph()
        p_sub.alignment = PP_ALIGN.CENTER
        p_sub.text = subtext
        p_sub.font.name = FONT_BODY
        p_sub.font.size = Pt(9)
        p_sub.font.color.rgb = CLR_MUTED

def add_speaker_notes(slide, timing, script, key_points=None):
    """
    Embeds structured presenter speaker notes in Thai.
    """
    notes_slide = slide.notes_slide
    text_frame = notes_slide.notes_text_frame
    text_frame.text = f"[เป้าหมายเวลา: {timing}]\n\n[บทพูดสำหรับผู้นำเสนอ]:\n{script}\n"
    if key_points:
        text_frame.text += "\n[ประเด็นเน้นย้ำและเตรียมตอบคำถาม]:\n"
        for kp in key_points:
            text_frame.text += f"- {kp}\n"
