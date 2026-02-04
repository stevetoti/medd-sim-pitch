#!/usr/bin/env python3
"""
MEDD SIM Pitch Deck - PowerPoint Generator (V5 Dark Theme with Images)
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import urllib.request
import os

# Brand colors - DARK THEME
MEDD_GREEN = RGBColor(0x0D, 0x6B, 0x56)
MEDD_GREEN_LIGHT = RGBColor(0x10, 0xA3, 0x7F)
MEDD_BG = RGBColor(0x1A, 0x1A, 0x1A)
MEDD_CARD = RGBColor(0x2A, 0x2A, 0x2A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MUTED = RGBColor(0xA0, 0xA0, 0xA0)

def set_slide_background(slide, color):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_text(slide, text, top, left, width, height, font_size, color, bold=False, align='left'):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = "Arial"
    if align == 'center':
        p.alignment = PP_ALIGN.CENTER
    return txBox

def add_card(slide, title, body, left, top, width=3.5, height=1.8, title_color=MEDD_GREEN_LIGHT):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = MEDD_CARD
    shape.line.fill.background()
    add_text(slide, title, top + 0.2, left + 0.2, width - 0.4, 0.4, 13, title_color, True)
    add_text(slide, body, top + 0.6, left + 0.2, width - 0.4, 1, 10, MUTED)

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    # ===== SLIDE 1: Hero =====
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide1, MEDD_BG)
    add_text(slide1, "medd sim", 2.2, 0, 13.333, 1, 64, MEDD_GREEN_LIGHT, True, 'center')
    add_text(slide1, "Giving your team the tools to practice\nthe moments that matter, before they matter.", 4, 1.5, 10.333, 1.5, 26, WHITE, False, 'center')
    
    # ===== SLIDE 2: Vision =====
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide2, MEDD_BG)
    add_text(slide2, "Our Vision", 0.8, 0.8, 6, 0.4, 16, MEDD_GREEN_LIGHT)
    add_text(slide2, "Make deliberate practice\na daily norm.", 1.3, 0.8, 6, 1.2, 36, WHITE, True)
    add_text(slide2, "MEDD Sim is the simulation studio you control.\nBuild any AI-powered coach, role-play, or examiner in minutes.", 3, 0.8, 5.5, 1.2, 14, MUTED)
    # Stats
    add_text(slide2, "73%", 4.5, 0.8, 2, 0.8, 44, MEDD_GREEN_LIGHT, True)
    add_text(slide2, "avoid role-play", 5.3, 0.8, 2, 0.4, 12, MUTED, False, 'center')
    add_text(slide2, "4x", 4.5, 3.5, 2, 0.8, 44, MEDD_GREEN_LIGHT, True)
    add_text(slide2, "faster learning", 5.3, 3.5, 2, 0.4, 12, MUTED, False, 'center')
    # Placeholder for image
    shape = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7), Inches(1.5), Inches(5.5), Inches(4.5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = MEDD_CARD
    shape.line.fill.background()
    add_text(slide2, "📸 Team Collaboration Image", 3.5, 7.5, 4.5, 0.5, 14, MUTED, False, 'center')
    
    # ===== SLIDE 3: Problems =====
    slide3 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide3, MEDD_BG)
    add_text(slide3, "The problems we face", 0.5, 0.8, 11, 0.7, 32, WHITE, True)
    
    # Problem cards
    cards = [
        ("Role-Play Anxiety", "Nobody enjoys role-plays, so they become filler sessions or dreaded training.", 0.8),
        ("Time-Starved Managers", "When managers are stretched thin, coaching is the first thing dropped.", 4.8),
        ("Fear of Exposure", "Professionals love coaching results — just don't love feeling exposed.", 8.8),
    ]
    for title, desc, x in cards:
        shape = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(2), Inches(3.8), Inches(4))
        shape.fill.solid()
        shape.fill.fore_color.rgb = MEDD_CARD
        shape.line.fill.background()
        add_text(slide3, "📸", 2.3, x + 1.5, 1, 1, 48, MUTED, False, 'center')
        add_text(slide3, title, 4, x + 0.3, 3.2, 0.5, 14, WHITE, True)
        add_text(slide3, desc, 4.6, x + 0.3, 3.2, 1.2, 11, MUTED)
    
    # ===== SLIDE 4: Solution =====
    slide4 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide4, MEDD_BG)
    add_text(slide4, "How we fix it", 0.5, 0.8, 6, 0.4, 16, MEDD_GREEN_LIGHT)
    add_text(slide4, "A behaviour rehearsal engine", 1, 0.8, 6, 0.7, 32, WHITE, True)
    add_text(slide4, "Build scenarios, practice in private, and make it happen in the real world.", 1.8, 0.8, 5.5, 0.5, 14, MUTED)
    
    solutions = [
        ("Total Control", "User-generated agents from templates or scratch"),
        ("Content Studio", "Generate resources with a click"),
        ("BYO Rubric", "Standardise feedback to your framework"),
        ("Learning Pathways", "Structured training for any goal"),
    ]
    y = 2.8
    for i, (title, desc) in enumerate(solutions):
        x = 0.8 if i % 2 == 0 else 4
        add_card(slide4, title, desc, x, y, 3, 1.3)
        if i % 2 == 1:
            y += 1.6
    
    # Image placeholder
    shape = slide4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.5), Inches(1.5), Inches(5), Inches(4.5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = MEDD_CARD
    shape.line.fill.background()
    add_text(slide4, "📸 Medical Professional Image", 3.5, 8, 4, 0.5, 14, MUTED, False, 'center')
    
    # ===== SLIDE 5: Use Cases =====
    slide5 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide5, MEDD_BG)
    add_text(slide5, "Build any coaching or roleplay experience", 0.5, 0.8, 11, 0.7, 28, WHITE, True)
    
    cases = [
        ("👔 Roleplay Customer", "Personas, objections, buying motives"),
        ("💬 Coach Any Situation", "Sales, business, HR templates"),
        ("📚 Workforce Readiness", "Orientation, compliance training"),
        ("🏥 Patient Case Study", "AI patient simulations for CPD"),
        ("📋 Examiner Mode", "OSCE clinical assessments"),
        ("🎯 User-Created Sims", "Personalized team coaching"),
    ]
    
    x, y = 0.8, 1.8
    for i, (title, desc) in enumerate(cases):
        add_card(slide5, title, desc, x, y, 3.8, 1.3, WHITE)
        x += 4.1
        if i == 2:
            x, y = 0.8, 3.4
    
    # ===== SLIDE 6: How It Works =====
    slide6 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide6, MEDD_BG)
    add_text(slide6, "How It Works", 0.5, 0.8, 11, 0.7, 32, WHITE, True)
    add_text(slide6, "From signup to simulation in minutes", 1.2, 0.8, 11, 0.4, 14, MUTED)
    
    steps = [("1", "Register", "Access dashboard, select package"),
             ("2", "Onboard", "Bring team, assign roles"),
             ("3", "Create", "Build agents and pathways"),
             ("4", "Practice", "Rehearse, feedback, improve")]
    
    x = 0.8
    for num, title, desc in steps:
        circle = slide6.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 1), Inches(2.5), Inches(0.9), Inches(0.9))
        circle.fill.solid()
        circle.fill.fore_color.rgb = MEDD_GREEN
        circle.line.fill.background()
        add_text(slide6, num, 2.6, x + 1.2, 0.6, 0.6, 28, WHITE, True)
        add_text(slide6, title, 3.8, x, 3, 0.5, 18, WHITE, True, 'center')
        add_text(slide6, desc, 4.4, x, 3, 0.8, 11, MUTED, False, 'center')
        x += 3.1
    
    # ===== SLIDE 7: Security =====
    slide7 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide7, MEDD_BG)
    add_text(slide7, "Security and governance", 0.5, 0.8, 11, 0.7, 32, WHITE, True)
    add_text(slide7, "Private by design — secure by default", 1.2, 0.8, 11, 0.4, 14, MUTED)
    
    security = [
        "🛡️ Data boundary — no model training",
        "🖥️ Customer isolation — separate instance",
        "🔐 Role-based access control",
        "📋 Full audit logs",
        "📄 Rubric versioning",
        "📤 Exportable evidence",
    ]
    
    x, y = 0.8, 2.2
    for i, item in enumerate(security):
        shape = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(5.8), Inches(0.7))
        shape.fill.solid()
        shape.fill.fore_color.rgb = MEDD_CARD
        shape.line.fill.background()
        add_text(slide7, item, y + 0.15, x + 0.3, 5.2, 0.5, 13, WHITE)
        x = 6.8 if x == 0.8 else 0.8
        if i % 2 == 1:
            y += 1
    
    # ===== SLIDE 8: Pricing =====
    slide8 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide8, MEDD_BG)
    add_text(slide8, "Pricing", 0.3, 0.8, 11, 0.6, 28, WHITE, True)
    
    # Essential
    shape = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.3), Inches(3.8), Inches(5.2))
    shape.fill.solid()
    shape.fill.fore_color.rgb = MEDD_CARD
    shape.line.fill.background()
    add_text(slide8, "ESSENTIAL", 1.5, 1, 3.4, 0.4, 11, MEDD_GREEN_LIGHT, True, 'center')
    add_text(slide8, "$39", 2, 1, 3.4, 0.7, 40, WHITE, True, 'center')
    add_text(slide8, "/user/mo", 2.7, 1, 3.4, 0.3, 11, MUTED, False, 'center')
    add_text(slide8, "✓ 10 min video / 60 min audio\n✓ Unlimited agent types\n✓ Learning pathways\n✓ Templates + rubrics\n✓ Email support", 3.3, 1.2, 3.2, 2.5, 10, MUTED)
    
    # Professional
    shape = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.8), Inches(1.3), Inches(3.8), Inches(5.2))
    shape.fill.solid()
    shape.fill.fore_color.rgb = MEDD_CARD
    shape.line.color.rgb = MEDD_GREEN
    shape.line.width = Pt(2)
    add_text(slide8, "★ POPULAR", 1.1, 5.8, 1.8, 0.3, 9, MEDD_GREEN_LIGHT, True, 'center')
    add_text(slide8, "PROFESSIONAL", 1.5, 5, 3.4, 0.4, 11, MEDD_GREEN_LIGHT, True, 'center')
    add_text(slide8, "$79", 2, 5, 3.4, 0.7, 40, WHITE, True, 'center')
    add_text(slide8, "/user/mo", 2.7, 5, 3.4, 0.3, 11, MUTED, False, 'center')
    add_text(slide8, "✓ 25 min video / 90 min audio\n✓ Everything in Essential\n✓ Team analytics + gamification\n✓ Pooled Content Studio\n✓ Priority support", 3.3, 5.2, 3.2, 2.5, 10, MUTED)
    
    # Enterprise
    shape = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.8), Inches(1.3), Inches(3.8), Inches(5.2))
    shape.fill.solid()
    shape.fill.fore_color.rgb = MEDD_CARD
    shape.line.fill.background()
    add_text(slide8, "ENTERPRISE", 1.5, 9, 3.4, 0.4, 11, MEDD_GREEN_LIGHT, True, 'center')
    add_text(slide8, "Custom", 2, 9, 3.4, 0.7, 40, WHITE, True, 'center')
    add_text(slide8, "contact us", 2.7, 9, 3.4, 0.3, 11, MUTED, False, 'center')
    add_text(slide8, "✓ Custom instance\n✓ High pooled allowances\n✓ SSO & governance\n✓ Volume rates\n✓ Phone support", 3.3, 9.2, 3.2, 2.5, 10, MUTED)
    
    # ===== SLIDE 9: CTA =====
    slide9 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide9, MEDD_BG)
    add_text(slide9, "medd sim", 1.8, 0, 13.333, 0.8, 48, MEDD_GREEN_LIGHT, True, 'center')
    add_text(slide9, "Ready to transform how\nyour team practices?", 3, 0, 13.333, 1.2, 36, WHITE, True, 'center')
    add_text(slide9, "Turn high-stakes conversations into rehearsed performances.", 4.8, 0, 13.333, 0.5, 16, MUTED, False, 'center')
    
    # CTA Button
    btn = slide9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.8), Inches(5.5), Inches(3.7), Inches(0.7))
    btn.fill.solid()
    btn.fill.fore_color.rgb = MEDD_GREEN
    btn.line.fill.background()
    add_text(slide9, "Start Free Trial", 5.6, 5.3, 2.8, 0.5, 16, WHITE, True, 'center')
    
    add_text(slide9, "📧 hello@medd.com.au   |   🌐 sim.medd.com.au", 6.5, 0, 13.333, 0.4, 12, MUTED, False, 'center')
    
    # Save
    output_path = "/home/toti/projects/medd-sim-pitch/MEDD-SIM-Pitch-Deck.pptx"
    prs.save(output_path)
    print(f"✅ Presentation saved: {output_path}")
    return output_path

if __name__ == "__main__":
    create_presentation()
