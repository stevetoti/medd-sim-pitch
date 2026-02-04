#!/usr/bin/env python3
"""
MEDD SIM Pitch Deck - PowerPoint Generator (Lime Green Style)
Matches Anthony's original presentation style
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# Brand colors - LIME GREEN STYLE
MEDD_LIME = RGBColor(0xC7, 0xF4, 0x64)
MEDD_GREEN = RGBColor(0x0D, 0x6B, 0x56)
MEDD_DARK = RGBColor(0x1A, 0x1A, 0x1A)
MEDD_GRAY = RGBColor(0x4A, 0x4A, 0x4A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

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

def add_card(slide, title, body, left, top, width=3.5, height=2):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = WHITE
    shape.line.fill.background()
    add_text(slide, title, top + 0.2, left + 0.2, width - 0.4, 0.5, 14, MEDD_GREEN, True)
    add_text(slide, body, top + 0.6, left + 0.2, width - 0.4, 1.2, 11, MEDD_GRAY)

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    # ===== SLIDE 1: Hero with Logo =====
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide1, MEDD_LIME)
    
    # Logo placeholder (text version)
    add_text(slide1, "medd sim", 2.5, 0, 13.333, 1, 72, MEDD_GREEN, True, 'center')
    add_text(slide1, "Giving your team the tools to practice the moments that matter,\nbefore they matter.", 4.5, 1.5, 10.333, 1.5, 24, MEDD_DARK, False, 'center')
    
    # ===== SLIDE 2: Vision =====
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide2, MEDD_LIME)
    
    add_text(slide2, "Our Vision", 1, 1, 6, 0.5, 18, MEDD_GRAY)
    add_text(slide2, "Make deliberate practice\na daily norm.", 1.5, 1, 6, 1.5, 40, MEDD_DARK, True)
    add_text(slide2, "MEDD Sim is the simulation studio you control.\nBuild any AI-powered coach, role-play, or examiner in minutes.", 3.5, 1, 5.5, 1.5, 16, MEDD_GRAY)
    
    # ===== SLIDE 3: The Problem =====
    slide3 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide3, MEDD_DARK)
    
    add_text(slide3, "The problems we face", 0.5, 1, 11, 0.8, 36, WHITE, True)
    
    # Problem cards
    add_card(slide3, "Role-Play Anxiety", "Nobody enjoys role-plays, so they become filler sessions or dreaded training.", 0.8, 2.5, 3.8, 2.2)
    add_card(slide3, "Time-Starved Managers", "When managers are stretched thin, coaching is the first thing dropped.", 4.8, 2.5, 3.8, 2.2)
    add_card(slide3, "Fear of Exposure", "Professionals love coaching results — just don't love feeling exposed.", 8.8, 2.5, 3.8, 2.2)
    
    # ===== SLIDE 4: How We Fix It =====
    slide4 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide4, MEDD_LIME)
    
    add_text(slide4, "How we fix it", 0.5, 1, 6, 0.5, 18, MEDD_GRAY)
    add_text(slide4, "A behaviour rehearsal engine", 1, 1, 6, 0.8, 36, MEDD_DARK, True)
    add_text(slide4, "Build scenarios, practice in private, and make it happen in the real world.", 1.8, 1, 6, 0.6, 16, MEDD_GRAY)
    
    add_card(slide4, "Total Control", "User-generated agents from templates or scratch", 1, 3, 3, 1.6)
    add_card(slide4, "Content Studio", "Generate resources with a click", 4.2, 3, 3, 1.6)
    add_card(slide4, "BYO Rubric", "Standardise feedback to your framework", 1, 4.8, 3, 1.6)
    add_card(slide4, "Learning Pathways", "Structured training for any goal", 4.2, 4.8, 3, 1.6)
    
    # ===== SLIDE 5: Build Any Experience =====
    slide5 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide5, MEDD_LIME)
    
    add_text(slide5, "Build any coaching or roleplay experience", 0.5, 1, 11, 0.8, 32, MEDD_DARK, True)
    
    # 6 use case cards (2 rows of 3)
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
        add_card(slide5, title, desc, x, y, 3.8, 1.4)
        x += 4.1
        if i == 2:
            x, y = 0.8, 3.5
    
    # ===== SLIDE 6: How It Works =====
    slide6 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide6, MEDD_LIME)
    
    add_text(slide6, "How It Works", 0.5, 1, 11, 0.8, 36, MEDD_DARK, True)
    
    steps = [
        ("1", "Register", "Access dashboard,\nselect package"),
        ("2", "Onboard", "Bring your team,\nassign roles"),
        ("3", "Create", "Build agents,\npathways, campaigns"),
        ("4", "Practice", "Rehearse, feedback,\nimprove"),
    ]
    
    x = 0.8
    for num, title, desc in steps:
        # Card
        shape = slide6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(2.5), Inches(2.8), Inches(3))
        shape.fill.solid()
        shape.fill.fore_color.rgb = WHITE
        shape.line.fill.background()
        
        # Number circle
        circle = slide6.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x + 1), Inches(2.8), Inches(0.8), Inches(0.8))
        circle.fill.solid()
        circle.fill.fore_color.rgb = MEDD_GREEN
        circle.line.fill.background()
        
        add_text(slide6, num, 2.9, x + 1.25, 0.5, 0.5, 24, WHITE, True)
        add_text(slide6, title, 3.9, x + 0.3, 2.2, 0.5, 18, MEDD_DARK, True, 'center')
        add_text(slide6, desc, 4.5, x + 0.3, 2.2, 1, 12, MEDD_GRAY, False, 'center')
        x += 3.1
    
    # ===== SLIDE 7: Security =====
    slide7 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide7, MEDD_LIME)
    
    add_text(slide7, "Security and governance", 0.5, 1, 11, 0.8, 36, MEDD_DARK, True)
    add_text(slide7, "Private by design — secure by default", 1.2, 1, 11, 0.5, 18, MEDD_GRAY)
    
    security = [
        "🛡️ Data boundary — no model training",
        "🖥️ Customer isolation — separate instance",
        "🔐 Role-based access control",
        "📋 Full audit logs",
        "📄 Rubric versioning",
        "📤 Exportable evidence",
    ]
    
    x, y = 1, 2.2
    for i, item in enumerate(security):
        shape = slide7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(5.5), Inches(0.7))
        shape.fill.solid()
        shape.fill.fore_color.rgb = WHITE
        shape.line.fill.background()
        add_text(slide7, item, y + 0.15, x + 0.3, 5, 0.5, 14, MEDD_DARK)
        x = 6.8 if x == 1 else 1
        if i % 2 == 1:
            y += 1
    
    # ===== SLIDE 8: Pricing =====
    slide8 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide8, MEDD_LIME)
    
    add_text(slide8, "Pricing", 0.3, 1, 11, 0.6, 32, MEDD_DARK, True)
    
    # Essential
    shape = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.2), Inches(3.8), Inches(5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = WHITE
    shape.line.fill.background()
    add_text(slide8, "ESSENTIAL", 1.4, 1, 3.5, 0.4, 12, MEDD_GREEN, True, 'center')
    add_text(slide8, "$39", 1.9, 1, 3.5, 0.6, 40, MEDD_DARK, True, 'center')
    add_text(slide8, "/user/mo", 2.5, 1, 3.5, 0.3, 12, MEDD_GRAY, False, 'center')
    features_e = "✓ 10 min video / 60 min audio\n✓ Unlimited agent types\n✓ Learning pathways\n✓ Templates + rubrics\n✓ Email support"
    add_text(slide8, features_e, 3.2, 1.2, 3.2, 2.5, 11, MEDD_GRAY)
    
    # Professional (featured)
    shape = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.8), Inches(1.2), Inches(3.8), Inches(5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = WHITE
    shape.line.color.rgb = MEDD_GREEN
    shape.line.width = Pt(2)
    add_text(slide8, "★ POPULAR", 1.0, 5.5, 2.5, 0.3, 10, MEDD_GREEN, True, 'center')
    add_text(slide8, "PROFESSIONAL", 1.4, 5, 3.5, 0.4, 12, MEDD_GREEN, True, 'center')
    add_text(slide8, "$79", 1.9, 5, 3.5, 0.6, 40, MEDD_DARK, True, 'center')
    add_text(slide8, "/user/mo", 2.5, 5, 3.5, 0.3, 12, MEDD_GRAY, False, 'center')
    features_p = "✓ 25 min video / 90 min audio\n✓ Everything in Essential\n✓ Team analytics\n✓ Gamification\n✓ Priority support"
    add_text(slide8, features_p, 3.2, 5.2, 3.2, 2.5, 11, MEDD_GRAY)
    
    # Enterprise
    shape = slide8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.8), Inches(1.2), Inches(3.8), Inches(5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = WHITE
    shape.line.fill.background()
    add_text(slide8, "ENTERPRISE", 1.4, 9, 3.5, 0.4, 12, MEDD_GREEN, True, 'center')
    add_text(slide8, "Custom", 1.9, 9, 3.5, 0.6, 40, MEDD_DARK, True, 'center')
    add_text(slide8, "contact us", 2.5, 9, 3.5, 0.3, 12, MEDD_GRAY, False, 'center')
    features_ent = "✓ Custom instance\n✓ High pooled allowances\n✓ SSO & governance\n✓ Volume rates\n✓ Phone support"
    add_text(slide8, features_ent, 3.2, 9.2, 3.2, 2.5, 11, MEDD_GRAY)
    
    # ===== SLIDE 9: CTA =====
    slide9 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide9, MEDD_LIME)
    
    add_text(slide9, "medd sim", 1.5, 0, 13.333, 0.8, 48, MEDD_GREEN, True, 'center')
    add_text(slide9, "Ready to transform how\nyour team practices?", 2.8, 0, 13.333, 1.2, 40, MEDD_DARK, True, 'center')
    add_text(slide9, "Turn high-stakes conversations into rehearsed performances.", 4.5, 0, 13.333, 0.6, 18, MEDD_GRAY, False, 'center')
    
    # CTA Button
    btn = slide9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.8), Inches(5.3), Inches(3.7), Inches(0.7))
    btn.fill.solid()
    btn.fill.fore_color.rgb = MEDD_GREEN
    btn.line.fill.background()
    add_text(slide9, "Start Free Trial", 5.4, 5.3, 2.8, 0.5, 18, WHITE, True, 'center')
    
    add_text(slide9, "📧 hello@medd.com.au   |   🌐 sim.medd.com.au", 6.3, 0, 13.333, 0.4, 14, MEDD_GRAY, False, 'center')
    
    # Save
    output_path = "/home/toti/projects/medd-sim-pitch/MEDD-SIM-Pitch-Deck.pptx"
    prs.save(output_path)
    print(f"✅ Presentation saved: {output_path}")
    return output_path

if __name__ == "__main__":
    create_presentation()
