#!/usr/bin/env python3
"""
Create a professional investor pitch deck for Vibe - AI-Native iOS Experience Platform
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

# Color scheme - Modern AI/Tech aesthetic
COLORS = {
    'primary': RGBColor(139, 92, 246),      # Purple #8B5CF6
    'secondary': RGBColor(6, 182, 212),      # Cyan #06B6D4
    'accent': RGBColor(244, 114, 182),       # Pink #F472B6
    'dark': RGBColor(15, 15, 26),            # Dark #0F0F1A
    'darker': RGBColor(8, 8, 15),            # Darker #08080F
    'white': RGBColor(255, 255, 255),
    'gray': RGBColor(156, 163, 175),         # Gray text
    'light_gray': RGBColor(200, 200, 210),
}

def set_slide_background(slide, color):
    """Set solid background color for a slide"""
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_gradient_shape(slide, left, top, width, height, color1, color2):
    """Add a gradient rectangle shape"""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.line.fill.background()
    fill = shape.fill
    fill.gradient()
    fill.gradient_stops[0].color.rgb = color1
    fill.gradient_stops[1].color.rgb = color2
    return shape

def add_text_box(slide, left, top, width, height, text, font_size=18,
                 font_color=COLORS['white'], bold=False, alignment=PP_ALIGN.LEFT,
                 font_name='Arial'):
    """Add a text box with specified styling"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = font_color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    return txBox

def add_card(slide, left, top, width, height, title, content, icon_text=""):
    """Add a card-style content block"""
    # Card background
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(30, 30, 50)
    card.line.color.rgb = RGBColor(60, 60, 80)
    card.line.width = Pt(1)

    # Icon circle if provided
    if icon_text:
        icon_size = Inches(0.4)
        icon = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                      left + Inches(0.2),
                                      top + Inches(0.2),
                                      icon_size, icon_size)
        icon.fill.gradient()
        icon.fill.gradient_stops[0].color.rgb = COLORS['primary']
        icon.fill.gradient_stops[1].color.rgb = COLORS['secondary']
        icon.line.fill.background()

    # Title
    title_top = top + Inches(0.2) if not icon_text else top + Inches(0.7)
    add_text_box(slide, left + Inches(0.2), title_top,
                 width - Inches(0.4), Inches(0.4),
                 title, font_size=14, bold=True, font_color=COLORS['primary'])

    # Content
    add_text_box(slide, left + Inches(0.2), title_top + Inches(0.35),
                 width - Inches(0.4), height - Inches(1),
                 content, font_size=11, font_color=COLORS['gray'])

def add_stat_card(slide, left, top, width, height, value, label):
    """Add a statistics card"""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(30, 30, 50)
    card.line.color.rgb = RGBColor(60, 60, 80)
    card.line.width = Pt(1)

    # Value with gradient effect (using purple)
    add_text_box(slide, left, top + Inches(0.3), width, Inches(0.6),
                 value, font_size=36, bold=True, font_color=COLORS['primary'],
                 alignment=PP_ALIGN.CENTER)

    # Label
    add_text_box(slide, left, top + Inches(0.9), width, Inches(0.4),
                 label, font_size=11, font_color=COLORS['gray'],
                 alignment=PP_ALIGN.CENTER)

def create_slide_1_cover(prs):
    """Slide 1: Cover"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    set_slide_background(slide, COLORS['darker'])

    # Decorative gradient circle (top right)
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7), Inches(-2), Inches(5), Inches(5))
    circle.fill.solid()
    circle.fill.fore_color.rgb = RGBColor(139, 92, 246)
    circle.fill.fore_color.brightness = 0.7
    circle.line.fill.background()

    # Logo placeholder (gradient square)
    logo = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   Inches(4.25), Inches(1.5), Inches(1.5), Inches(1.5))
    logo.fill.gradient()
    logo.fill.gradient_stops[0].color.rgb = COLORS['primary']
    logo.fill.gradient_stops[1].color.rgb = COLORS['secondary']
    logo.line.fill.background()

    # Company name
    add_text_box(slide, Inches(0), Inches(3.2), Inches(10), Inches(1),
                 "Vibe", font_size=72, bold=True, font_color=COLORS['primary'],
                 alignment=PP_ALIGN.CENTER)

    # Tagline
    add_text_box(slide, Inches(0), Inches(4.2), Inches(10), Inches(0.6),
                 "AI-Native iOS Experience Platform", font_size=24,
                 font_color=COLORS['gray'], alignment=PP_ALIGN.CENTER)

    # Funding ask badge
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                    Inches(3.5), Inches(5.2), Inches(3), Inches(0.6))
    badge.fill.gradient()
    badge.fill.gradient_stops[0].color.rgb = COLORS['primary']
    badge.fill.gradient_stops[1].color.rgb = COLORS['secondary']
    badge.line.fill.background()

    add_text_box(slide, Inches(3.5), Inches(5.25), Inches(3), Inches(0.5),
                 "Seed Round: $5M", font_size=18, bold=True,
                 font_color=COLORS['white'], alignment=PP_ALIGN.CENTER)

def create_slide_2_problem(prs):
    """Slide 2: Problem"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide, COLORS['darker'])

    # Section tag
    add_text_box(slide, Inches(0.5), Inches(0.5), Inches(2), Inches(0.4),
                 "THE PROBLEM", font_size=12, font_color=COLORS['gray'])

    # Title
    add_text_box(slide, Inches(0.5), Inches(0.9), Inches(9), Inches(0.8),
                 "Current AI apps fail iOS users", font_size=40, bold=True,
                 font_color=COLORS['white'])

    # Problem cards - 2x2 grid
    problems = [
        ("Non-Native Experience", "Most AI apps are web wrappers that feel foreign on iOS. No deep system integration."),
        ("Privacy Concerns", "All data sent to cloud servers. Users worried about sensitive information exposure."),
        ("Fragmented Workflows", "Multiple AI tools for different tasks. Context lost between applications."),
        ("Generic, Not Personal", "AI doesn't truly understand individual users. Every conversation starts from zero."),
    ]

    positions = [
        (Inches(0.5), Inches(2)),
        (Inches(5), Inches(2)),
        (Inches(0.5), Inches(4)),
        (Inches(5), Inches(4)),
    ]

    for (title, content), (left, top) in zip(problems, positions):
        add_card(slide, left, top, Inches(4.3), Inches(1.7), title, content, icon_text="!")

def create_slide_3_solution(prs):
    """Slide 3: Solution"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide, COLORS['darker'])

    # Section tag
    add_text_box(slide, Inches(0.5), Inches(0.5), Inches(2), Inches(0.4),
                 "OUR SOLUTION", font_size=12, font_color=COLORS['gray'])

    # Title
    add_text_box(slide, Inches(0.5), Inches(0.9), Inches(5), Inches(0.8),
                 "AI that lives on your iPhone", font_size=36, bold=True,
                 font_color=COLORS['white'])

    # Solution points
    solutions = [
        ("iOS-Native", "Built with Swift & SwiftUI, deep system integration"),
        ("Privacy-First", "On-device AI processing, data never leaves"),
        ("Truly Personal", "Learns your patterns, remembers context"),
        ("All-in-One", "Assistant, creator, and productivity suite"),
    ]

    y_pos = Inches(2)
    for title, desc in solutions:
        # Checkmark icon
        check = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.5), y_pos, Inches(0.3), Inches(0.3))
        check.fill.gradient()
        check.fill.gradient_stops[0].color.rgb = COLORS['primary']
        check.fill.gradient_stops[1].color.rgb = COLORS['secondary']
        check.line.fill.background()

        add_text_box(slide, Inches(1), y_pos - Inches(0.05), Inches(4.5), Inches(0.3),
                     title, font_size=16, bold=True, font_color=COLORS['white'])
        add_text_box(slide, Inches(1), y_pos + Inches(0.25), Inches(4.5), Inches(0.3),
                     desc, font_size=12, font_color=COLORS['gray'])
        y_pos += Inches(0.8)

    # iPhone mockup (right side)
    phone = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                    Inches(6), Inches(1.5), Inches(2.5), Inches(4.5))
    phone.fill.solid()
    phone.fill.fore_color.rgb = RGBColor(25, 25, 40)
    phone.line.color.rgb = COLORS['primary']
    phone.line.width = Pt(3)

    # App UI elements inside phone
    for i, color in enumerate([COLORS['primary'], COLORS['secondary'], COLORS['accent']]):
        ui_element = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                            Inches(6.3), Inches(2 + i * 1.2), Inches(1.9), Inches(0.9))
        ui_element.fill.solid()
        ui_element.fill.fore_color.rgb = color
        ui_element.fill.fore_color.brightness = 0.6
        ui_element.line.fill.background()

def create_slide_4_product(prs):
    """Slide 4: Product"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide, COLORS['darker'])

    # Section tag
    add_text_box(slide, Inches(0.5), Inches(0.5), Inches(2), Inches(0.4),
                 "PRODUCT", font_size=12, font_color=COLORS['gray'])

    # Title
    add_text_box(slide, Inches(0.5), Inches(0.9), Inches(9), Inches(0.8),
                 "Four pillars of intelligent living", font_size=36, bold=True,
                 font_color=COLORS['white'])

    # Product cards - 4 columns
    products = [
        ("VibeGPT", "Context-aware AI assistant with long-term memory and voice interaction"),
        ("Creator Suite", "AI image generation, smart editing, and content creation tools"),
        ("Health+", "Apple Health integration with AI insights and smart recommendations"),
        ("Productivity", "Smart scheduling, task management, and workflow automation"),
    ]

    card_width = Inches(2.1)
    start_x = Inches(0.5)
    gap = Inches(0.2)

    for i, (title, content) in enumerate(products):
        left = start_x + i * (card_width + gap)
        add_card(slide, left, Inches(2.2), card_width, Inches(3), title, content, icon_text="*")

def create_slide_5_market(prs):
    """Slide 5: Market Opportunity"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide, COLORS['darker'])

    # Section tag
    add_text_box(slide, Inches(0.5), Inches(0.5), Inches(3), Inches(0.4),
                 "MARKET OPPORTUNITY", font_size=12, font_color=COLORS['gray'])

    # Title
    add_text_box(slide, Inches(0.5), Inches(0.9), Inches(9), Inches(0.8),
                 "A massive, growing market", font_size=36, bold=True,
                 font_color=COLORS['white'])

    # TAM/SAM/SOM Funnel (left side)
    funnel_data = [
        ("$180B", "TAM - Global Mobile App Market", Inches(4.5), COLORS['primary']),
        ("$45B", "SAM - AI-Enhanced Productivity", Inches(3.5), COLORS['secondary']),
        ("$450M", "SOM - iOS AI-Native (3yr)", Inches(2.5), COLORS['accent']),
    ]

    y_pos = Inches(2)
    for value, label, width, color in funnel_data:
        # Funnel segment
        segment = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                         Inches(0.5), y_pos, width, Inches(1.1))
        segment.fill.solid()
        segment.fill.fore_color.rgb = color
        segment.fill.fore_color.brightness = 0.5
        segment.line.fill.background()

        add_text_box(slide, Inches(0.7), y_pos + Inches(0.15), width - Inches(0.4), Inches(0.5),
                     value, font_size=28, bold=True, font_color=COLORS['white'])
        add_text_box(slide, Inches(0.7), y_pos + Inches(0.6), width - Inches(0.4), Inches(0.4),
                     label, font_size=10, font_color=COLORS['light_gray'])
        y_pos += Inches(1.3)

    # Market Tailwinds (right side)
    add_text_box(slide, Inches(5.5), Inches(2), Inches(4), Inches(0.4),
                 "Market Tailwinds", font_size=16, bold=True, font_color=COLORS['primary'])

    tailwinds = [
        ("AI App Growth YoY", "+340%", 0.9),
        ("iOS User Premium (vs Android)", "4.5x", 0.85),
        ("On-device AI Interest", "72%", 0.72),
    ]

    y_pos = Inches(2.6)
    for label, value, progress in tailwinds:
        add_text_box(slide, Inches(5.5), y_pos, Inches(3), Inches(0.3),
                     label, font_size=11, font_color=COLORS['gray'])
        add_text_box(slide, Inches(8.5), y_pos, Inches(1), Inches(0.3),
                     value, font_size=11, bold=True, font_color=COLORS['primary'],
                     alignment=PP_ALIGN.RIGHT)

        # Progress bar background
        bar_bg = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                        Inches(5.5), y_pos + Inches(0.35), Inches(4), Inches(0.15))
        bar_bg.fill.solid()
        bar_bg.fill.fore_color.rgb = RGBColor(40, 40, 60)
        bar_bg.line.fill.background()

        # Progress bar fill
        bar_fill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                          Inches(5.5), y_pos + Inches(0.35),
                                          Inches(4 * progress), Inches(0.15))
        bar_fill.fill.gradient()
        bar_fill.fill.gradient_stops[0].color.rgb = COLORS['primary']
        bar_fill.fill.gradient_stops[1].color.rgb = COLORS['secondary']
        bar_fill.line.fill.background()

        y_pos += Inches(0.8)

def create_slide_6_business_model(prs):
    """Slide 6: Business Model"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide, COLORS['darker'])

    # Section tag
    add_text_box(slide, Inches(0.5), Inches(0.5), Inches(3), Inches(0.4),
                 "BUSINESS MODEL", font_size=12, font_color=COLORS['gray'])

    # Title
    add_text_box(slide, Inches(0.5), Inches(0.9), Inches(9), Inches(0.8),
                 "Proven subscription economics", font_size=36, bold=True,
                 font_color=COLORS['white'])

    # Pricing table (left)
    tiers = [
        ("Free", "$0", "User acquisition"),
        ("Pro", "$9.99/mo", "Power users"),
        ("Ultimate", "$19.99/mo", "Professionals"),
        ("Team", "$15/user/mo", "Small teams"),
    ]

    # Table header
    add_text_box(slide, Inches(0.5), Inches(2), Inches(1.5), Inches(0.4),
                 "Tier", font_size=11, bold=True, font_color=COLORS['gray'])
    add_text_box(slide, Inches(2), Inches(2), Inches(1.5), Inches(0.4),
                 "Price", font_size=11, bold=True, font_color=COLORS['gray'])
    add_text_box(slide, Inches(3.5), Inches(2), Inches(1.5), Inches(0.4),
                 "Target", font_size=11, bold=True, font_color=COLORS['gray'])

    y_pos = Inches(2.5)
    for tier, price, target in tiers:
        # Row separator
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                      Inches(0.5), y_pos - Inches(0.1), Inches(4.5), Pt(1))
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(50, 50, 70)
        line.line.fill.background()

        add_text_box(slide, Inches(0.5), y_pos, Inches(1.5), Inches(0.4),
                     tier, font_size=12, bold=True, font_color=COLORS['white'])
        add_text_box(slide, Inches(2), y_pos, Inches(1.5), Inches(0.4),
                     price, font_size=12, font_color=COLORS['light_gray'])
        add_text_box(slide, Inches(3.5), y_pos, Inches(1.5), Inches(0.4),
                     target, font_size=12, font_color=COLORS['gray'])
        y_pos += Inches(0.55)

    # Key metrics (right) - 2x2 grid
    metrics = [
        ("24:1", "LTV:CAC Ratio"),
        ("$85", "Customer LTV"),
        ("$3.50", "CAC"),
        ("3.5%", "Monthly Churn"),
    ]

    positions = [
        (Inches(5.5), Inches(2)),
        (Inches(7.5), Inches(2)),
        (Inches(5.5), Inches(3.7)),
        (Inches(7.5), Inches(3.7)),
    ]

    for (value, label), (left, top) in zip(metrics, positions):
        add_stat_card(slide, left, top, Inches(1.8), Inches(1.5), value, label)

def create_slide_7_competition(prs):
    """Slide 7: Competition"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide, COLORS['darker'])

    # Section tag
    add_text_box(slide, Inches(0.5), Inches(0.5), Inches(3), Inches(0.4),
                 "COMPETITIVE LANDSCAPE", font_size=12, font_color=COLORS['gray'])

    # Title
    add_text_box(slide, Inches(0.5), Inches(0.9), Inches(9), Inches(0.8),
                 "Differentiated positioning", font_size=36, bold=True,
                 font_color=COLORS['white'])

    # Position chart (left)
    chart_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                      Inches(0.5), Inches(2), Inches(4.5), Inches(3.5))
    chart_bg.fill.solid()
    chart_bg.fill.fore_color.rgb = RGBColor(25, 25, 40)
    chart_bg.line.color.rgb = RGBColor(50, 50, 70)

    # Axis labels
    add_text_box(slide, Inches(0.6), Inches(2.1), Inches(1.5), Inches(0.3),
                 "Feature Depth ↑", font_size=9, font_color=COLORS['gray'])
    add_text_box(slide, Inches(3.5), Inches(5.2), Inches(1.5), Inches(0.3),
                 "iOS Quality →", font_size=9, font_color=COLORS['gray'])

    # Competitor positions
    competitors = [
        ("ChatGPT", Inches(2.5), Inches(3)),
        ("Copilot", Inches(1.5), Inches(3.5)),
        ("Perplexity", Inches(1.2), Inches(4)),
        ("Notion AI", Inches(1.8), Inches(3.2)),
        ("Apple AI", Inches(3.5), Inches(4.2)),
    ]

    for name, x, y in competitors:
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, Inches(0.15), Inches(0.15))
        dot.fill.solid()
        dot.fill.fore_color.rgb = COLORS['gray']
        dot.line.fill.background()
        add_text_box(slide, x + Inches(0.2), y - Inches(0.05), Inches(1), Inches(0.3),
                     name, font_size=9, font_color=COLORS['gray'])

    # Vibe position (highlighted)
    vibe_dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(3.8), Inches(2.3), Inches(0.2), Inches(0.2))
    vibe_dot.fill.gradient()
    vibe_dot.fill.gradient_stops[0].color.rgb = COLORS['primary']
    vibe_dot.fill.gradient_stops[1].color.rgb = COLORS['secondary']
    vibe_dot.line.fill.background()
    add_text_box(slide, Inches(4.05), Inches(2.25), Inches(0.8), Inches(0.3),
                 "Vibe", font_size=11, bold=True, font_color=COLORS['primary'])

    # Our Moat (right)
    add_text_box(slide, Inches(5.5), Inches(2), Inches(4), Inches(0.4),
                 "Our Moat", font_size=16, bold=True, font_color=COLORS['primary'])

    moats = [
        ("1", "Proprietary Models", "Optimized for iOS & Apple Silicon"),
        ("2", "Data Flywheel", "Personalization improves with usage"),
        ("3", "Network Effects", "Community templates & sharing"),
        ("4", "Brand Position", "\"AI-Native\" mindshare"),
    ]

    y_pos = Inches(2.5)
    for num, title, desc in moats:
        # Number circle
        circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(5.5), y_pos, Inches(0.3), Inches(0.3))
        circle.fill.gradient()
        circle.fill.gradient_stops[0].color.rgb = COLORS['primary']
        circle.fill.gradient_stops[1].color.rgb = COLORS['secondary']
        circle.line.fill.background()

        add_text_box(slide, Inches(5.55), y_pos + Inches(0.02), Inches(0.2), Inches(0.25),
                     num, font_size=11, bold=True, font_color=COLORS['white'],
                     alignment=PP_ALIGN.CENTER)
        add_text_box(slide, Inches(5.95), y_pos, Inches(3.5), Inches(0.3),
                     title, font_size=12, bold=True, font_color=COLORS['white'])
        add_text_box(slide, Inches(5.95), y_pos + Inches(0.3), Inches(3.5), Inches(0.3),
                     desc, font_size=10, font_color=COLORS['gray'])
        y_pos += Inches(0.85)

def create_slide_8_traction(prs):
    """Slide 8: Traction"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide, COLORS['darker'])

    # Section tag
    add_text_box(slide, Inches(0.5), Inches(0.5), Inches(3), Inches(0.4),
                 "TRACTION & VALIDATION", font_size=12, font_color=COLORS['gray'])

    # Title
    add_text_box(slide, Inches(0.5), Inches(0.9), Inches(9), Inches(0.8),
                 "Strong early signals", font_size=36, bold=True,
                 font_color=COLORS['white'])

    # Key metrics - 4 columns
    metrics = [
        ("2.5K", "TestFlight Users"),
        ("52%", "Day 7 Retention"),
        ("47", "NPS Score"),
        ("34%", "Payment Intent"),
    ]

    start_x = Inches(0.5)
    card_width = Inches(2.1)
    gap = Inches(0.2)

    for i, (value, label) in enumerate(metrics):
        left = start_x + i * (card_width + gap)
        add_stat_card(slide, left, Inches(2), card_width, Inches(1.5), value, label)

    # User testimonials
    add_text_box(slide, Inches(0.5), Inches(3.8), Inches(3), Inches(0.4),
                 "User Feedback", font_size=14, bold=True, font_color=COLORS['primary'])

    testimonials = [
        ("\"Finally an AI app that feels like it belongs on my iPhone.\"", "Product Designer"),
        ("\"The privacy-first approach is exactly what I've been looking for.\"", "Lawyer"),
        ("\"It actually remembers our previous conversations. Game changer.\"", "Entrepreneur"),
    ]

    start_x = Inches(0.5)
    card_width = Inches(3)

    for i, (quote, role) in enumerate(testimonials):
        left = start_x + i * (card_width + Inches(0.15))
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      left, Inches(4.3), card_width, Inches(1.4))
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(30, 30, 50)
        card.line.color.rgb = RGBColor(60, 60, 80)

        add_text_box(slide, left + Inches(0.15), Inches(4.4), card_width - Inches(0.3), Inches(0.9),
                     quote, font_size=10, font_color=COLORS['gray'])
        add_text_box(slide, left + Inches(0.15), Inches(5.3), card_width - Inches(0.3), Inches(0.3),
                     f"— Beta User, {role}", font_size=9, font_color=COLORS['primary'])

def create_slide_9_roadmap(prs):
    """Slide 9: Roadmap"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide, COLORS['darker'])

    # Section tag
    add_text_box(slide, Inches(0.5), Inches(0.5), Inches(2), Inches(0.4),
                 "ROADMAP", font_size=12, font_color=COLORS['gray'])

    # Title
    add_text_box(slide, Inches(0.5), Inches(0.9), Inches(9), Inches(0.8),
                 "Path to scale", font_size=36, bold=True,
                 font_color=COLORS['white'])

    # Timeline
    milestones = [
        ("Q1-Q2 2024", "MVP Launch", "Core AI, subscriptions, App Store"),
        ("Q3-Q4 2024", "Expansion", "Full features, 100K users"),
        ("H1 2025", "Scale", "International, enterprise, 1M users"),
        ("H2 2025", "Platform", "Open API, plugins, Series A"),
    ]

    # Timeline line
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                  Inches(0.5), Inches(2.3), Inches(9), Pt(4))
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(50, 50, 70)
    line.line.fill.background()

    start_x = Inches(0.5)
    segment_width = Inches(2.25)

    for i, (date, title, desc) in enumerate(milestones):
        x = start_x + i * segment_width

        # Active segment (first one)
        if i == 0:
            active_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                                 x, Inches(2.3), segment_width, Pt(4))
            active_line.fill.gradient()
            active_line.fill.gradient_stops[0].color.rgb = COLORS['primary']
            active_line.fill.gradient_stops[1].color.rgb = COLORS['secondary']
            active_line.line.fill.background()

        # Dot
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, Inches(2.2), Inches(0.2), Inches(0.2))
        dot.fill.gradient()
        dot.fill.gradient_stops[0].color.rgb = COLORS['primary']
        dot.fill.gradient_stops[1].color.rgb = COLORS['secondary']
        dot.line.fill.background()

        add_text_box(slide, x, Inches(2.5), segment_width, Inches(0.3),
                     date, font_size=11, bold=True, font_color=COLORS['primary'])
        add_text_box(slide, x, Inches(2.85), segment_width, Inches(0.3),
                     title, font_size=12, bold=True, font_color=COLORS['white'])
        add_text_box(slide, x, Inches(3.15), segment_width - Inches(0.2), Inches(0.4),
                     desc, font_size=9, font_color=COLORS['gray'])

    # Financial projections table
    add_text_box(slide, Inches(0.5), Inches(4), Inches(3), Inches(0.4),
                 "Financial Projections", font_size=14, bold=True, font_color=COLORS['primary'])

    # Table headers
    headers = ["Year", "Users", "Conversion", "ARR"]
    x_positions = [Inches(0.5), Inches(2.5), Inches(4.5), Inches(6.5)]

    for header, x in zip(headers, x_positions):
        add_text_box(slide, x, Inches(4.5), Inches(2), Inches(0.3),
                     header, font_size=10, bold=True, font_color=COLORS['gray'])

    # Table data
    data = [
        ("Year 1", "100K", "8%", "$960K"),
        ("Year 2", "500K", "12%", "$7.2M"),
        ("Year 3", "2M", "15%", "$36M"),
    ]

    y_pos = Inches(4.9)
    for row in data:
        # Row line
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                      Inches(0.5), y_pos - Inches(0.05), Inches(8), Pt(1))
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(50, 50, 70)
        line.line.fill.background()

        for value, x in zip(row, x_positions):
            color = COLORS['primary'] if value.startswith('$') else COLORS['white']
            add_text_box(slide, x, y_pos, Inches(2), Inches(0.35),
                         value, font_size=11, font_color=color,
                         bold=value.startswith('$'))
        y_pos += Inches(0.45)

def create_slide_10_team(prs):
    """Slide 10: Team"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide, COLORS['darker'])

    # Section tag
    add_text_box(slide, Inches(0.5), Inches(0.5), Inches(2), Inches(0.4),
                 "TEAM", font_size=12, font_color=COLORS['gray'])

    # Title
    add_text_box(slide, Inches(0.5), Inches(0.9), Inches(9), Inches(0.8),
                 "Built by industry veterans", font_size=36, bold=True,
                 font_color=COLORS['white'])

    # Team members
    team = [
        ("JC", "James Chen", "CEO / Co-founder", "Ex-Apple Senior Engineer\n10 years iOS experience\nPrevious startup acquired"),
        ("SW", "Sarah Wang", "CTO / Co-founder", "Ex-Google Brain Researcher\nML PhD, Stanford\nMobile AI specialist"),
        ("MK", "Michael Kim", "CPO / Co-founder", "Ex-Instagram Product Lead\nMultiple 10M+ user products"),
    ]

    start_x = Inches(0.5)
    card_width = Inches(3)
    gap = Inches(0.15)

    for i, (initials, name, role, bio) in enumerate(team):
        left = start_x + i * (card_width + gap)

        # Card background
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      left, Inches(2), card_width, Inches(2.8))
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(30, 30, 50)
        card.line.color.rgb = RGBColor(60, 60, 80)

        # Avatar circle
        avatar = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                        left + Inches(1), Inches(2.2), Inches(1), Inches(1))
        avatar.fill.gradient()
        avatar.fill.gradient_stops[0].color.rgb = COLORS['primary']
        avatar.fill.gradient_stops[1].color.rgb = COLORS['secondary']
        avatar.line.fill.background()

        add_text_box(slide, left + Inches(1), Inches(2.5), Inches(1), Inches(0.5),
                     initials, font_size=24, bold=True, font_color=COLORS['white'],
                     alignment=PP_ALIGN.CENTER)

        add_text_box(slide, left, Inches(3.3), card_width, Inches(0.3),
                     name, font_size=14, bold=True, font_color=COLORS['white'],
                     alignment=PP_ALIGN.CENTER)
        add_text_box(slide, left, Inches(3.6), card_width, Inches(0.3),
                     role, font_size=10, font_color=COLORS['primary'],
                     alignment=PP_ALIGN.CENTER)
        add_text_box(slide, left + Inches(0.2), Inches(3.95), card_width - Inches(0.4), Inches(0.8),
                     bio, font_size=9, font_color=COLORS['gray'],
                     alignment=PP_ALIGN.CENTER)

    # Advisors
    add_text_box(slide, Inches(0.5), Inches(5), Inches(2), Inches(0.3),
                 "Advisors", font_size=12, bold=True, font_color=COLORS['primary'])

    advisors = [
        ("Alex Turner", "Partner, Top-tier VC"),
        ("Emily Zhang", "Former Apple VP"),
        ("Dr. Robert Lee", "AI Professor, Stanford"),
        ("Lisa Park", "Growth Lead, Unicorn"),
    ]

    x_pos = Inches(0.5)
    for name, title in advisors:
        add_text_box(slide, x_pos, Inches(5.35), Inches(2.2), Inches(0.25),
                     name, font_size=10, bold=True, font_color=COLORS['white'])
        add_text_box(slide, x_pos, Inches(5.55), Inches(2.2), Inches(0.25),
                     title, font_size=8, font_color=COLORS['gray'])
        x_pos += Inches(2.4)

def create_slide_11_ask(prs):
    """Slide 11: The Ask"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide, COLORS['darker'])

    # Decorative orbs
    circle1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-2), Inches(-2), Inches(5), Inches(5))
    circle1.fill.solid()
    circle1.fill.fore_color.rgb = COLORS['primary']
    circle1.fill.fore_color.brightness = 0.7
    circle1.line.fill.background()

    circle2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(8), Inches(4), Inches(4), Inches(4))
    circle2.fill.solid()
    circle2.fill.fore_color.rgb = COLORS['secondary']
    circle2.fill.fore_color.brightness = 0.7
    circle2.line.fill.background()

    # Section tag
    add_text_box(slide, Inches(0), Inches(0.8), Inches(10), Inches(0.4),
                 "THE ASK", font_size=12, font_color=COLORS['gray'],
                 alignment=PP_ALIGN.CENTER)

    # Title
    add_text_box(slide, Inches(0), Inches(1.2), Inches(10), Inches(0.6),
                 "Join us in building the future of AI on iOS", font_size=24, bold=True,
                 font_color=COLORS['white'], alignment=PP_ALIGN.CENTER)

    # Big number
    add_text_box(slide, Inches(0), Inches(2), Inches(10), Inches(1.2),
                 "$5M", font_size=96, bold=True, font_color=COLORS['primary'],
                 alignment=PP_ALIGN.CENTER)

    add_text_box(slide, Inches(0), Inches(3.2), Inches(10), Inches(0.5),
                 "Seed Round", font_size=20, font_color=COLORS['gray'],
                 alignment=PP_ALIGN.CENTER)

    # Use of funds - 4 cards
    funds = [
        ("50%", "Product Dev"),
        ("25%", "Marketing"),
        ("15%", "Team Growth"),
        ("10%", "Operations"),
    ]

    start_x = Inches(1)
    card_width = Inches(1.8)
    gap = Inches(0.3)

    for i, (pct, label) in enumerate(funds):
        left = start_x + i * (card_width + gap)
        add_stat_card(slide, left, Inches(3.9), card_width, Inches(1.3), pct, label)

    # Contact info
    add_text_box(slide, Inches(2), Inches(5.5), Inches(3), Inches(0.4),
                 "✉  investors@vibe.app", font_size=12, font_color=COLORS['gray'])
    add_text_box(slide, Inches(5), Inches(5.5), Inches(3), Inches(0.4),
                 "🌐  www.vibe.app", font_size=12, font_color=COLORS['gray'])

def create_slide_12_thanks(prs):
    """Slide 12: Thank You"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_background(slide, COLORS['darker'])

    # Decorative elements
    circle1 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6), Inches(-3), Inches(6), Inches(6))
    circle1.fill.solid()
    circle1.fill.fore_color.rgb = COLORS['primary']
    circle1.fill.fore_color.brightness = 0.7
    circle1.line.fill.background()

    circle2 = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-2), Inches(4), Inches(5), Inches(5))
    circle2.fill.solid()
    circle2.fill.fore_color.rgb = COLORS['secondary']
    circle2.fill.fore_color.brightness = 0.7
    circle2.line.fill.background()

    # Logo
    logo = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   Inches(4.25), Inches(1.8), Inches(1.5), Inches(1.5))
    logo.fill.gradient()
    logo.fill.gradient_stops[0].color.rgb = COLORS['primary']
    logo.fill.gradient_stops[1].color.rgb = COLORS['secondary']
    logo.line.fill.background()

    # Thank you text
    add_text_box(slide, Inches(0), Inches(3.5), Inches(10), Inches(0.9),
                 "Thank You", font_size=56, bold=True, font_color=COLORS['white'],
                 alignment=PP_ALIGN.CENTER)

    add_text_box(slide, Inches(0), Inches(4.4), Inches(10), Inches(0.5),
                 "Let's shape the future of AI together", font_size=18,
                 font_color=COLORS['gray'], alignment=PP_ALIGN.CENTER)

    # Contact
    add_text_box(slide, Inches(0), Inches(5.3), Inches(10), Inches(0.4),
                 "investors@vibe.app  |  www.vibe.app", font_size=12,
                 font_color=COLORS['gray'], alignment=PP_ALIGN.CENTER)

def main():
    # Create presentation
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)  # 16:9 aspect ratio

    # Create all slides
    create_slide_1_cover(prs)
    create_slide_2_problem(prs)
    create_slide_3_solution(prs)
    create_slide_4_product(prs)
    create_slide_5_market(prs)
    create_slide_6_business_model(prs)
    create_slide_7_competition(prs)
    create_slide_8_traction(prs)
    create_slide_9_roadmap(prs)
    create_slide_10_team(prs)
    create_slide_11_ask(prs)
    create_slide_12_thanks(prs)

    # Save
    output_path = '/home/user/iOS-Exploration/Vibe_Investor_Pitch_Deck.pptx'
    prs.save(output_path)
    print(f"Pitch deck saved to: {output_path}")

if __name__ == "__main__":
    main()
