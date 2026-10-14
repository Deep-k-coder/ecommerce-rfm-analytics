"""
build_phase2_presentation.py
Generates the comprehensive 10-slide PowerPoint presentation (.pptx) for the
Second Phase Work (Mid-Term Exam Evaluation - 20 Marks):
- Degree: B.Sc. in Artificial Intelligence & Data Science, Semester V (5th Semester)
- Institution: Institute of Advanced Research (IAR), Gandhinagar
- NO DEPARTMENT NAME mentioned anywhere in the presentation.
- Official IAR banner at top of Title & Thank You slides.
- 16:9 Widescreen format (13.333" x 7.5").
- Exactly 10 slides, culminating in an elegant "Thank You & Q&A" slide.
- High-res UML Class Diagram, Sequence Diagram, and Screen Layouts visuals.
- Color Palette: IAR Maroon (#7A003C), Deep Navy (#1E1B4B), Slate (#475569), Clean Whites & Light Grays.
"""

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

# Brand Colors
COLOR_MAROON = RGBColor(122, 0, 60)      # IAR Brand Maroon #7A003C
COLOR_NAVY = RGBColor(30, 27, 75)        # Deep Navy #1E1B4B
COLOR_DARK = RGBColor(15, 23, 42)        # Body Dark #0F172A
COLOR_SLATE = RGBColor(71, 85, 105)      # Slate Gray #475569
COLOR_MUTED = RGBColor(148, 163, 184)    # Light Slate #94A3B8
COLOR_CARD_BG = RGBColor(248, 250, 252)  # Card Background #F8FAFC
COLOR_BORDER = RGBColor(226, 232, 240)   # Border #E2E8F0
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_EMERALD = RGBColor(16, 185, 129)   # Accent Green #10B981
COLOR_BLUE = RGBColor(37, 99, 235)       # Accent Blue #2563EB
COLOR_AMBER = RGBColor(217, 119, 6)      # Accent Amber #D97706

def create_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    base_dir = Path("/Users/deep/.gemini/antigravity/scratch/ecommerce_rfm_analytics")
    img_dir = base_dir / "report_images"

    def add_header(slide, category_text, title_text, slide_num):
        # Top Maroon accent line
        accent_line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.08)
        )
        accent_line.fill.solid()
        accent_line.fill.fore_color.rgb = COLOR_MAROON
        accent_line.line.color.rgb = COLOR_MAROON

        # Category Chip
        tb_cat = slide.shapes.add_textbox(Inches(0.8), Inches(0.25), Inches(8.5), Inches(0.35))
        p_cat = tb_cat.text_frame.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.name = 'Arial'
        p_cat.font.size = Pt(10)
        p_cat.font.bold = True
        p_cat.font.color.rgb = COLOR_MAROON

        # Slide Title
        tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.55), Inches(9.5), Inches(0.65))
        p_title = tb_title.text_frame.paragraphs[0]
        p_title.text = title_text
        p_title.font.name = 'Arial'
        p_title.font.size = Pt(22)
        p_title.font.bold = True
        p_title.font.color.rgb = COLOR_NAVY

        # Slide Number Badge
        tb_num = slide.shapes.add_textbox(Inches(11.2), Inches(0.35), Inches(1.3), Inches(0.4))
        p_num = tb_num.text_frame.paragraphs[0]
        p_num.alignment = PP_ALIGN.RIGHT
        p_num.text = f"{slide_num:02d} / 10"
        p_num.font.name = 'Arial'
        p_num.font.size = Pt(11)
        p_num.font.bold = True
        p_num.font.color.rgb = COLOR_SLATE

        # Subtle divider
        div = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.22), Inches(11.733), Inches(0.015)
        )
        div.fill.solid()
        div.fill.fore_color.rgb = COLOR_BORDER
        div.line.color.rgb = COLOR_BORDER

    def add_footer(slide):
        tb_foot = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.35))
        p_foot = tb_foot.text_frame.paragraphs[0]
        p_foot.text = "B.Sc. in Artificial Intelligence & Data Science (Sem V)  •  Institute of Advanced Research  •  Second Progress Work (13/10/2026)"
        p_foot.font.name = 'Arial'
        p_foot.font.size = Pt(9)
        p_foot.font.color.rgb = COLOR_MUTED

    def add_card(slide, left, top, width, height, bg_color=COLOR_CARD_BG, border_color=COLOR_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1)
        return card

    # =========================================================================
    # SLIDE 1: COVER SLIDE
    # =========================================================================
    slide1 = prs.slides.add_slide(blank_layout)

    top_bar = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = COLOR_MAROON
    top_bar.line.color.rgb = COLOR_MAROON

    bot_bar = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.38), Inches(13.333), Inches(0.12))
    bot_bar.fill.solid()
    bot_bar.fill.fore_color.rgb = COLOR_NAVY
    bot_bar.line.color.rgb = COLOR_NAVY

    iar_banner = img_dir / "iar_banner_hd.png"
    if iar_banner.exists():
        slide1.shapes.add_picture(str(iar_banner), Inches(3.9), Inches(0.45), width=Inches(5.5))

    tb_univ = slide1.shapes.add_textbox(Inches(1.0), Inches(1.65), Inches(11.333), Inches(0.6))
    p_u = tb_univ.text_frame.paragraphs[0]
    p_u.alignment = PP_ALIGN.CENTER
    p_u.text = "INSTITUTE OF ADVANCED RESEARCH (IAR)"
    p_u.font.name = 'Arial'
    p_u.font.size = Pt(17)
    p_u.font.bold = True
    p_u.font.color.rgb = COLOR_MAROON

    p_u2 = tb_univ.text_frame.add_paragraph()
    p_u2.alignment = PP_ALIGN.CENTER
    p_u2.text = "The University for Innovation  |  Gandhinagar, Gujarat"
    p_u2.font.name = 'Arial'
    p_u2.font.size = Pt(11)
    p_u2.font.italic = True
    p_u2.font.color.rgb = COLOR_SLATE

    badge = add_card(slide1, Inches(3.0), Inches(2.35), Inches(7.333), Inches(0.42), bg_color=RGBColor(236, 253, 245), border_color=COLOR_EMERALD)
    tb_badge = slide1.shapes.add_textbox(Inches(3.0), Inches(2.35), Inches(7.333), Inches(0.42))
    p_b = tb_badge.text_frame.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    p_b.text = "SECOND PROGRESS WORK (MID-TERM EXAM)  •  13TH OCTOBER 2026  •  20 MARKS"
    p_b.font.name = 'Arial'
    p_b.font.size = Pt(10)
    p_b.font.bold = True
    p_b.font.color.rgb = COLOR_EMERALD

    tb_t = slide1.shapes.add_textbox(Inches(1.0), Inches(2.9), Inches(11.333), Inches(1.3))
    p_t = tb_t.text_frame.paragraphs[0]
    p_t.alignment = PP_ALIGN.CENTER
    p_t.text = "E-COMMERCE CUSTOMER LIFECYCLE &\nRFM SEGMENTATION DASHBOARD"
    p_t.font.name = 'Arial'
    p_t.font.size = Pt(28)
    p_t.font.bold = True
    p_t.font.color.rgb = COLOR_NAVY

    p_sub = tb_t.text_frame.add_paragraph()
    p_sub.alignment = PP_ALIGN.CENTER
    p_sub.text = "Second Phase Technical Report: UML Class & Sequence Models, Data Dictionary & UI Layout Validations"
    p_sub.font.name = 'Arial'
    p_sub.font.size = Pt(13)
    p_sub.font.italic = True
    p_sub.font.color.rgb = COLOR_MAROON

    c1 = add_card(slide1, Inches(1.5), Inches(4.5), Inches(5.0), Inches(2.3))
    tb_c1 = slide1.shapes.add_textbox(Inches(1.7), Inches(4.65), Inches(4.6), Inches(2.0))
    tf1 = tb_c1.text_frame
    tf1.word_wrap = True
    p = tf1.paragraphs[0]
    p.text = "PROJECT SCHOLAR"
    p.font.name = 'Arial'
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    lines_c1 = [
        ("Candidate Name: ", "Mr. Deep Koshiya"),
        ("Program: ", "B.Sc. in AI & Data Science"),
        ("Semester: ", "Semester V (5th Semester)"),
        ("Academic Year: ", "2026 – 2027"),
        ("Course: ", "Project - 1 (CE-502 / PRJ-701)")
    ]
    for lbl, val in lines_c1:
        p = tf1.add_paragraph()
        r1 = p.add_run()
        r1.text = lbl
        r1.font.bold = True
        r1.font.size = Pt(10)
        r1.font.color.rgb = COLOR_DARK
        r2 = p.add_run()
        r2.text = val
        r2.font.size = Pt(10)
        r2.font.color.rgb = COLOR_DARK

    c2 = add_card(slide1, Inches(6.8), Inches(4.5), Inches(5.0), Inches(2.3))
    tb_c2 = slide1.shapes.add_textbox(Inches(7.0), Inches(4.65), Inches(4.6), Inches(2.0))
    tf2 = tb_c2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "EVALUATION & DELIVERABLE SCOPE"
    p.font.name = 'Arial'
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    lines_c2 = [
        ("Milestone: ", "Second Progress Work (Mid-Term)"),
        ("Evaluation Date: ", "13th October 2026 (10 AM - 5 PM)"),
        ("Marks Allotted: ", "20 Marks (Mid-Term Assessment)"),
        ("Syllabus Scope: ", "Sections 2.3, 2.4, 3.1, 4.1"),
        ("Supervision: ", "Assigned Project Supervisor (IAR)")
    ]
    for lbl, val in lines_c2:
        p = tf2.add_paragraph()
        r1 = p.add_run()
        r1.text = lbl
        r1.font.bold = True
        r1.font.size = Pt(10)
        r1.font.color.rgb = COLOR_DARK
        r2 = p.add_run()
        r2.text = val
        r2.font.size = Pt(10)
        r2.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 2: EXECUTIVE SUMMARY & PHASE 1 FOUNDATION RECAP
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    add_header(slide2, "Progressive Architecture  •  Overview", "System Architecture & Phase 1 Foundations Recap", 2)
    add_footer(slide2)

    col_w = Inches(3.64)
    gap = Inches(0.4)
    y_top = Inches(1.5)
    card_h = Inches(5.2)

    # Card 1: Completed First Progress Deliverables
    add_card(slide2, Inches(0.8), y_top, col_w, card_h)
    tb = slide2.shapes.add_textbox(Inches(0.95), y_top + Inches(0.15), col_w - Inches(0.3), card_h - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "1. Phase 1 Foundations (10 Marks)"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    b1 = [
        ("Problem Formulation (1.1): ", "Defined retention economics, surging CAC, and OLTP operational gaps."),
        ("Requirement Matrix (1.2): ", "Specified FR-01 through FR-06 and NFR-01 to NFR-03 with hardware specs."),
        ("Tech Stack Layering (1.3): ", "Configured Python 3.9, Pandas 2.3, Scikit-Learn 1.6, and Chart.js."),
        ("DFD Architecture (2.1): ", "Constructed Context Level 0, Level 1 Pipeline, and Level 2 Decomposition DFDs.")
    ]
    for lbl, val in b1:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r0 = p.add_run()
        r0.text = "• " + lbl
        r0.font.bold = True
        r0.font.size = Pt(10)
        r0.font.color.rgb = COLOR_NAVY
        r1 = p.add_run()
        r1.text = val
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = COLOR_DARK

    # Card 2: Mid-Term Milestones (Phase 2 Focus)
    add_card(slide2, Inches(0.8) + col_w + gap, y_top, col_w, card_h, bg_color=RGBColor(240, 253, 244), border_color=COLOR_EMERALD)
    tb = slide2.shapes.add_textbox(Inches(0.95) + col_w + gap, y_top + Inches(0.15), col_w - Inches(0.3), card_h - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "2. Phase 2 Deliverables (20 Marks)"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_EMERALD

    b2 = [
        ("UML Class Diagram (2.3): ", "Domain entity model & pipeline class architecture (TransactionRecord, Refinery, RFM, KMeans)."),
        ("UML Sequence Diagram (2.4): ", "11-step synchronous/asynchronous message dispatch protocol across system lifelines."),
        ("Data Dictionary (3.1): ", "Complete star-schema catalog: FACT_TRANSACTIONS, FACT_REVERSALS, DIM_CUSTOMERS."),
        ("Screen Layouts (4.1): ", "Multi-screen UI layouts with operational input validations (VAL-01 to VAL-04).")
    ]
    for lbl, val in b2:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r0 = p.add_run()
        r0.text = "• " + lbl
        r0.font.bold = True
        r0.font.size = Pt(10)
        r0.font.color.rgb = COLOR_NAVY
        r1 = p.add_run()
        r1.text = val
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = COLOR_DARK

    # Card 3: Business & Academic Impact
    add_card(slide2, Inches(0.8) + (col_w + gap) * 2, y_top, col_w, card_h)
    tb = slide2.shapes.add_textbox(Inches(0.95) + (col_w + gap) * 2, y_top + Inches(0.15), col_w - Inches(0.3), card_h - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "3. Technical Value Proposition"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    b3 = [
        ("Zero Data Leakage: ", "Strict segregation of reversals ensures return items never pollute positive purchase frequency counts."),
        ("Skew-Resistant Clustering: ", "Log1p scaling neutralizes Pareto monetary power laws for robust K-Means clustering."),
        ("BI Mart Compatibility: ", "Star schema facilitates native query performance in Tableau, Power BI, and SQL engines."),
        ("Instant CRM Activation: ", "Interactive UI enables commercial teams to filter and export high-value at-risk cohorts instantly.")
    ]
    for lbl, val in b3:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r0 = p.add_run()
        r0.text = "• " + lbl
        r0.font.bold = True
        r0.font.size = Pt(10)
        r0.font.color.rgb = COLOR_NAVY
        r1 = p.add_run()
        r1.text = val
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 3: 2.3 UML CLASS DIAGRAM
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    add_header(slide3, "Section 2.3  •  System Design", "UML Class Diagram (Domain Model & Object Architecture)", 3)
    add_footer(slide3)

    # Left: Image Visual Card (Width 6.8")
    add_card(slide3, Inches(0.8), Inches(1.5), Inches(6.8), Inches(5.2))
    cls_img = img_dir / "fig_class_diagram.png"
    if cls_img.exists():
        slide3.shapes.add_picture(str(cls_img), Inches(0.95), Inches(1.7), width=Inches(6.5))
    tb_cap = slide3.shapes.add_textbox(Inches(0.8), Inches(6.25), Inches(6.8), Inches(0.35))
    p = tb_cap.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "Figure 2.5: UML Class Diagram (Domain Model & Pipeline Services)"
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_SLATE

    # Right: Technical Walkthrough (Width 4.6")
    add_card(slide3, Inches(7.9), Inches(1.5), Inches(4.633), Inches(5.2))
    tb_desc = slide3.shapes.add_textbox(Inches(8.1), Inches(1.65), Inches(4.2), Inches(4.9))
    tf = tb_desc.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Core Pipeline Classes & Relationships"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    cls_points = [
        ("TransactionRecord: ", "Domain entity representing single line-item purchases with validation and pricing methods."),
        ("CleanRefineryPipeline: ", "Orchestrates ingestion, typing, guest session filtering, reversal quarantine, and Tukey IQR capping."),
        ("RFMFeatureEngineer: ", "Vectorized engine computing customer Recency (T_obs), Frequency, Spend, and AOV."),
        ("KMeansSegmentationEngine: ", "Encapsulates log1p compression, StandardScaler, K-Means fitting (k=4), and Silhouette diagnostics."),
        ("QuantileStratifier: ", "Assigns 1–5 quantile scores (pd.qcut) and maps accounts to 8 business operational cohorts."),
        ("StarSchemaMartExporter: ", "Generates FACT_TRANSACTIONS, FACT_REVERSALS, DIM_CUSTOMERS, and summary JSON.")
    ]
    for c_lbl, c_desc in cls_points:
        p = tf.add_paragraph()
        p.space_before = Pt(4)
        r0 = p.add_run()
        r0.text = "• " + c_lbl
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        r0.font.color.rgb = COLOR_NAVY
        r1 = p.add_run()
        r1.text = c_desc
        r1.font.size = Pt(8.8)
        r1.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 4: 2.4 UML SEQUENCE DIAGRAM
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    add_header(slide4, "Section 2.4  •  Execution Architecture", "UML Sequence Diagram (Pipeline Execution Workflow)", 4)
    add_footer(slide4)

    # Left: Image Visual Card (Width 6.8")
    add_card(slide4, Inches(0.8), Inches(1.5), Inches(6.8), Inches(5.2))
    seq_img = img_dir / "fig_sequence_diagram.png"
    if seq_img.exists():
        slide4.shapes.add_picture(str(seq_img), Inches(0.95), Inches(1.7), width=Inches(6.5))
    tb_cap = slide4.shapes.add_textbox(Inches(0.8), Inches(6.25), Inches(6.8), Inches(0.35))
    p = tb_cap.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "Figure 2.6: UML Sequence Diagram (End-to-End Processing & Activation)"
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_SLATE

    # Right: Message Dispatch Protocol (Width 4.6")
    add_card(slide4, Inches(7.9), Inches(1.5), Inches(4.633), Inches(5.2))
    tb_desc = slide4.shapes.add_textbox(Inches(8.1), Inches(1.65), Inches(4.2), Inches(4.9))
    tf = tb_desc.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "11-Step Message Dispatch Protocol"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    seq_points = [
        ("Step 1 (Trigger): ", "Data Engineer initiates trigger_pipeline(raw_csv) to commence ingestion."),
        ("Steps 2-3 (Refinery): ", "Self-calls filter unauthenticated sessions, quarantine cancellations, and cap Tukey IQR outliers."),
        ("Steps 4-5 (Feature Extraction): ", "Cleansed records dispatched to RFMFeatureEngineer, deriving snapshot boundary T_obs."),
        ("Steps 6-7 (ML Modeling): ", "RFM Analytical Base Table dispatched to KMeansEngine for log1p scaling and k=4 clustering."),
        ("Steps 8-9 (Data Mart Export): ", "Segments and fact tables persisted to StarSchemaDataMart."),
        ("Steps 10-11 (Activation): ", "Summary JSON exported to ExecutiveDashboard, rendering KPI cards and spatial scatterplots.")
    ]
    for s_lbl, s_desc in seq_points:
        p = tf.add_paragraph()
        p.space_before = Pt(4)
        r0 = p.add_run()
        r0.text = "• " + s_lbl
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        r0.font.color.rgb = COLOR_NAVY
        r1 = p.add_run()
        r1.text = s_desc
        r1.font.size = Pt(8.8)
        r1.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 5: 3.1 DATA DICTIONARY - FACT TABLES
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    add_header(slide5, "Section 3.1  •  Data Dictionary", "Data Dictionary: Fact Tables (FACT_TRANSACTIONS & FACT_REVERSALS)", 5)
    add_footer(slide5)

    # Left: FACT_TRANSACTIONS (Width 5.7")
    add_card(slide5, Inches(0.8), Inches(1.5), Inches(5.7), Inches(5.2))
    tb1 = slide5.shapes.add_textbox(Inches(1.0), Inches(1.65), Inches(5.3), Inches(4.9))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p = tf1.paragraphs[0]
    p.text = "Table 3.1: FACT_TRANSACTIONS (Order Fact Mart)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    f_items = [
        ("transaction_id (BIGINT, PK): ", "Surrogate primary key generated monotonically."),
        ("invoice_no (VARCHAR(20)): ", "6-digit order identifier (e.g. 536365)."),
        ("stock_code (VARCHAR(20), FK): ", "Foreign key referencing DIM_PRODUCTS."),
        ("quantity (INTEGER): ", "Item units purchased; strictly Quantity > 0."),
        ("invoice_date (TIMESTAMP, FK): ", "Transaction timestamp referencing DIM_DATE."),
        ("unit_price (DECIMAL(10,2)): ", "Product sale unit price; strictly UnitPrice > 0."),
        ("customer_id (VARCHAR(20), FK): ", "Foreign key referencing DIM_CUSTOMERS."),
        ("gross_amount (DECIMAL(12,2)): ", "Line spend: Quantity * UnitPrice.")
    ]
    for lbl, desc in f_items:
        p = tf1.add_paragraph()
        p.space_before = Pt(3)
        r0 = p.add_run()
        r0.text = "• " + lbl
        r0.font.bold = True
        r0.font.size = Pt(9)
        r0.font.color.rgb = COLOR_NAVY
        r1 = p.add_run()
        r1.text = desc
        r1.font.size = Pt(8.5)
        r1.font.color.rgb = COLOR_DARK

    # Right: FACT_REVERSALS (Width 5.7")
    add_card(slide5, Inches(6.8), Inches(1.5), Inches(5.733), Inches(5.2))
    tb2 = slide5.shapes.add_textbox(Inches(7.0), Inches(1.65), Inches(5.3), Inches(4.9))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "Table 3.2: FACT_REVERSALS (Audit Fact Mart)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    r_items = [
        ("reversal_id (BIGINT, PK): ", "Surrogate primary key for return audit records."),
        ("invoice_no (VARCHAR(20)): ", "Cancellation invoice matching regex '^C[0-9]+'."),
        ("stock_code (VARCHAR(20), FK): ", "Foreign key referencing returned product SKU."),
        ("return_quantity (INTEGER): ", "Negative units returned (Quantity < 0)."),
        ("reversal_date (TIMESTAMP, FK): ", "Timestamp when cancellation was processed."),
        ("unit_price (DECIMAL(10,2)): ", "Original sale unit price of returned item."),
        ("customer_id (VARCHAR(20), FK): ", "Foreign key referencing customer returning items."),
        ("Operational Integrity: ", "Quarantined in fact_reversals_audit.csv to ensure return events never distort positive spend metrics.")
    ]
    for lbl, desc in r_items:
        p = tf2.add_paragraph()
        p.space_before = Pt(3)
        r0 = p.add_run()
        r0.text = "• " + lbl
        r0.font.bold = True
        r0.font.size = Pt(9)
        r0.font.color.rgb = COLOR_NAVY
        r1 = p.add_run()
        r1.text = desc
        r1.font.size = Pt(8.5)
        r1.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 6: 3.1 DATA DICTIONARY - DIMENSION TABLES
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    add_header(slide6, "Section 3.1  •  Data Dictionary", "Data Dictionary: Dimension Tables (DIM_CUSTOMERS & DIM_PRODUCTS)", 6)
    add_footer(slide6)

    # Left: DIM_CUSTOMERS (Width 6.0")
    add_card(slide6, Inches(0.8), Inches(1.5), Inches(6.0), Inches(5.2))
    tb1 = slide6.shapes.add_textbox(Inches(1.0), Inches(1.65), Inches(5.6), Inches(4.9))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p = tf1.paragraphs[0]
    p.text = "Table 3.3: DIM_CUSTOMERS (RFM Behavioral Mart)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    c_items = [
        ("customer_key (BIGINT, PK): ", "Surrogate primary key for customer dimension."),
        ("customer_id (VARCHAR(20)): ", "Unique authenticated natural customer ID."),
        ("recency_days (INTEGER): ", "Days elapsed since last order relative to T_obs."),
        ("frequency_orders (INTEGER): ", "Count of distinct checkout trips (unique orders)."),
        ("monetary_total (DECIMAL(12,2)): ", "Cumulative net customer spend across all valid orders."),
        ("avg_order_value (DECIMAL(10,2)): ", "Average Order Value: Monetary / Frequency."),
        ("r_score, f_score, m_score (SMALLINT): ", "Deterministic 1 to 5 quantile rank scores (pd.qcut)."),
        ("segment_tier (VARCHAR(40)): ", "Mapped business cohort ('Champions', 'At Risk Whales').")
    ]
    for lbl, desc in c_items:
        p = tf1.add_paragraph()
        p.space_before = Pt(3)
        r0 = p.add_run()
        r0.text = "• " + lbl
        r0.font.bold = True
        r0.font.size = Pt(9)
        r0.font.color.rgb = COLOR_NAVY
        r1 = p.add_run()
        r1.text = desc
        r1.font.size = Pt(8.5)
        r1.font.color.rgb = COLOR_DARK

    # Right: DIM_PRODUCTS & DIM_DATE (Width 5.4")
    add_card(slide6, Inches(7.1), Inches(1.5), Inches(5.433), Inches(5.2))
    tb2 = slide6.shapes.add_textbox(Inches(7.3), Inches(1.65), Inches(5.0), Inches(4.9))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "Table 3.4 & 3.5: DIM_PRODUCTS & DIM_DATE"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    p_items = [
        ("DIM_PRODUCTS Catalog Mart: ", ""),
        ("  • product_key (BIGINT, PK): ", "Surrogate primary key for product catalog."),
        ("  • stock_code (VARCHAR(20)): ", "Unique SKU inventory identifier."),
        ("  • description (VARCHAR(255)): ", "Cleaned, sanitized item description text."),
        ("  • unit_price_ref (DECIMAL): ", "Baseline catalog reference price."),
        ("DIM_DATE Temporal Mart: ", ""),
        ("  • date_key (INTEGER, PK): ", "YYYYMMDD formatted integer key (e.g. 20261013)."),
        ("  • full_date (DATE): ", "Standard ISO-8601 calendar date."),
        ("  • day_of_week, month, quarter: ", "Calendar attributes for cohort time-series."),
        ("  • is_weekend (BOOLEAN): ", "Flag isolating weekend purchasing behavioral shifts.")
    ]
    for lbl, desc in p_items:
        p = tf2.add_paragraph()
        p.space_before = Pt(2)
        r0 = p.add_run()
        r0.text = lbl
        r0.font.bold = True
        r0.font.size = Pt(9)
        r0.font.color.rgb = COLOR_NAVY
        r1 = p.add_run()
        r1.text = desc
        r1.font.size = Pt(8.5)
        r1.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 7: 4.1 IMPLEMENTATION SCREEN LAYOUTS
    # =========================================================================
    slide7 = prs.slides.add_slide(blank_layout)
    add_header(slide7, "Section 4.1  •  Implementation", "Screen Layouts: Executive BI Dashboard Architecture", 7)
    add_footer(slide7)

    # Left: Visual Layout Mockup (Width 6.8")
    add_card(slide7, Inches(0.8), Inches(1.5), Inches(6.8), Inches(5.2))
    scr_img = img_dir / "fig_screen_layouts.png"
    if scr_img.exists():
        slide7.shapes.add_picture(str(scr_img), Inches(0.95), Inches(1.7), width=Inches(6.5))
    tb_cap = slide7.shapes.add_textbox(Inches(0.8), Inches(6.25), Inches(6.8), Inches(0.35))
    p = tb_cap.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "Figure 4.1: Implementation Screen Layouts (4 Synchronized Views)"
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_SLATE

    # Right: Screen Descriptions (Width 4.6")
    add_card(slide7, Inches(7.9), Inches(1.5), Inches(4.633), Inches(5.2))
    tb_desc = slide7.shapes.add_textbox(Inches(8.1), Inches(1.65), Inches(4.2), Inches(4.9))
    tf = tb_desc.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "4 Synchronized BI Dashboard Views"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    scr_points = [
        ("Screen 1: Executive KPI Scorecards: ", "High-level summary displaying gross spend ($65,049.20), active customer accounts (256), average order value ($254.10), and at-risk revenue share (7.3%)."),
        ("Screen 2: RFM Behavioral Scatterplot: ", "Interactive 2D/3D canvas mapping Recency against Monetary Spend with bubble radii scaled to Frequency."),
        ("Screen 3: Customer Activation Ledger: ", "Filterable data table with instant search, pagination, and targeted CSV audience export controls."),
        ("Screen 4: Reverse Logistics Audit Ledger: ", "Isolated compliance table displaying quarantined cancellations and credit notes.")
    ]
    for s_lbl, s_desc in scr_points:
        p = tf.add_paragraph()
        p.space_before = Pt(5)
        r0 = p.add_run()
        r0.text = "• " + s_lbl
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        r0.font.color.rgb = COLOR_NAVY
        r1 = p.add_run()
        r1.text = s_desc
        r1.font.size = Pt(9)
        r1.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 8: 4.1 OPERATIONAL INPUT VALIDATIONS
    # =========================================================================
    slide8 = prs.slides.add_slide(blank_layout)
    add_header(slide8, "Section 4.1  •  Data Integrity", "Implementation: Operational Input Validations (VAL-01 to VAL-04)", 8)
    add_footer(slide8)

    # 4 Validation Cards
    v_cards = [
        ("VAL-01: Scorecard Numerical Integrity Validation",
         "Enforces non-negative monetary constraints (Spend >= 0.00). Active customer counts strictly verify unique authenticated non-null CustomerIDs. At-risk percentages are evaluated as: (Sum(Spend_AtRisk) / Total_Spend) * 100.",
         COLOR_MAROON, RGBColor(254, 242, 242)),
        ("VAL-02: Spatial Canvas Coordinate & Boundary Constraints",
         "Interactive scatterplot coordinates enforce strict domain bounds: Recency in [1..365], Frequency in [1..50], Spend in [0..10,000]. Hover tooltips execute null-safety checks before canvas event firing.",
         COLOR_BLUE, RGBColor(239, 246, 255)),
        ("VAL-03: Search Query Sanitization & Pagination Bounds",
         "Client-side search input is sanitized against cross-site scripting (XSS) and SQL injection. Table pagination enforces bounded integer intervals (1 <= Page <= MaxPage).",
         COLOR_EMERALD, RGBColor(236, 253, 245)),
        ("VAL-04: Reverse Logistics Quarantine Constraint",
         "Cancellation transactions strictly enforce regex constraint (InvoiceNo.startswith('C')) and negative quantities (Quantity < 0) before insertion into audit stores.",
         COLOR_NAVY, RGBColor(248, 250, 252))
    ]

    card_w = Inches(11.733)
    c_h = Inches(1.15)
    c_gap = Inches(0.18)
    y_start = Inches(1.5)

    for idx, (v_title, v_desc, v_color, v_bg) in enumerate(v_cards):
        y_pos = y_start + (c_h + c_gap) * idx
        add_card(slide8, Inches(0.8), y_pos, card_w, c_h, bg_color=v_bg, border_color=COLOR_BORDER)

        tb = slide8.shapes.add_textbox(Inches(1.0), y_pos + Inches(0.08), card_w - Inches(0.4), c_h - Inches(0.16))
        tf = tb.text_frame
        tf.word_wrap = True

        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = v_title
        r0.font.name = 'Arial'
        r0.font.size = Pt(11)
        r0.font.bold = True
        r0.font.color.rgb = v_color

        p1 = tf.add_paragraph()
        p1.space_before = Pt(3)
        r1 = p1.add_run()
        r1.text = v_desc
        r1.font.name = 'Arial'
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 9: MILESTONE STATUS & PROJECT ROADMAP
    # =========================================================================
    slide9 = prs.slides.add_slide(blank_layout)
    add_header(slide9, "Milestone Tracking  •  CE-502 / PRJ-701", "Project Evaluation Schedule & Milestone Status (Mid-Term)", 9)
    add_footer(slide9)

    add_card(slide9, Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.2))
    tb_m = slide9.shapes.add_textbox(Inches(1.0), Inches(1.65), Inches(11.333), Inches(4.9))
    tf_m = tb_m.text_frame
    tf_m.word_wrap = True

    p = tf_m.paragraphs[0]
    p.text = "Table 4.1: Comprehensive Project Progress Work Schedule"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    milestones = [
        ("24/09/2026", "First Progress Work Evaluation", "Sections 1.1, 1.2, 1.3, 2.1, 2.2", "10 Marks", "COMPLETED & VERIFIED", COLOR_SLATE, False),
        ("13/10/2026", "Mid-Term Examination (Second Progress)", "Sections 2.3, 2.4, 3.1, 4.1", "20 Marks", "DEMONSTRATED (Current Submission)", COLOR_EMERALD, True),
        ("27/10/2026", "Third Progress Work Evaluation", "Sections 4.2 (Sample Code), 4.3 (Navigation), 5.1-5.2, 6.1", "10 Marks", "Scheduled for Third Milestone", COLOR_BLUE, False),
        ("06/11/2026", "Report Submission (Spiral Binding)", "Final Signed Spiral Copy & Media", "10 Marks", "Scheduled for Final Submission", COLOR_SLATE, False),
        ("18-20/11/2026", "Final Examination (SEE)", "Project Execution Demonstration & Viva Voce", "50 Marks", "Scheduled for Semester Examination", COLOR_MAROON, False)
    ]

    for dt, part, scope, mrk, stat, clr, is_active in milestones:
        p = tf_m.add_paragraph()
        p.space_before = Pt(8)

        r0 = p.add_run()
        r0.text = f"• [{dt}]  {part} ({mrk}) — "
        r0.font.bold = True
        r0.font.size = Pt(11)
        r0.font.color.rgb = clr

        r1 = p.add_run()
        r1.text = stat + "\n"
        r1.font.bold = True
        r1.font.size = Pt(10.5)
        r1.font.color.rgb = clr

        r2 = p.add_run()
        r2.text = f"   Curriculum Scope: {scope}"
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = COLOR_DARK

    add_card(slide9, Inches(1.1), Inches(5.6), Inches(11.133), Inches(0.9), bg_color=RGBColor(240, 253, 244), border_color=COLOR_EMERALD)
    tb_call = slide9.shapes.add_textbox(Inches(1.2), Inches(5.65), Inches(10.9), Inches(0.8))
    p = tb_call.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = "MID-TERM MILESTONE STATUS (13/10/2026): "
    r.font.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = COLOR_EMERALD
    r_txt = p.add_run()
    r_txt.text = "Cumulative completion of 30 Marks (First Progress 10 Marks + Mid-Term 20 Marks). All Class, Sequence, Data Dictionary, and UI Validation specifications fully implemented."
    r_txt.font.size = Pt(9.5)
    r_txt.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 10: THANK YOU SLIDE (CONCLUSION & Q&A)
    # =========================================================================
    slide10 = prs.slides.add_slide(blank_layout)

    top_bar = slide10.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = COLOR_MAROON
    top_bar.line.color.rgb = COLOR_MAROON

    bot_bar = slide10.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(7.38), Inches(13.333), Inches(0.12))
    bot_bar.fill.solid()
    bot_bar.fill.fore_color.rgb = COLOR_NAVY
    bot_bar.line.color.rgb = COLOR_NAVY

    if iar_banner.exists():
        slide10.shapes.add_picture(str(iar_banner), Inches(3.9), Inches(0.5), width=Inches(5.5))

    tb_thx = slide10.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.333), Inches(1.4))
    p_thx = tb_thx.text_frame.paragraphs[0]
    p_thx.alignment = PP_ALIGN.CENTER
    p_thx.text = "THANK YOU!"
    p_thx.font.name = 'Arial'
    p_thx.font.size = Pt(46)
    p_thx.font.bold = True
    p_thx.font.color.rgb = COLOR_MAROON

    p_qa = tb_thx.text_frame.add_paragraph()
    p_qa.alignment = PP_ALIGN.CENTER
    p_qa.text = "Questions & Technical Discussion"
    p_qa.font.name = 'Arial'
    p_qa.font.size = Pt(20)
    p_qa.font.bold = True
    p_qa.font.color.rgb = COLOR_NAVY

    add_card(slide10, Inches(2.5), Inches(3.4), Inches(8.333), Inches(3.4), bg_color=COLOR_CARD_BG, border_color=COLOR_BORDER)
    tb_end = slide10.shapes.add_textbox(Inches(2.8), Inches(3.55), Inches(7.733), Inches(3.1))
    tf_end = tb_end.text_frame
    tf_end.word_wrap = True

    p = tf_end.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "E-COMMERCE CUSTOMER LIFECYCLE & RFM SEGMENTATION DASHBOARD"
    p.font.name = 'Arial'
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_NAVY

    p_div = tf_end.add_paragraph()
    p_div.alignment = PP_ALIGN.CENTER
    p_div.text = "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    p_div.font.size = Pt(8.5)
    p_div.font.color.rgb = COLOR_MAROON

    end_details = [
        ("Candidate: ", "Mr. Deep Koshiya (B.Sc. in AI & Data Science, Semester V)"),
        ("Institution: ", "Institute of Advanced Research (IAR), Gandhinagar"),
        ("Assessment Milestone: ", "Second Progress Work / Mid-Term Exam (13/10/2026  •  20 Marks)"),
        ("Course & Project: ", "Project - 1 (CE-502 / PRJ-701)"),
        ("GitHub Repository: ", "https://github.com/Deep-k-coder/ecommerce-rfm-analytics"),
        ("Project Status: ", "Class & Sequence Diagrams, Star Schema Marts & UI Validations Complete")
    ]
    for lbl, val in end_details:
        p = tf_end.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        p.space_before = Pt(3)
        r0 = p.add_run()
        r0.text = lbl
        r0.font.bold = True
        r0.font.size = Pt(10)
        r0.font.color.rgb = COLOR_DARK
        r1 = p.add_run()
        r1.text = val
        r1.font.size = Pt(10)
        r1.font.color.rgb = COLOR_SLATE

    out_pptx_path = base_dir / "Deep_Koshiya_ProjectPresentation_Progress2.pptx"
    prs.save(str(out_pptx_path))
    prs.save(str(base_dir / "Deep_Koshiya_ProjectPresentation.pptx"))
    print(f"[SUCCESS] Phase 2 Presentation saved to {out_pptx_path}")

if __name__ == "__main__":
    create_deck()
