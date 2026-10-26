"""
build_phase3_presentation.py
Generates the comprehensive 10-slide PowerPoint presentation (.pptx) for the
Third Phase Work (Third Progress Work Evaluation - 10 Marks, Cumulative 40 Marks):
- Degree: B.Sc. in Artificial Intelligence & Data Science, Semester V (5th Semester)
- Institution: Institute of Advanced Research (IAR), Gandhinagar
- NO DEPARTMENT NAME mentioned anywhere in the presentation.
- Official IAR banner at top of Title & Thank You slides.
- 16:9 Widescreen format (13.333" x 7.5").
- Exactly 10 slides, culminating in an elegant "Thank You & Q&A" slide.
- High-res Screen Layouts, Diagnostic Profiles, and Navigation Visuals.
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
        p_foot.text = "B.Sc. in Artificial Intelligence & Data Science (Sem V)  •  Institute of Advanced Research  •  Third Progress Work (27/10/2026)"
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

    badge = add_card(slide1, Inches(2.9), Inches(2.35), Inches(7.533), Inches(0.42), bg_color=RGBColor(236, 253, 245), border_color=COLOR_EMERALD)
    tb_badge = slide1.shapes.add_textbox(Inches(2.9), Inches(2.35), Inches(7.533), Inches(0.42))
    p_b = tb_badge.text_frame.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    p_b.text = "THIRD PROGRESS WORK EVALUATION  •  27TH OCTOBER 2026  •  40 MARKS CUMULATIVE"
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
    p_sub.text = "Third Phase Technical Presentation: Sample Coding, Cybersecurity Controls, Execution & Navigation"
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
    p.text = "EVALUATION & MILESTONE SCOPE"
    p.font.name = 'Arial'
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    lines_c2 = [
        ("Milestone: ", "Third Progress Work Evaluation"),
        ("Evaluation Date: ", "27th October 2026 (10 AM - 5 PM)"),
        ("Curriculum Scope: ", "Sections 4.2, 4.3, 5.1, 5.2, 6.1"),
        ("Cumulative Marks: ", "40 Marks Evaluated (10M + 20M + 10M)"),
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
    # SLIDE 2: PROJECT OVERVIEW & CUMULATIVE ARCHITECTURE
    # =========================================================================
    slide2 = prs.slides.add_slide(blank_layout)
    add_header(slide2, "Progressive Architecture  •  Overview", "Cumulative Project Architecture (Chapters 1 to 3 Summary)", 2)
    add_footer(slide2)

    col_w = Inches(3.64)
    gap = Inches(0.4)
    y_top = Inches(1.5)
    card_h = Inches(5.2)

    # Card 1: Chapter 1 Summary
    add_card(slide2, Inches(0.8), y_top, col_w, card_h)
    tb = slide2.shapes.add_textbox(Inches(0.95), y_top + Inches(0.15), col_w - Inches(0.3), card_h - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "1. About the System (Ch. 1)"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    b1 = [
        ("Problem Definition: ", "Surging customer acquisition costs and lack of longitudinal cohort tracking in OLTP order logs."),
        ("Requirement Matrix: ", "FR-01 to FR-06 data contracts, sub-5s latency, deterministic reproducibility (`seed=42`)."),
        ("Technology Stack: ", "Python 3.9, Pandas 2.3, NumPy 2.0, Scikit-Learn 1.6, SciPy, Tailwind CSS, Chart.js."),
        ("Architectural Decoupling: ", "Raw data refinery separated from dimensional marts and presentation layers.")
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

    # Card 2: Chapter 2 Summary
    add_card(slide2, Inches(0.8) + col_w + gap, y_top, col_w, card_h)
    tb = slide2.shapes.add_textbox(Inches(0.95) + col_w + gap, y_top + Inches(0.15), col_w - Inches(0.3), card_h - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "2. UML & DFD Design (Ch. 2)"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    b2 = [
        ("Level 0, 1, 2 DFDs: ", "Global system boundary, 5-stage transformation pipeline, and algorithmic decomposition."),
        ("Use Case Model: ", "6 formal operational capabilities (UC-01 to UC-06) across 4 primary actors."),
        ("Class Diagram: ", "Object domain model: TransactionRecord, CleanRefineryPipeline, RFMFeatureEngineer, KMeansEngine."),
        ("Sequence Diagram: ", "11-step synchronous/asynchronous execution protocol across lifelines.")
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

    # Card 3: Chapter 3 Summary
    add_card(slide2, Inches(0.8) + (col_w + gap) * 2, y_top, col_w, card_h, bg_color=RGBColor(240, 253, 244), border_color=COLOR_EMERALD)
    tb = slide2.shapes.add_textbox(Inches(0.95) + (col_w + gap) * 2, y_top + Inches(0.15), col_w - Inches(0.3), card_h - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "3. Star-Schema Marts (Ch. 3)"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_EMERALD

    b3 = [
        ("FACT_TRANSACTIONS: ", "Granular order facts with gross spend and foreign keys to dimensions."),
        ("FACT_REVERSALS: ", "Audit ledger isolating credit notes ('^C') and negative items."),
        ("DIM_CUSTOMERS: ", "Longitudinal customer profiles with Recency, Frequency, Spend, and 8 business tiers."),
        ("DIM_PRODUCTS & DATE: ", "Standardized SKU reference catalog and temporal attributes.")
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
    # SLIDE 3: 4.1 SCREEN LAYOUTS & VALIDATIONS
    # =========================================================================
    slide3 = prs.slides.add_slide(blank_layout)
    add_header(slide3, "Section 4.1  •  Implementation UI", "Screen Layouts & Operational Input Validations", 3)
    add_footer(slide3)

    add_card(slide3, Inches(0.8), Inches(1.5), Inches(6.8), Inches(5.2))
    scr_img = img_dir / "fig_screen_layouts.png"
    if scr_img.exists():
        slide3.shapes.add_picture(str(scr_img), Inches(0.95), Inches(1.7), width=Inches(6.5))
    tb_cap = slide3.shapes.add_textbox(Inches(0.8), Inches(6.25), Inches(6.8), Inches(0.35))
    p = tb_cap.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "Figure 4.1: Implementation Screen Layouts (4 Synchronized UI Views)"
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_SLATE

    add_card(slide3, Inches(7.9), Inches(1.5), Inches(4.633), Inches(5.2))
    tb_desc = slide3.shapes.add_textbox(Inches(8.1), Inches(1.65), Inches(4.2), Inches(4.9))
    tf = tb_desc.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Input & Operational Validations"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    v_points = [
        ("VAL-01 (Scorecard Integrity): ", "Enforces non-negative spend constraints (Spend >= 0.00) and cross-checks active counts against distinct non-null IDs."),
        ("VAL-02 (Canvas Boundary Constraints): ", "Scatterplot coordinates bounded to Recency [1..365], Frequency [1..50], Spend [0..10,000]. Hover tooltips validate customer existence."),
        ("VAL-03 (Search Sanitization): ", "Client search input sanitized against script injection. Table pagination enforces integer bounds (1 <= Page <= MaxPage)."),
        ("VAL-04 (Audit Quarantine Logic): ", "Reversals strictly enforce cancellation regex ('^C') and negative quantities (Quantity < 0) before insertion into audit stores.")
    ]
    for v_lbl, v_desc in v_points:
        p = tf.add_paragraph()
        p.space_before = Pt(6)
        r0 = p.add_run()
        r0.text = "• " + v_lbl
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        r0.font.color.rgb = COLOR_NAVY
        r1 = p.add_run()
        r1.text = v_desc
        r1.font.size = Pt(9)
        r1.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 4: 4.2 SAMPLE CODING ARCHITECTURE
    # =========================================================================
    slide4 = prs.slides.add_slide(blank_layout)
    add_header(slide4, "Section 4.2  •  Implementation Code", "Sample Coding: Modular Pipeline Architecture (Excerpts)", 4)
    add_footer(slide4)

    card_w = Inches(11.733)
    c_h = Inches(1.6)
    c_gap = Inches(0.18)
    y_start = Inches(1.5)

    code_excerpts = [
        ("1. Data Refinery & Anomaly Quarantine (clean_refinery_pipeline.py)",
         "is_cancellation = df['InvoiceNo'].str.startswith('C', na=False) | (df['Quantity'] <= 0)\nreversals_df = df[is_cancellation].copy()\nsales_df = df[~is_cancellation].copy()\nreversals_df.to_csv(self.output_dir / 'fact_reversals_audit.csv', index=False)",
         COLOR_MAROON, RGBColor(254, 242, 242)),
        ("2. Vectorized RFM Feature Extraction (rfm_feature_engineering.py)",
         "snapshot_date = df['InvoiceDate'].max() + pd.Timedelta(days=1)\nrfm = df.groupby('CustomerID').agg({\n    'InvoiceDate': lambda x: (snapshot_date - x.max()).days,\n    'InvoiceNo': 'nunique', 'LineTotal': 'sum'\n}).reset_index()",
         COLOR_BLUE, RGBColor(239, 246, 255)),
        ("3. Skew-Resistant Feature Scaling & K-Means (segmentation_engine.py)",
         "rfm_log = np.log1p(rfm_df[['Recency', 'Frequency', 'Monetary']])\nscaler = StandardScaler()\nrfm_scaled = scaler.fit_transform(rfm_log)\nkmeans = KMeans(n_clusters=4, init='k-means++', random_state=42)\nrfm_df['Cluster'] = kmeans.fit_predict(rfm_scaled)",
         COLOR_EMERALD, RGBColor(236, 253, 245))
    ]

    for idx, (title, code_str, tag_col, bg_col) in enumerate(code_excerpts):
        y_pos = y_start + (c_h + c_gap) * idx
        add_card(slide4, Inches(0.8), y_pos, card_w, c_h, bg_color=bg_col, border_color=COLOR_BORDER)

        tb = slide4.shapes.add_textbox(Inches(1.0), y_pos + Inches(0.08), card_w - Inches(0.4), c_h - Inches(0.16))
        tf = tb.text_frame
        tf.word_wrap = True

        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = title
        r0.font.name = 'Arial'
        r0.font.size = Pt(11)
        r0.font.bold = True
        r0.font.color.rgb = tag_col

        p1 = tf.add_paragraph()
        p1.space_before = Pt(3)
        r1 = p1.add_run()
        r1.text = code_str
        r1.font.name = 'Courier New'
        r1.font.size = Pt(8.5)
        r1.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 5: 4.2 CYBERSECURITY CONTROLS & DIAGNOSTICS
    # =========================================================================
    slide5 = prs.slides.add_slide(blank_layout)
    add_header(slide5, "Section 4.2  •  Security & Validation", "Cybersecurity Controls & Algorithmic Diagnostics", 5)
    add_footer(slide5)

    # Left: Diagnostics Image (Width 6.8")
    add_card(slide5, Inches(0.8), Inches(1.5), Inches(6.8), Inches(5.2))
    diag_img = img_dir / "fig_kmeans_diagnostics.png"
    if diag_img.exists():
        slide5.shapes.add_picture(str(diag_img), Inches(0.95), Inches(1.7), width=Inches(6.5))
    tb_cap = slide5.shapes.add_textbox(Inches(0.8), Inches(6.25), Inches(6.8), Inches(0.35))
    p = tb_cap.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "Figure 4.2: K-Means Inertia Elbow Curve & Silhouette Diagnostic Profile"
    p.font.size = Pt(9.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_SLATE

    # Right: Cybersecurity Controls (Width 4.6")
    add_card(slide5, Inches(7.9), Inches(1.5), Inches(4.633), Inches(5.2))
    tb_desc = slide5.shapes.add_textbox(Inches(8.1), Inches(1.65), Inches(4.2), Inches(4.9))
    tf = tb_desc.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "5 Core Cybersecurity Tasks"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    sec_points = [
        ("1. PII Masking & Privacy: ", "Customer IDs synthesized into opaque alphanumeric strings (CUST-14911), preventing personal data leakage."),
        ("2. Input Sanitization: ", "Client-side search sanitization rejects special characters, mitigating script injection and XSS."),
        ("3. Algorithmic Determinism: ", "Fixed random state (seed=42) guarantees 100% reproducible cluster centroids across runtime environments."),
        ("4. Immutable Audit Ledger: ", "Quarantined returns stored in fact_reversals_audit.csv with row integrity checks."),
        ("5. Safe Export Protocols: ", "CSV audience downloads enforce header whitelisting to eliminate CSV formula injection vulnerabilities.")
    ]
    for s_lbl, s_desc in sec_points:
        p = tf.add_paragraph()
        p.space_before = Pt(4)
        r0 = p.add_run()
        r0.text = s_lbl
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        r0.font.color.rgb = COLOR_NAVY
        r1 = p.add_run()
        r1.text = s_desc
        r1.font.size = Pt(8.8)
        r1.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 6: 4.3 STEPS OF PROJECT EXECUTION
    # =========================================================================
    slide6 = prs.slides.add_slide(blank_layout)
    add_header(slide6, "Section 4.3  •  Operations", "Steps of the Execution Project (Automated Pipeline Protocol)", 6)
    add_footer(slide6)

    card_w = Inches(11.733)
    step_h = Inches(0.92)
    step_gap = Inches(0.12)
    y_start = Inches(1.5)

    steps = [
        ("Step 1: Environment Setup & Dependency Verification",
         "source venv/bin/activate && pip install -r requirements.txt",
         "Initializes Python 3.9 virtual environment and verifies Pandas, NumPy, Scikit-Learn, and SciPy packages."),
        ("Step 2: Synthetic Event Generation & Ingestion",
         "python generate_dataset.py",
         "Generates 15,000+ realistic transaction rows with Pareto spend distributions, cancellations, and guest sessions."),
        ("Step 3: End-to-End Pipeline & Star-Schema Build",
         "python main.py",
         "Executes data refinery, vectorized RFM feature extraction, K-Means clustering, and star-schema CSV mart exports."),
        ("Step 4: Automated Testing & Quality Assurance",
         "pytest tests/ -v",
         "Executes unit and integration test suite validating schema typing, Tukey IQR fences, and cluster bounds."),
        ("Step 5: Executive Dashboard Launch & Hosting",
         "python -m http.server 8000 --directory dashboard/",
         "Launches client-side executive dashboard on port 8000, serving KPI scorecards and interactive scatterplots.")
    ]

    for idx, (title, cmd, desc) in enumerate(steps):
        y_pos = y_start + (step_h + step_gap) * idx
        add_card(slide6, Inches(0.8), y_pos, card_w, step_h)

        tb = slide6.shapes.add_textbox(Inches(1.0), y_pos + Inches(0.06), card_w - Inches(0.4), step_h - Inches(0.12))
        tf = tb.text_frame
        tf.word_wrap = True

        p0 = tf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = title + "  "
        r0.font.bold = True
        r0.font.size = Pt(10.5)
        r0.font.color.rgb = COLOR_MAROON

        r1 = p0.add_run()
        r1.text = "• " + desc
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = COLOR_DARK

        p1 = tf.add_paragraph()
        p1.space_before = Pt(2)
        r2 = p1.add_run()
        r2.text = "$ " + cmd
        r2.font.name = 'Courier New'
        r2.font.size = Pt(8.5)
        r2.font.bold = True
        r2.font.color.rgb = COLOR_BLUE

    # =========================================================================
    # SLIDE 7: 4.3 DASHBOARD NAVIGATION & VISUALIZATION
    # =========================================================================
    slide7 = prs.slides.add_slide(blank_layout)
    add_header(slide7, "Section 4.3  •  Dashboard Navigation", "Interactive Dashboard Navigation & Visual Analytics", 7)
    add_footer(slide7)

    # Left: Revenue Donut (Width 5.6")
    add_card(slide7, Inches(0.8), Inches(1.5), Inches(5.6), Inches(5.2))
    rev_img = img_dir / "fig_revenue_share.png"
    if rev_img.exists():
        slide7.shapes.add_picture(str(rev_img), Inches(1.0), Inches(1.7), width=Inches(5.2))
    tb_cap = slide7.shapes.add_textbox(Inches(0.8), Inches(6.25), Inches(5.6), Inches(0.35))
    p = tb_cap.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = "Figure 4.3: Segment Revenue Share Donut (72.5% Champions)"
    p.font.size = Pt(9)
    p.font.bold = True
    p.font.color.rgb = COLOR_SLATE

    # Right: Spatial Scatterplot (Width 5.8")
    add_card(slide7, Inches(6.7), Inches(1.5), Inches(5.833), Inches(5.2))
    spat_img = img_dir / "fig_spatial_scatter.png"
    if spat_img.exists():
        slide7.shapes.add_picture(str(spat_img), Inches(6.9), Inches(1.7), width=Inches(5.4))
    tb_cap2 = slide7.shapes.add_textbox(Inches(6.7), Inches(6.25), Inches(5.833), Inches(0.35))
    p2 = tb_cap2.text_frame.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    p2.text = "Figure 4.4: Multi-Dimensional Spatial RFM Bubble Scatterplot"
    p2.font.size = Pt(9)
    p2.font.bold = True
    p2.font.color.rgb = COLOR_SLATE

    # =========================================================================
    # SLIDE 8: 5.1 & 5.2 CONCLUSION & FUTURE ENHANCEMENTS
    # =========================================================================
    slide8 = prs.slides.add_slide(blank_layout)
    add_header(slide8, "Chapter 5  •  Conclusion & Roadmaps", "Importance of Worked-Out Tasks & Future Enhancements", 8)
    add_footer(slide8)

    # Left: 5.1 Importance
    add_card(slide8, Inches(0.8), Inches(1.5), Inches(5.7), Inches(5.2))
    tb1 = slide8.shapes.add_textbox(Inches(1.0), Inches(1.65), Inches(5.3), Inches(4.9))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    p = tf1.paragraphs[0]
    p.text = "5.1 Importance of Worked-Out Tasks"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    imp_points = [
        ("Elimination of Analytical Distortion: ", "Isolating returns into fact_reversals_audit.csv and capping bulk outliers ensures RFM reflects true consumer buying behavior."),
        ("Commercial Margin Optimization: ", "Replacing blanket discounts with 8 behavioral tiers protects gross profit margins and focuses spend on high-value retention."),
        ("Enterprise Dimensional Governance: ", "Star-schema dimensional mart allows commercial teams to query multi-dimensional customer features using SQL, Tableau, and Power BI."),
        ("Operational Activation: ", "Interactive scorecards and CSV activation lists bridge the gap between machine learning models and daily marketing campaigns.")
    ]
    for lbl, desc in imp_points:
        p = tf1.add_paragraph()
        p.space_before = Pt(6)
        r0 = p.add_run()
        r0.text = "• " + lbl
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        r0.font.color.rgb = COLOR_NAVY
        r1 = p.add_run()
        r1.text = desc
        r1.font.size = Pt(9)
        r1.font.color.rgb = COLOR_DARK

    # Right: 5.2 Future Enhancements
    add_card(slide8, Inches(6.8), Inches(1.5), Inches(5.733), Inches(5.2))
    tb2 = slide8.shapes.add_textbox(Inches(7.0), Inches(1.65), Inches(5.3), Inches(4.9))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "5.2 Further Enhancement in Future"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    fut_points = [
        ("Deep Learning CLV Forecasting: ", "Implementing Recurrent Neural Networks (LSTM) and Transformers to predict customer lifetime value and churn probabilities."),
        ("Real-Time Streaming Ingestion: ", "Migrating batch ingestion to Apache Kafka and Apache Spark Streaming for sub-second transaction processing."),
        ("Automated CRM Webhooks: ", "Building RESTful webhooks to dispatch targeted cohorts automatically into HubSpot, Salesforce Marketing Cloud, and Klaviyo."),
        ("Multi-Touch Attribution: ", "Incorporating Shapley Value and Markov Chain attribution models to quantify marketing touchpoint impact across channels.")
    ]
    for lbl, desc in fut_points:
        p = tf2.add_paragraph()
        p.space_before = Pt(6)
        r0 = p.add_run()
        r0.text = "• " + lbl
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        r0.font.color.rgb = COLOR_NAVY
        r1 = p.add_run()
        r1.text = desc
        r1.font.size = Pt(9)
        r1.font.color.rgb = COLOR_DARK

    # =========================================================================
    # SLIDE 9: 6.1 REFERENCES & MILESTONE SCHEDULE
    # =========================================================================
    slide9 = prs.slides.add_slide(blank_layout)
    add_header(slide9, "Milestone Tracking  •  CE-502 / PRJ-701", "References & Complete Project Progress Work Schedule", 9)
    add_footer(slide9)

    add_card(slide9, Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.2))
    tb_m = slide9.shapes.add_textbox(Inches(1.0), Inches(1.65), Inches(11.333), Inches(4.9))
    tf_m = tb_m.text_frame
    tf_m.word_wrap = True

    p = tf_m.paragraphs[0]
    p.text = "Table 6.1: Complete Project Progress Work Evaluation Schedule"
    p.font.size = Pt(13.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_MAROON

    milestones = [
        ("24/09/2026", "First Progress Work Evaluation", "Sections 1.1, 1.2, 1.3, 2.1, 2.2", "10 Marks", "COMPLETED & VERIFIED", COLOR_SLATE, False),
        ("13/10/2026", "Mid-Term Examination (Second Progress)", "Sections 2.3, 2.4, 3.1, 4.1", "20 Marks", "COMPLETED & VERIFIED", COLOR_SLATE, False),
        ("27/10/2026", "Third Progress Work Evaluation", "Sections 4.2, 4.3, 5.1, 5.2, 6.1", "10 Marks", "DEMONSTRATED (Current Submission)", COLOR_EMERALD, True),
        ("06/11/2026", "Report Submission (Spiral Binding)", "Final Signed Spiral Copy & Media", "10 Marks", "Scheduled for Final Submission", COLOR_BLUE, False),
        ("18-20/11/2026", "Final Examination (SEE)", "Project Execution Demonstration & Viva Voce", "50 Marks", "Scheduled for Semester Examination", COLOR_MAROON, False)
    ]

    for dt, part, scope, mrk, stat, clr, is_active in milestones:
        p = tf_m.add_paragraph()
        p.space_before = Pt(7)

        r0 = p.add_run()
        r0.text = f"• [{dt}]  {part} ({mrk}) — "
        r0.font.bold = True
        r0.font.size = Pt(10.5)
        r0.font.color.rgb = clr

        r1 = p.add_run()
        r1.text = stat + "\n"
        r1.font.bold = True
        r1.font.size = Pt(10)
        r1.font.color.rgb = clr

        r2 = p.add_run()
        r2.text = f"   Curriculum Scope: {scope}"
        r2.font.size = Pt(9)
        r2.font.color.rgb = COLOR_DARK

    add_card(slide9, Inches(1.1), Inches(5.6), Inches(11.133), Inches(0.9), bg_color=RGBColor(240, 253, 244), border_color=COLOR_EMERALD)
    tb_call = slide9.shapes.add_textbox(Inches(1.2), Inches(5.65), Inches(10.9), Inches(0.8))
    p = tb_call.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = "THIRD PROGRESS MILESTONE STATUS (27/10/2026): "
    r.font.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = COLOR_EMERALD
    r_txt = p.add_run()
    r_txt.text = "Cumulative completion of 40 Marks across all progress work milestones (10M + 20M + 10M). All technical chapters (Chapters 1 to 6) fully implemented, documented, and validated."
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
        ("Assessment Milestone: ", "Third Progress Work Evaluation (27/10/2026  •  10 Marks)"),
        ("Course & Project: ", "Project - 1 (CE-502 / PRJ-701)"),
        ("GitHub Repository: ", "https://github.com/Deep-k-coder/ecommerce-rfm-analytics"),
        ("Project Status: ", "Chapters 1 through 6 Complete; Ready for Final Spiral Binding & SEE Examination")
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

    out_pptx_path = base_dir / "Deep_Koshiya_ProjectPresentation_Progress3.pptx"
    prs.save(str(out_pptx_path))
    prs.save(str(base_dir / "Deep_Koshiya_ProjectPresentation.pptx"))
    print(f"[SUCCESS] Phase 3 Presentation saved to {out_pptx_path}")

if __name__ == "__main__":
    create_deck()
