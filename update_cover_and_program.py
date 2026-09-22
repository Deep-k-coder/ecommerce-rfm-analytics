"""
update_cover_and_program.py
Updates the project report with:
1. Premium, elegant academic Cover Page with formal layout and borders.
2. Degree updated to: B.Sc. AI and Data Science (Bachelor of Science in Artificial Intelligence & Data Science).
3. Department updated to: Department of Artificial Intelligence & Data Science.
"""

from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, hex_color):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_cell_border(cell, **kwargs):
    """
    kwargs can be top, bottom, left, right.
    val: 'single', 'double', 'dashed', etc.
    color: '000000'
    sz: '12' (1/8 pt)
    """
    tcPr = cell._element.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = f'<w:{edge} {nsdecls("w")} w:val="{edge_data.get("val", "single")}" w:sz="{edge_data.get("sz", "8")}" w:space="0" w:color="{edge_data.get("color", "000000")}"/>'
        else:
            tag = f'<w:{edge} {nsdecls("w")} w:val="none"/>'
        tcBorders.append(parse_xml(tag))
    tcPr.append(tcBorders)

def build_report():
    base_dir = Path("/Users/deep/.gemini/antigravity/scratch/ecommerce_rfm_analytics")
    img_dir = base_dir / "report_images"
    doc = Document()

    # Configure Margins: 1.15 in left (for spiral binding), 0.85 in top/bottom/right
    for section in doc.sections:
        section.top_margin = Inches(0.85)
        section.bottom_margin = Inches(0.85)
        section.left_margin = Inches(1.15)
        section.right_margin = Inches(0.85)

    NAVY = RGBColor(15, 23, 42)
    INDIGO = RGBColor(30, 27, 75)       # Deep Royal Navy #1e1b4b
    ACCENT_BLUE = RGBColor(37, 99, 235) # Vibrant Blue #2563eb
    SLATE = RGBColor(71, 85, 105)

    style_normal = doc.styles['Normal']
    font_normal = style_normal.font
    font_normal.name = 'Times New Roman'
    font_normal.size = Pt(11)
    font_normal.color.rgb = NAVY

    def add_para(text, bold_prefix=None, space_after=4, space_before=0, line_spacing=1.12, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        if bold_prefix:
            r0 = p.add_run(bold_prefix)
            r0.font.name = 'Times New Roman'
            r0.font.size = Pt(11)
            r0.font.bold = True
            r0.font.color.rgb = NAVY
        r1 = p.add_run(text)
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(11)
        r1.font.italic = italic
        r1.font.color.rgb = NAVY
        return p

    def add_h1(text, space_before=10, space_after=4):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(15)
        r.font.bold = True
        r.font.color.rgb = INDIGO
        return p

    def add_h2(text, space_before=8, space_after=3):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.color.rgb = NAVY
        return p

    def add_h3(text, space_before=6, space_after=2):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(11.5)
        r.font.bold = True
        r.font.color.rgb = NAVY
        return p

    def add_caption(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(5)
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = SLATE
        return p

    # =========================================================================
    # PAGE 1: PREMIUM ACADEMIC COVER PAGE
    # =========================================================================
    
    # Outer Border Box table to frame the entire cover page
    outer_table = doc.add_table(rows=1, cols=1)
    outer_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    outer_cell = outer_table.rows[0].cells[0]
    set_cell_background(outer_cell, "FFFFFF")
    set_cell_margins(outer_cell, top=140, bottom=140, left=180, right=180)
    set_cell_border(outer_cell, 
                    top={"val": "double", "sz": "18", "color": "1E1B4B"},
                    bottom={"val": "double", "sz": "18", "color": "1E1B4B"},
                    left={"val": "double", "sz": "18", "color": "1E1B4B"},
                    right={"val": "double", "sz": "18", "color": "1E1B4B"})

    # Content inside the framed cover
    cp0 = outer_cell.paragraphs[0]
    cp0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp0.paragraph_format.space_after = Pt(2)
    r_hdr_inst = cp0.add_run("DEPARTMENT OF ARTIFICIAL INTELLIGENCE & DATA SCIENCE")
    r_hdr_inst.font.name = 'Times New Roman'
    r_hdr_inst.font.size = Pt(13)
    r_hdr_inst.font.bold = True
    r_hdr_inst.font.color.rgb = INDIGO

    cp_deg = outer_cell.add_paragraph()
    cp_deg.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp_deg.paragraph_format.space_after = Pt(16)
    r_deg_text = cp_deg.add_run("B.Sc. in Artificial Intelligence & Data Science | Semester VII\nAcademic Year: 2026 – 2027")
    r_deg_text.font.name = 'Times New Roman'
    r_deg_text.font.size = Pt(11)
    r_deg_text.font.bold = True
    r_deg_text.font.color.rgb = ACCENT_BLUE

    # Thin decorative divider
    cp_div = outer_cell.add_paragraph()
    cp_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp_div.paragraph_format.space_after = Pt(14)
    r_div = cp_div.add_run("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    r_div.font.name = 'Times New Roman'
    r_div.font.size = Pt(10)
    r_div.font.color.rgb = SLATE

    # Title Block
    cp_title = outer_cell.add_paragraph()
    cp_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp_title.paragraph_format.space_after = Pt(8)
    r_t1 = cp_title.add_run("E-COMMERCE CUSTOMER LIFECYCLE &\nRFM SEGMENTATION DASHBOARD")
    r_t1.font.name = 'Times New Roman'
    r_t1.font.size = Pt(18)
    r_t1.font.bold = True
    r_t1.font.color.rgb = INDIGO

    cp_sub = outer_cell.add_paragraph()
    cp_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp_sub.paragraph_format.space_after = Pt(16)
    r_s1 = cp_sub.add_run("An Applied Machine Learning & Analytics Engineering System\nfor Commercial Retention & Behavioral Cohort Optimization")
    r_s1.font.name = 'Times New Roman'
    r_s1.font.size = Pt(11.5)
    r_s1.font.italic = True
    r_s1.font.color.rgb = NAVY

    # Callout Badge Table for First Progress Work (24/09/2026)
    badge_table = outer_cell.add_table(rows=5, cols=2)
    badge_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    b_rows = [
        ("Assessment Milestone:", "FIRST PROGRESS WORK EVALUATION"),
        ("Scheduled Submission Date:", "24th September 2026 (10:00 AM – 5:00 PM)"),
        ("Target Work Completed:", "Sections 1.1, 1.2, 1.3, 2.1, and 2.2 (DFD & Use Case)"),
        ("Course & Project Code:", "Project - 1 (CE-502 / PRJ-701)"),
        ("Allotted Assessment Marks:", "10 Marks (First Progress Work)")
    ]
    for idx, (label, val) in enumerate(b_rows):
        row = badge_table.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        c0.text, c1.text = label, val
        c0.paragraphs[0].runs[0].font.name = 'Times New Roman'
        c0.paragraphs[0].runs[0].font.bold = True
        c0.paragraphs[0].runs[0].font.size = Pt(9.5)
        c1.paragraphs[0].runs[0].font.name = 'Times New Roman'
        c1.paragraphs[0].runs[0].font.size = Pt(9.5)
        set_cell_background(c0, "F8FAFC")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, top=45, bottom=45, left=90, right=90)
        set_cell_margins(c1, top=45, bottom=45, left=90, right=90)

    # Spacing
    cp_sp = outer_cell.add_paragraph()
    cp_sp.paragraph_format.space_after = Pt(18)

    # Scholar & Guidance Information Block
    meta_sub_table = outer_cell.add_table(rows=2, cols=2)
    meta_sub_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_m1 = meta_sub_table.rows[0].cells[0]
    cell_m2 = meta_sub_table.rows[0].cells[1]
    cell_m1.text = "SUBMITTED BY:\nMr. Deep Koshiya\nB.Sc. AI and Data Science\nSemester: VII"
    cell_m2.text = "SUPERVISED & VERIFIED BY:\nProject Supervisor\nDepartment of AI & Data Science\nEvaluation & Signature Stamp"
    for c in [cell_m1, cell_m2]:
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(9.5)
        set_cell_background(c, "F1F5F9")
        set_cell_margins(c, top=60, bottom=60, left=80, right=80)

    # Signature line
    cell_sig1 = meta_sub_table.rows[1].cells[0]
    cell_sig2 = meta_sub_table.rows[1].cells[1]
    cell_sig1.text = "_________________________\n(Signature of Scholar)"
    cell_sig2.text = "_________________________\n(Signature of Supervisor)"
    for c in [cell_sig1, cell_sig2]:
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(9)
        set_cell_background(c, "FFFFFF")
        set_cell_margins(c, top=50, bottom=50, left=80, right=80)

    doc.add_page_break()

    # =========================================================================
    # PAGE 2: ROMAN PAGES - ACKNOWLEDGMENT & ABSTRACT
    # =========================================================================
    add_para("Page I - II", space_after=2, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h1("ACKNOWLEDGMENT", space_before=2, space_after=4)
    add_para(
        "I express my sincere gratitude to my assigned project supervisor and the esteemed faculty members of the Department of Artificial Intelligence & Data Science for their invaluable guidance, encouragement, and insightful feedback throughout the problem identification and system design phases of this project."
    )
    add_para(
        "I also thank the Project Coordinator and the Head of the School for facilitating the computational lab environment and resources necessary to execute Project-1 (PRJ-701) in accordance with the prescribed university schedule."
    )
    add_para(
        "This First Progress Work report, submitted on 24/09/2026, encompasses the foundational requirements, data engineering pipeline architecture, and complete UML/DFD models (Sections 1.1, 1.2, 1.3, 2.1, and 2.2). I remain dedicated to incorporating all supervisory feedback for the upcoming Mid-Term evaluation."
    )
    add_para("Deep Koshiya | B.Sc. AI and Data Science (Sem VII) | Date: 24th September 2026", space_after=10, space_before=2, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)

    add_h1("ABSTRACT", space_before=6, space_after=4)
    add_para(
        "In modern multi-channel retail, treating an entire customer base as a homogenous entity leads to suboptimal marketing spend, inflated Customer Acquisition Costs (CAC), and customer churn blindspots. The 'E-Commerce Customer Lifecycle & RFM Segmentation Dashboard' resolves these commercial challenges through an end-to-end applied machine learning and analytics engineering pipeline."
    )
    add_para(
        "The automated pipeline ingests raw transaction event logs, enforces schema validation, isolates unauthenticated sessions, segregates reverse logistics (credit notes and cancellations) into an audit ledger, and neutralizes bulk B2B distribution anomalies using Tukey's Interquartile Range (IQR) fences. A vectorized feature engineering engine computes customer-level Recency, Frequency, and Monetary (RFM) dimensions alongside Average Order Value (AOV). Features undergo log1p compression and Z-score standardization to alleviate Pareto spend skewness, enabling robust K-Means clustering evaluated via Elbow and Silhouette diagnostics. Concurrently, a rule-based quantile stratification engine (pd.qcut) computes deterministic 1–5 behavioral tiers mapping customers into 8 actionable cohorts (such as 'Champions', 'Loyal Customers', and 'At Risk Whales')."
    )
    add_para(
        "This First Progress Work report documents the system foundation and software engineering design for the 24/09/2026 milestone. It presents complete Level-0, Level-1, and Level-2 Data Flow Diagrams (DFDs) alongside a comprehensive Use Case Model and specification matrix, fulfilling all requirements for the first 10-mark evaluation."
    )

    doc.add_page_break()

    # =========================================================================
    # PAGE 3: INDEX OF THE PROJECT REPORT & LISTS
    # =========================================================================
    add_para("Page VI - VIII", space_after=2, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h1("INDEX OF THE PROJECT REPORT", space_before=2, space_after=3)
    add_para("(Structured strictly according to Syllabus Guidelines for B.Sc. AI and Data Science SEM VII)", space_after=4, italic=True)

    idx_table = doc.add_table(rows=18, cols=3)
    idx_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    idx_headers = ["SR. NO.", "CONTENTS / TOPIC", "STATUS & PAGE"]
    for j, h in enumerate(idx_headers):
        c = idx_table.rows[0].cells[j]
        c.text = h
        c.paragraphs[0].runs[0].font.name = 'Times New Roman'
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.size = Pt(8.5)
        set_cell_background(c, "0F172A")
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    index_entries = [
        ("*", "Acknowledgment", "Page I"),
        ("*", "Abstract", "Page II"),
        ("*", "Index of the Project Report, List of Figures & Tables", "Page VI - VIII"),
        ("1", "ABOUT THE SYSTEM:", "Page 1"),
        ("1.1", "  Problem definition (Identification of needs)", "[Completed 24/09] Page 1"),
        ("1.2", "  Requirement Specifications (Product/System Tasks)", "[Completed 24/09] Page 2"),
        ("1.3", "  Tools and Technology Used (Front-End/Back-End/Framework)", "[Completed 24/09] Page 3"),
        ("2", "SYSTEM DESIGN USING UML:", "Page 4"),
        ("2.1", "  Data Flow Diagrams (Level 0 Context, Level 1, Level 2 DFDs)", "[Completed 24/09] Page 4"),
        ("2.2", "  Use Case Diagram & Actor Specifications", "[Completed 24/09] Page 7"),
        ("2.3", "  Class Diagram", "[Scheduled 13/10 Mid-Term]"),
        ("2.4", "  Sequence Diagrams", "[Scheduled 13/10 Mid-Term]"),
        ("3", "DATA DICTIONARY: (Database Tables SQL/NoSQL)", "[Scheduled 13/10 Mid-Term]"),
        ("4", "IMPLEMENTATION: (Screen Layouts, Coding, Navigation)", "[Scheduled 13/10 & 27/10]"),
        ("5", "CONCLUSION: (Importance of Worked Tasks, Future Scope)", "[Scheduled 27/10 Third Progress]"),
        ("6", "REFERENCES: (Books, Papers, Technology, Webs)", "[Scheduled 27/10 Third Progress]")
    ]

    for idx, (sr, title, pg) in enumerate(index_entries):
        row = idx_table.rows[idx + 1]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.text, c1.text, c2.text = sr, title, pg
        for cell in [c0, c1, c2]:
            cell.paragraphs[0].runs[0].font.name = 'Times New Roman'
            cell.paragraphs[0].runs[0].font.size = Pt(8)
            set_cell_margins(cell, top=25, bottom=25, left=50, right=50)
            if "[Completed 24/09]" in pg:
                cell.paragraphs[0].runs[0].font.bold = True
                set_cell_background(cell, "EFF6FF")
            else:
                set_cell_background(cell, "FFFFFF" if idx % 2 == 0 else "F8FAFC")

    add_h2("LIST OF FIGURES & TABLES (Progress Work - 1)", space_before=6, space_after=2)
    add_para("Figure 2.1: Level 0 Context DFD (Context Diagram) ................................................................ Page 4")
    add_para("Figure 2.2: Level 1 Data Flow Diagram (Pipeline Architecture) ............................................. Page 5")
    add_para("Figure 2.3: Level 2 Data Flow Diagram (Refinery & ML Decomposition) ................................ Page 6")
    add_para("Figure 2.4: UML Use Case Diagram (System Actors & Capabilities) ....................................... Page 7")
    add_para("Table 1.1: System & Hardware/Software Specifications ........................................................... Page 2")
    add_para("Table 1.2: Technology Stack & Architectural Layer Mapping ..................................................... Page 3")
    add_para("Table 2.1: Formal Use Case Specification Matrix (UC-01 to UC-06) ........................................ Page 7")
    add_para("Table 2.2: Project Progress Work Evaluation & Milestone Tracking .......................................... Page 7")

    doc.add_page_break()

    # =========================================================================
    # PAGE 4: 1. ABOUT THE SYSTEM - 1.1 PROBLEM DEFINITION
    # =========================================================================
    add_para("Page 1", space_after=2, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h1("1. ABOUT THE SYSTEM:")
    add_h2("1.1 Problem definition (Identification of needs):")
    add_para(
        "Modern digital commerce platforms generate continuous transactional event streams across mobile apps, online web checkouts, and physical POS gateways. Relational Online Transaction Processing (OLTP) repositories effectively record line-item orders; however, they inherently track individual checkouts rather than longitudinal customer relationships. Without an analytics engineering framework that transforms low-level orders into customer-centric behavioral dimensions, commercial leadership operates with severe blindspots."
    )

    add_h3("1.1.1 Industry Background & Domain Challenges")
    add_para(
        "E-commerce organizations face escalating paid Customer Acquisition Costs (CAC) across digital ad networks alongside declining organic retention rates. Treating an entire customer portfolio as a monolith leads to severe resource misallocations: high-value repeat shoppers receive unnecessary margin-diluting discount coupons, while early churn-risk signals from historically loyal accounts are overlooked until permanent dormancy occurs."
    )

    add_h3("1.1.2 Identification of Operational & System Needs")
    add_para("The system addresses four primary operational requirements identified in enterprise retail analytics:", bold_prefix="• Core Needs: ")
    add_para("Raw transactional feeds contain anonymous guest checkouts (missing CustomerID), cancelled invoices (prefixed with 'C'), zero-priced administrative adjustments, and extreme B2B bulk orders (e.g., 1,500 units). An automated filter applying Tukey's Interquartile Range (IQR) fences is required to isolate anomalies without compromising retail consumer behavioral patterns.", bold_prefix="1. Automated Data Cleaning & Outlier Isolation: ")
    add_para("The system must condense disparate orders into normalized behavioral dimensions: Recency (days elapsed since last transaction relative to a fixed operational boundary T_obs = max(InvoiceDate) + 1 day), Frequency (distinct checkout visits via nunique), and Monetary spend (net cumulative spend).", bold_prefix="2. Granular Behavioral Metric Derivation (RFM): ")
    add_para("Consumer spend follows a Pareto distribution (extreme positive skew). Euclidean distance-based clustering fails on raw values because monetary variance completely dominates frequency and recency. Log1p compression and Z-score standardization are required prior to K-Means modeling, coupled with a deterministic 1–5 quantile fallback scoring engine.", bold_prefix="3. Skew-Resistant Algorithmic Clustering: ")
    add_para("Executives and marketing teams require an interactive BI interface that visualizes portfolio health, highlights at-risk capital, and exports actionable customer cohorts linked to distinct CRM retention playbooks.", bold_prefix="4. Strategic BI Visualization & Activation: ")

    add_h3("1.1.3 Project Scope & Targeted Objectives")
    add_para(
        "The project delivers a production-grade analytics system bridging raw OLTP logs and strategic decision-making. Deliverables encompass an automated Python refinery pipeline, vectorized RFM engine, unsupervised K-Means clustering, star-schema data marts (FACT_TRANSACTIONS, FACT_REVERSALS, DIM_CUSTOMERS), and an interactive executive web dashboard."
    )

    doc.add_page_break()

    # =========================================================================
    # PAGE 5: 1.2 REQUIREMENT SPECIFICATIONS
    # =========================================================================
    add_para("Page 2", space_after=2, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h2("1.2 Requirement Specifications (Product/System Tasks)")
    add_para(
        "The requirement specifications formally define the operational tasks, behavioral functions, and system performance standards for the E-Commerce Customer Lifecycle & RFM Segmentation system."
    )

    add_h3("1.2.1 Functional Requirements Specifications (FR)")
    add_para("Must ingest raw CSV transactional records, enforce typing (InvoiceNo string, InvoiceDate datetime64[ns], UnitPrice and Quantity float64), and standardize SKU text strings.", bold_prefix="• FR-01: Ingestion & Schema Enforcement: ")
    add_para("Must identify records lacking authenticated CustomerID and drop them to prevent corrupting customer-level longitudinal cohorts.", bold_prefix="• FR-02: Missing Entity Resolution: ")
    add_para("Must segregate credit notes, cancellations (invoices matching '^C'), and non-positive quantities into an audit ledger (fact_reversals_audit.csv) without affecting positive frequency metrics.", bold_prefix="• FR-03: Reverse Logistics Isolation: ")
    add_para("Must compute Tukey's fences (IQR = Q3 - Q1) across Quantity and UnitPrice to remove bulk B2B anomalies.", bold_prefix="• FR-04: Tukey IQR Outlier Capping: ")
    add_para("Must set T_obs = max(InvoiceDate) + 1 day and compute Recency days, Frequency (unique invoices), Monetary spend, and Average Order Value (AOV = Spend / Frequency).", bold_prefix="• FR-05: Vectorized RFM Derivation: ")
    add_para("Must apply log1p skew reduction, StandardScaler standardization, evaluate Elbow Inertia and Silhouette scores across k in [2..6], and fit production K-Means.", bold_prefix="• FR-06: ML Clustering & Diagnostics: ")
    add_para("Must provide fallback 1–5 percentile scoring (pd.qcut) mapping accounts to 8 business segments ('Champions', 'Loyal Customers', 'At Risk Whales', etc.).", bold_prefix="• FR-07: Quantile RFM Scoring: ")
    add_para("Must export clean star-schema tables and render an interactive executive dashboard with KPI scorecards, spatial scatterplots, and a searchable activation ledger.", bold_prefix="• FR-08: Star Schema & Dashboard: ")

    add_h3("1.2.2 Non-Functional Requirements Specifications (NFR)")
    add_para("The pipeline must process 15,000+ transactional records in under 5 seconds on standard x86/ARM hardware.", bold_prefix="• NFR-01: Performance & Latency: ")
    add_para("All feature calculations must maintain linear O(N) complexity using vectorized Pandas/NumPy C-routines.", bold_prefix="• NFR-02: Scalability: ")
    add_para("Fixed random seeds (random_state=42) ensure deterministic clustering across distributed environments.", bold_prefix="• NFR-03: Reproducibility: ")

    add_caption("Table 1.1: System & Hardware/Software Specifications")
    req_table = doc.add_table(rows=6, cols=3)
    req_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    r_hdrs = ["SYSTEM CATEGORY", "MINIMUM REQUIREMENT", "DEPLOYED SPECIFICATION"]
    for j, h in enumerate(r_hdrs):
        c = req_table.rows[0].cells[j]
        c.text = h
        c.paragraphs[0].runs[0].font.name = 'Times New Roman'
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.size = Pt(8.5)
        set_cell_background(c, "0F172A")
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    r_specs = [
        ("Processor / CPU", "Dual-Core x86_64 or ARM64", "Apple Silicon M-Series / Intel Core i5/i7 (8-Core+)"),
        ("System Memory (RAM)", "4 GB DDR3/DDR4", "8 GB to 16 GB Unified Memory"),
        ("Storage Subsystem", "500 MB Free Disk Space", "NVMe Solid State Drive (SSD)"),
        ("Operating System", "macOS 12+, Ubuntu 20.04+, Windows 10+", "macOS Darwin (Unix Environment)"),
        ("Runtime Environment", "Python 3.9+, Modern Web Browser", "Python 3.9.6, Safari, Google Chrome, VS Code")
    ]
    for idx, (cat, min_s, tar_s) in enumerate(r_specs):
        row = req_table.rows[idx + 1]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.text, c1.text, c2.text = cat, min_s, tar_s
        for c in [c0, c1, c2]:
            c.paragraphs[0].runs[0].font.name = 'Times New Roman'
            c.paragraphs[0].runs[0].font.size = Pt(8)
            set_cell_margins(c, top=35, bottom=35, left=70, right=70)
            set_cell_background(c, "F8FAFC" if idx % 2 == 0 else "FFFFFF")

    doc.add_page_break()

    # =========================================================================
    # PAGE 6: 1.3 TOOLS AND TECHNOLOGY USED
    # =========================================================================
    add_para("Page 3", space_after=2, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h2("1.3 Tools and Technology Used (Front-End/Back-End/Framework/Protocol/API/Services/App. etc.)")
    add_para(
        "The project architecture integrates an enterprise-ready analytical engineering stack. Each technology component has been chosen to maximize processing throughput, algorithmic precision, and visual interactivity."
    )

    add_h3("1.3.1 Core Processing & Numerical Computing")
    add_para("Primary runtime environment offering native execution of optimized C-extensions for data science.", bold_prefix="• Python 3.9+: ")
    add_para("High-performance data structures executing vectorized group aggregations, type coercion, datetime logic, and quantile ranking (pd.qcut).", bold_prefix="• Pandas (v2.3.3): ")
    add_para("Vectorized numerical array operations, logarithmic transformations (log1p), and Tukey fence computations.", bold_prefix="• NumPy (v2.0.2): ")

    add_h3("1.3.2 Machine Learning & Statistical Modeling")
    add_para("Implements StandardScaler for Z-score normalization, KMeans with k-means++ initialization, and silhouette_score for cluster quality validation.", bold_prefix="• Scikit-Learn (v1.6.1): ")
    add_para("Provides statistical functions supporting spatial distance evaluations and distribution analysis.", bold_prefix="• SciPy (v1.13.1): ")

    add_h3("1.3.3 Presentation, Dashboards & Storage")
    add_para("Utility-first styling framework providing responsive dark-mode executive UI layouts.", bold_prefix="• HTML5 & Tailwind CSS: ")
    add_para("Interactive charting engine rendering segment revenue donuts, spatial scatterplots, and diagnostic curves.", bold_prefix="• Chart.js: ")
    add_para("Data marts formatted as a star schema (fact table linked to customer dimension) for native Tableau/Power BI integration.", bold_prefix="• Star-Schema BI Marts: ")
    add_para("Git for version control, GitHub for cloud repository hosting, and python-docx for automated report generation.", bold_prefix="• Engineering Tooling: ")

    add_caption("Table 1.2: Technology Stack & Architectural Layer Mapping")
    tech_table = doc.add_table(rows=6, cols=3)
    tech_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_hdrs = ["SYSTEM LAYER", "TECHNOLOGY / FRAMEWORK", "PRIMARY OPERATIONAL TASK"]
    for j, h in enumerate(t_hdrs):
        c = tech_table.rows[0].cells[j]
        c.text = h
        c.paragraphs[0].runs[0].font.name = 'Times New Roman'
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.size = Pt(8.5)
        set_cell_background(c, "0F172A")
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    t_data = [
        ("Data Refinery Pipeline", "Python, Pandas, NumPy", "Schema typing, null dropping, reversal isolation, IQR outlier capping"),
        ("Feature Engineering", "Pandas GroupBy, Datetime", "Vectorized derivation of Recency, Frequency, Monetary, and AOV"),
        ("Machine Learning", "Scikit-Learn, SciPy", "Log1p transformation, StandardScaler, K-Means clustering, Silhouette score"),
        ("Data Mart Layer", "CSV / Parquet Data Marts", "Star-schema fact & dimension tables for BI compatibility"),
        ("Executive BI Dashboard", "HTML5, Tailwind CSS, Chart.js", "Interactive KPI cards, spatial bubble plot, searchable activation grid")
    ]
    for idx, (lay, tool, func) in enumerate(t_data):
        row = tech_table.rows[idx + 1]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.text, c1.text, c2.text = lay, tool, func
        for c in [c0, c1, c2]:
            c.paragraphs[0].runs[0].font.name = 'Times New Roman'
            c.paragraphs[0].runs[0].font.size = Pt(8)
            set_cell_margins(c, top=35, bottom=35, left=70, right=70)
            set_cell_background(c, "F8FAFC" if idx % 2 == 0 else "FFFFFF")

    doc.add_page_break()

    # =========================================================================
    # PAGE 7: 2. SYSTEM DESIGN USING UML - 2.1 LEVEL 0 CONTEXT DFD
    # =========================================================================
    add_para("Page 4", space_after=2, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h1("2. SYSTEM DESIGN USING UML:")
    add_h2("2.1 Data Flow Diagrams (1 Level DFD, 2 Level DFD must).")
    add_para(
        "Data Flow Diagrams (DFDs) provide a graphical representation of information flow across the system, modeling processes, external entities, data stores, and directional data pathways following Gane & Sarson conventions:"
    )
    add_para("External Entities (Rectangles), Processes (Rounded Rectangles/Circles), Data Stores (Open Horizontal Lines), and Data Flows (Directed Arrows).", bold_prefix="• Notation Standards: ")

    add_h3("2.1.1 Level 0 Data Flow Diagram (Context-Level DFD)")
    add_para(
        "The Level 0 Context DFD defines the global operational boundary of Process 0.0 ('E-Commerce Customer Lifecycle & RFM Analytics System') interacting with four external entities: (1) OLTP Transaction Stream supplying sales logs; (2) Data Scientist supplying IQR tuning parameters; (3) CRM Marketing Engine consuming targeted segments; and (4) Commercial Executive team receiving portfolio KPI scorecards."
    )

    df0_path = img_dir / "fig_dfd_level_0.png"
    if df0_path.exists():
        doc.add_picture(str(df0_path), width=Inches(5.4))
        add_caption("Figure 2.1: Level 0 Data Flow Diagram (Context-Level DFD)")

    doc.add_page_break()

    # =========================================================================
    # PAGE 8: 2.1.2 LEVEL 1 DATA FLOW DIAGRAM
    # =========================================================================
    add_para("Page 5", space_after=2, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h3("2.1.2 Level 1 Data Flow Diagram (End-to-End Pipeline Architecture)")
    add_para(
        "The Level 1 DFD decomposes the system into five primary sequential operational processes and identifies the underlying persistent data stores:"
    )
    add_para("Ingests raw orders, validates data types, and logs records to D1 (Raw Transaction Log).", bold_prefix="• Process 1.0 (Ingestion & Schema Typing): ")
    add_para("Purges missing CustomerIDs, routes credit notes and cancellations to D2 (Reversals Audit Ledger), and applies Tukey IQR fences on Quantity and Price.", bold_prefix="• Process 2.0 (Data Quality & Refinery Pipeline): ")
    add_para("Computes snapshot boundary T_obs, aggregates unique invoices for Frequency, calculates Recency days, and sums net spend, writing to D3 (RFM Analytical Base Table).", bold_prefix="• Process 3.0 (Vectorized RFM Feature Engineering): ")
    add_para("Applies log1p transformation and StandardScaler, fits K-Means, and assigns quantile 1–5 scoring rules, writing to D4 (Dimensional Customer Mart).", bold_prefix="• Process 4.0 (ML Clustering & Quantile Scoring): ")
    add_para("Builds the star schema and renders the interactive executive web dashboard.", bold_prefix="• Process 5.0 (Star-Schema & Dashboard Export): ")

    df1_path = img_dir / "fig_dfd_level_1.png"
    if df1_path.exists():
        doc.add_picture(str(df1_path), width=Inches(5.4))
        add_caption("Figure 2.2: Level 1 Data Flow Diagram (End-to-End Pipeline Architecture)")

    doc.add_page_break()

    # =========================================================================
    # PAGE 9: 2.1.3 LEVEL 2 DATA FLOW DIAGRAM (DECOMPOSITION)
    # =========================================================================
    add_para("Page 6", space_after=2, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h3("2.1.3 Level 2 Data Flow Diagram (Sub-Process Decomposition)")
    add_para(
        "To provide rigorous engineering granularity as mandated by syllabus guidelines, the Level 2 DFD decomposes the two critical computational processes: Process 2.0 (Data Refinery Pipeline) and Process 4.0 (Algorithmic & Quantile Engine)."
    )
    add_para("Sub-process 2.1 checks CustomerID presence, dropping guest sessions. Sub-process 2.2 identifies cancellation identifiers ('^C' or Quantity <= 0), segregating them to an audit ledger. Sub-process 2.3 computes Tukey fences across positive items, eliminating wholesale distributor anomalies.", bold_prefix="• Sub-Process 2.0 Decomposition (Refinery): ")
    add_para("Sub-process 4.1 performs log1p compression to alleviate Pareto right-skewness. Sub-process 4.2 executes Z-score standardization. Sub-process 4.3 runs K-Means clustering across k in [2..6] with Silhouette scoring. Sub-process 4.4 runs quantile binning (pd.qcut) to generate 1–5 scores and map 8 business operating tiers.", bold_prefix="• Sub-Process 4.0 Decomposition (ML & Scoring): ")

    df2_path = img_dir / "fig_dfd_level_2.png"
    if df2_path.exists():
        doc.add_picture(str(df2_path), width=Inches(5.4))
        add_caption("Figure 2.3: Level 2 Data Flow Diagram (Refinery & ML Decomposition)")

    doc.add_page_break()

    # =========================================================================
    # PAGE 10: 2.2 USE CASE DIAGRAM & MILESTONE SPECIFICATION
    # =========================================================================
    add_para("Page 7", space_after=2, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h2("2.2 Use Case Diagram & Milestone Tracking.")
    add_para(
        "The Use Case Model delineates the functional boundaries of the system by illustrating interactions between primary human and automated actors and the system's operational capabilities."
    )

    uc_path = img_dir / "fig_use_case_diagram.png"
    if uc_path.exists():
        doc.add_picture(str(uc_path), width=Inches(3.7))
        add_caption("Figure 2.4: UML Use Case Diagram (System Actors & Capabilities)")

    add_caption("Table 2.1: Formal Use Case Specification Matrix (UC-01 to UC-06)")
    uc_table = doc.add_table(rows=7, cols=4)
    uc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    u_hdrs = ["UC ID", "USE CASE NAME", "PRIMARY ACTOR", "OPERATIONAL GOAL & PRE-CONDITION"]
    for j, h in enumerate(u_hdrs):
        c = uc_table.rows[0].cells[j]
        c.text = h
        c.paragraphs[0].runs[0].font.name = 'Times New Roman'
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.size = Pt(8)
        set_cell_background(c, "0F172A")
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    u_data = [
        ("UC-01", "Ingest Transaction Stream", "Data Engineer / Admin", "Ingests raw event logs, validates schema, enforces datatypes."),
        ("UC-02", "Isolate Reversals & IQR Filter", "Data Engineer / Admin", "Purges null IDs, isolates returns (^C), caps Tukey bulk outliers."),
        ("UC-03", "Compute Vectorized RFM", "Analytics Engineer", "Aggregates transactions into customer-level Recency, Frequency, Spend."),
        ("UC-04", "Train & Tune K-Means Model", "Data Scientist", "Applies log1p + StandardScaler, evaluates Elbow/Silhouette, fits k=4."),
        ("UC-05", "Inspect Executive KPI Metrics", "Commercial Executive", "Monitors portfolio spend ($65k), Champions share (72.5%), at-risk capital."),
        ("UC-06", "Export Targeted CRM Audiences", "Marketing Manager", "Filters segments in activation ledger and exports audience lists for campaigns.")
    ]
    for idx, (uid, uname, uact, ugoal) in enumerate(u_data):
        row = uc_table.rows[idx + 1]
        c0, c1, c2, c3 = row.cells[0], row.cells[1], row.cells[2], row.cells[3]
        c0.text, c1.text, c2.text, c3.text = uid, uname, uact, ugoal
        for c in [c0, c1, c2, c3]:
            c.paragraphs[0].runs[0].font.name = 'Times New Roman'
            c.paragraphs[0].runs[0].font.size = Pt(7.5)
            set_cell_margins(c, top=20, bottom=20, left=40, right=40)
            set_cell_background(c, "F8FAFC" if idx % 2 == 0 else "FFFFFF")

    add_h2("MILESTONE EVALUATION SUMMARY (24/09/2026)", space_before=5, space_after=2)
    add_para(
        "This submission formally completes the First Progress Work milestone (10 Marks allotted) scheduled for 24/09/2026 for B.Sc. AI and Data Science. All required index sections (1.1, 1.2, 1.3, 2.1, and 2.2) have been fully developed and validated. The technical artifacts, UML models, and data pipelines are complete and ready for supervisor review."
    )

    add_caption("Table 2.2: Project Progress Work Evaluation & Milestone Tracking")
    prog_table = doc.add_table(rows=6, cols=5)
    prog_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    p_hdrs = ["DATE", "ASSESSMENT PARTICULARS", "INDEX SECTIONS", "MARKS", "CURRENT STATUS"]
    for j, h in enumerate(p_hdrs):
        c = prog_table.rows[0].cells[j]
        c.text = h
        c.paragraphs[0].runs[0].font.name = 'Times New Roman'
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.size = Pt(7.5)
        set_cell_background(c, "0F172A")
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    p_data = [
        ("24/09/2026", "First Progress Work", "1.1, 1.2, 1.3, 2.1, 2.2", "10 Marks", "COMPLETED (Current Submission)"),
        ("13/10/2026", "Mid-Term Exam (Second Progress)", "2.3, 2.4, 3.1, 4.1", "20 Marks", "Scheduled for Next Milestone"),
        ("27/10/2026", "Third Progress Work", "4.2, 4.3, 5.1, 5.2, 6.1", "10 Marks", "Scheduled for Third Milestone"),
        ("06/11/2026", "Report Submission (Spiral Binding)", "Final Signed Spiral Copy", "10 Marks", "Scheduled for Final Submission"),
        ("18-20/11/2026", "Final Exam (SEE)", "Exam & Viva Voce", "50 Marks", "Scheduled for Semester Examination")
    ]
    for idx, (dt, part, sec, mrk, stat) in enumerate(p_data):
        row = prog_table.rows[idx + 1]
        c0, c1, c2, c3, c4 = row.cells[0], row.cells[1], row.cells[2], row.cells[3], row.cells[4]
        c0.text, c1.text, c2.text, c3.text, c4.text = dt, part, sec, mrk, stat
        for c in [c0, c1, c2, c3, c4]:
            c.paragraphs[0].runs[0].font.name = 'Times New Roman'
            c.paragraphs[0].runs[0].font.size = Pt(7)
            set_cell_margins(c, top=20, bottom=20, left=40, right=40)
            if idx == 0:
                c.paragraphs[0].runs[0].font.bold = True
                set_cell_background(c, "DCFCE7")
            else:
                set_cell_background(c, "F8FAFC" if idx % 2 == 0 else "FFFFFF")

    out_docx_path = base_dir / "Deep_Koshiya_ProjectReport_Progress1.docx"
    doc.save(str(out_docx_path))
    doc.save(str(base_dir / "Deep_Koshiya_ProjectReport.docx"))
    print(f"[SUCCESS] Saved beautifully styled report to {out_docx_path}")

if __name__ == "__main__":
    build_report()
