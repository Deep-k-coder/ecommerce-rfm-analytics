"""
build_detailed_10page_report.py
Generates the comprehensive ~10-page First Progress Work Project Report (Date: 24/09/2026)
Course: Project -1 (CE-502, PRJ-701) | B. Tech CEs & IT SEM VII
Candidate: Deep Koshiya
Conforming strictly to the provided syllabus schedule and report index:
- Sections Covered: Preliminary Roman Pages + 1.1, 1.2, 1.3, 2.1, 2.2
- Formatting: Times New Roman, 16pt main title, 14pt subtitle, 12pt regular text, justified paragraphs.
"""

from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, hex_color):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def create_report():
    base_dir = Path(__file__).resolve().parent
    img_dir = base_dir / "report_images"
    doc = Document()

    # Configure Margins: 1.25 inch Left (for spiral binding as noted in syllabus), 1.0 inch Top, Bottom, Right
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.25)
        section.right_margin = Inches(1.0)

    # Base Colors
    NAVY = RGBColor(15, 23, 42)      # Main text
    INDIGO = RGBColor(49, 46, 129)   # Main titles (Dark Navy/Indigo)
    CHARCOAL = RGBColor(30, 41, 59)  # Subtitles

    # Global Style Configuration
    style_normal = doc.styles['Normal']
    font_normal = style_normal.font
    font_normal.name = 'Times New Roman'
    font_normal.size = Pt(12)
    font_normal.color.rgb = NAVY

    # Helper Functions
    def add_para(text, bold_prefix=None, space_after=6, space_before=0, line_spacing=1.15, italic=False):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.font.name = 'Times New Roman'
            r_pre.font.size = Pt(12)
            r_pre.font.bold = True
            r_pre.font.color.rgb = NAVY
            
        r_text = p.add_run(text)
        r_text.font.name = 'Times New Roman'
        r_text.font.size = Pt(12)
        r_text.font.italic = italic
        r_text.font.color.rgb = NAVY
        return p

    def add_h1(text, space_before=14, space_after=6):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(16)
        r.font.bold = True
        r.font.color.rgb = INDIGO
        return p

    def add_h2(text, space_before=12, space_after=4):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = CHARCOAL
        return p

    def add_h3(text, space_before=10, space_after=3):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12.5)
        r.font.bold = True
        r.font.color.rgb = NAVY
        return p

    def add_caption(text, is_table=False):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4 if not is_table else 8)
        p.paragraph_format.space_after = Pt(8 if not is_table else 4)
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.color.rgb = CHARCOAL
        return p

    # =========================================================================
    # PAGE 1: FORMAL COVER PAGE & PROGRESS WORK - 1 CERTIFICATE
    # =========================================================================
    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_after = Pt(2)
    r_inst = p_inst.add_run("DEPARTMENT OF COMPUTER ENGINEERING & INFORMATION TECHNOLOGY")
    r_inst.font.name = 'Times New Roman'
    r_inst.font.size = Pt(13)
    r_inst.font.bold = True
    r_inst.font.color.rgb = NAVY

    p_deg = doc.add_paragraph()
    p_deg.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_deg.paragraph_format.space_after = Pt(28)
    r_deg = p_deg.add_run("B. Tech (CEs & IT) - Semester VII | Academic Year 2026-2027")
    r_deg.font.name = 'Times New Roman'
    r_deg.font.size = Pt(11)
    r_deg.font.color.rgb = CHARCOAL

    p_proj_title = doc.add_paragraph()
    p_proj_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_proj_title.paragraph_format.space_after = Pt(8)
    r_proj = p_proj_title.add_run("E-COMMERCE CUSTOMER LIFECYCLE &\nRFM SEGMENTATION DASHBOARD")
    r_proj.font.name = 'Times New Roman'
    r_proj.font.size = Pt(18)
    r_proj.font.bold = True
    r_proj.font.color.rgb = INDIGO

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(24)
    r_sub = p_sub.add_run("A Project Report Submitted for the Milestone Assessment of:\nPROJECT - 1 (CE-502, PRJ-701)")
    r_sub.font.name = 'Times New Roman'
    r_sub.font.size = Pt(13)
    r_sub.font.bold = True
    r_sub.font.color.rgb = NAVY

    # Milestone Box Table
    m_table = doc.add_table(rows=6, cols=2)
    m_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    m_rows = [
        ("Assessment Particular:", "FIRST PROGRESS WORK EVALUATION"),
        ("Scheduled Submission Date:", "24/09/2026 (Time: 10:00 AM to 5:00 PM)"),
        ("Work Completed (as per Index):", "Chapters 1.1, 1.2, 1.3, 2.1, and 2.2"),
        ("Evaluation Weightage:", "10 Marks (First Progress Work)"),
        ("Candidate Name & Scholar:", "Deep Koshiya"),
        ("Program & Semester:", "B. Tech Computer Engineering & IT, Semester VII")
    ]
    for idx, (label, val) in enumerate(m_rows):
        r = m_table.rows[idx]
        c0, c1 = r.cells[0], r.cells[1]
        c0.text = label
        c1.text = val
        c0.paragraphs[0].runs[0].font.name = 'Times New Roman'
        c0.paragraphs[0].runs[0].font.bold = True
        c0.paragraphs[0].runs[0].font.size = Pt(10.5)
        c1.paragraphs[0].runs[0].font.name = 'Times New Roman'
        c1.paragraphs[0].runs[0].font.size = Pt(10.5)
        set_cell_background(c0, "F1F5F9")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, top=100, bottom=100, left=140, right=140)
        set_cell_margins(c1, top=100, bottom=100, left=140, right=140)

    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_after = Pt(36)

    # Signature Block Table
    sig_table = doc.add_table(rows=2, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_s1 = sig_table.rows[0].cells[0]
    cell_s2 = sig_table.rows[0].cells[1]
    cell_s1.text = "______________________________\nMr. Deep Koshiya\n(Project Scholar)"
    cell_s2.text = "______________________________\nAssigned Project Supervisor\n(Signature & Verification Stamp)"
    for row in sig_table.rows:
        for c in row.cells:
            for p in c.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(10)
            set_cell_background(c, "FFFFFF")

    doc.add_page_break()

    # =========================================================================
    # PAGE 2: ROMAN PAGES - ACKNOWLEDGMENT & ABSTRACT
    # =========================================================================
    p_rom_hdr1 = doc.add_paragraph()
    p_rom_hdr1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_rom1 = p_rom_hdr1.add_run("Page I")
    r_rom1.font.name = 'Times New Roman'
    r_rom1.font.size = Pt(10)
    r_rom1.font.italic = True

    add_h1("ACKNOWLEDGMENT", space_before=10, space_after=12)
    add_para(
        "I would like to express my deepest gratitude to my assigned project supervisor and the respected faculty members of the Department of Computer Engineering & Information Technology for their invaluable guidance, consistent encouragement, and critical feedback throughout the conceptualization, system analysis, and design phases of this project."
    )
    add_para(
        "I am particularly indebted to the Project Coordinator and the Head of the School for facilitating the required computing infrastructure, laboratory environments, and administrative support necessary to execute Project-1 (CE-502, PRJ-701) in accordance with the prescribed university schedule."
    )
    add_para(
        "This First Progress Work report, submitted for the formal assessment milestone on 24/09/2026, encompasses the core problem formulation, requirements specification, technological architecture, and complete UML/DFD structural designs (Sections 1.1, 1.2, 1.3, 2.1, and 2.2). I remain dedicated to incorporating all supervisory recommendations as the project transitions into database physical modeling and algorithmic implementation for the forthcoming Mid-Term milestone."
    )

    p_sig_ack = doc.add_paragraph()
    p_sig_ack.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig_ack.paragraph_format.space_before = Pt(24)
    r_ack_sig = p_sig_ack.add_run("Deep Koshiya\nB. Tech CE & IT (Sem VII)\nDate: 24th September 2026")
    r_ack_sig.font.name = 'Times New Roman'
    r_ack_sig.font.size = Pt(11)
    r_ack_sig.font.bold = True

    p_spacer2 = doc.add_paragraph()
    p_spacer2.paragraph_format.space_after = Pt(20)

    p_rom_hdr2 = doc.add_paragraph()
    p_rom_hdr2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_rom2 = p_rom_hdr2.add_run("Page II")
    r_rom2.font.name = 'Times New Roman'
    r_rom2.font.size = Pt(10)
    r_rom2.font.italic = True

    add_h1("ABSTRACT", space_before=10, space_after=12)
    add_para(
        "In the contemporary landscape of digital retail and global multi-channel e-commerce, enterprises encounter severe inefficiencies by treating their customer base as a homogenous entity. Generic marketing blasts, unfocused promotional discounts, and late-stage reactive churn interventions lead to high Customer Acquisition Costs (CAC) and progressive revenue attrition. This project, titled 'E-Commerce Customer Lifecycle & RFM Segmentation Dashboard', delivers an end-to-end data analytics and machine learning system engineered to resolve these commercial blindspots."
    )
    add_para(
        "The system architecture ingests raw, unindexed transaction logs and feeds them into an automated data refinery pipeline. This pipeline executes schema validation, isolates unauthenticated sessions, segregates reverse logistics (credit notes and cancellations) into an audit ledger, and neutralizes bulk B2B distribution anomalies through Tukey's Interquartile Range (IQR) method. A vectorized feature engineering engine aggregates the cleansed stream into customer-level Recency, Frequency, and Monetary (RFM) dimensions. To account for heavy right-skewed spend distributions, features undergo log-transformation and standard Z-score scaling before being modeled by an unsupervised K-Means algorithm evaluated via Elbow and Silhouette diagnostics. Concurrently, a rule-based quantile stratification engine (pd.qcut) computes deterministic 1–5 behavioral tiers mapping customers into 8 actionable cohorts (such as 'Champions', 'Loyal Customers', and 'At Risk Whales')."
    )
    add_para(
        "This First Progress Work report (dated 24/09/2026) fully documents the system foundation, operational requirements, technical stack, and software engineering design using UML. It presents complete Level-0, Level-1, and Level-2 Data Flow Diagrams (DFDs) alongside a comprehensive Use Case Model and specification matrix, fulfilling all requirements stipulated for the first 10-mark evaluation milestone."
    )

    doc.add_page_break()

    # =========================================================================
    # PAGE 3: PROJECT REPORT INDEX & LISTS OF FIGURES/TABLES
    # =========================================================================
    p_rom_hdr3 = doc.add_paragraph()
    p_rom_hdr3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_rom3 = p_rom_hdr3.add_run("Page VI - VIII")
    r_rom3.font.name = 'Times New Roman'
    r_rom3.font.size = Pt(10)
    r_rom3.font.italic = True

    add_h1("INDEX OF THE PROJECT REPORT", space_before=8, space_after=10)
    p_idx_note = doc.add_paragraph()
    p_idx_note.paragraph_format.space_after = Pt(8)
    r_idx_note = p_idx_note.add_run("(Structured strictly according to Syllabus Guidelines for B. Tech CEs & IT SEM VII)")
    r_idx_note.font.name = 'Times New Roman'
    r_idx_note.font.size = Pt(10)
    r_idx_note.font.italic = True

    # Master Index Table
    idx_table = doc.add_table(rows=19, cols=3)
    idx_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    idx_headers = ["SR. NO.", "CONTENTS / TOPIC", "STATUS & MILESTONE PAGE"]
    for j, h in enumerate(idx_headers):
        c = idx_table.rows[0].cells[j]
        c.text = h
        c.paragraphs[0].runs[0].font.name = 'Times New Roman'
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.size = Pt(9.5)
        set_cell_background(c, "0F172A")
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    index_entries = [
        ("*", "Acknowledgment", "Page I"),
        ("*", "Abstract", "Page II"),
        ("*", "Index of the Project Report", "Page VI"),
        ("*", "List of Figures & List of Tables", "Page VII - VIII"),
        ("1", "ABOUT THE SYSTEM:", "Page 1"),
        ("1.1", "  Problem definition (Identification of needs)", "[Completed 24/09] Page 1"),
        ("1.2", "  Requirement Specifications (Product/System Tasks)", "[Completed 24/09] Page 2"),
        ("1.3", "  Tools and Technology Used (Front-End/Back-End/Framework)", "[Completed 24/09] Page 4"),
        ("2", "SYSTEM DESIGN USING UML:", "Page 6"),
        ("2.1", "  Data Flow Diagrams (Level 0, Level 1, Level 2 DFDs)", "[Completed 24/09] Page 6"),
        ("2.2", "  Use Case Diagram & Actor Specifications", "[Completed 24/09] Page 9"),
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
        c0.text = sr
        c1.text = title
        c2.text = pg
        for cell in [c0, c1, c2]:
            cell.paragraphs[0].runs[0].font.name = 'Times New Roman'
            cell.paragraphs[0].runs[0].font.size = Pt(9)
            if "[Completed 24/09]" in pg:
                cell.paragraphs[0].runs[0].font.bold = True
                set_cell_background(cell, "EFF6FF")
            else:
                set_cell_background(cell, "FFFFFF" if idx % 2 == 0 else "F8FAFC")

    # List of Figures & Tables
    add_h2("LIST OF FIGURES & LIST OF TABLES (Progress Work - 1)", space_before=14, space_after=6)
    add_para("Figure 2.1: Level 0 Data Flow Diagram (Context-Level DFD) ..................................................... Page 7", bold_prefix=None)
    add_para("Figure 2.2: Level 1 Data Flow Diagram (End-to-End Pipeline Architecture) .............................. Page 8", bold_prefix=None)
    add_para("Figure 2.3: Level 2 Data Flow Diagram (Refinery & ML Decomposition) ................................. Page 9", bold_prefix=None)
    add_para("Figure 2.4: UML Use Case Diagram (System Actors & Functional Capabilities) ......................... Page 10", bold_prefix=None)
    add_para("Table 1.1: Comprehensive System & Hardware/Software Specifications .................................... Page 3", bold_prefix=None)
    add_para("Table 1.2: Technology Stack & Architectural Layer Mapping ...................................................... Page 5", bold_prefix=None)
    add_para("Table 2.1: Formal Use Case Specification Matrix (UC-01 to UC-06) ......................................... Page 10", bold_prefix=None)
    add_para("Table 2.2: Project Progress Work Evaluation & Milestone Tracking ........................................... Page 10", bold_prefix=None)

    doc.add_page_break()

    # =========================================================================
    # PAGE 4: 1. ABOUT THE SYSTEM - 1.1 PROBLEM DEFINITION
    # =========================================================================
    p_body_hdr1 = doc.add_paragraph()
    p_body_hdr1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_bhdr1 = p_body_hdr1.add_run("Page 1")
    r_bhdr1.font.name = 'Times New Roman'
    r_bhdr1.font.size = Pt(10)
    r_bhdr1.font.italic = True

    add_h1("1. ABOUT THE SYSTEM:")
    add_h2("1.1 Problem definition (Identification of needs):")
    
    add_para(
        "Modern electronic commerce systems generate massive streams of transactional data across web portals, mobile applications, and point-of-sale systems. In standard commercial operations, relational Online Transaction Processing (OLTP) databases record millions of sales rows. However, a significant operational challenge arises: raw transaction logs record orders, not customer relationships. Without an analytics engineering framework that transforms line-item records into customer-level behavioral metrics, commercial leadership operates under severe revenue blindspots."
    )
    
    add_h3("1.1.1 Industry Background & Commercial Challenges")
    add_para(
        "Commercial enterprises face intense competitive pressures, escalating paid acquisition costs (CAC on platforms like Google and Meta), and diminishing organic customer retention. In standard practice, marketing departments execute promotional campaigns uniformly across the entire customer ledger. High-spending repeat buyers receive unnecessary blanket discounts that erode gross margins, while dormant accounts are ignored until permanent churn occurs."
    )

    add_h3("1.1.2 Identification of Operational Needs")
    add_para(
        "To achieve commercial efficiency and capital preservation, the system identifies four critical operational needs:",
        bold_prefix="• Core Analytical Needs: "
    )
    add_para(
        "Raw transactional datasets are plagued by missing customer identifiers (guest checkouts), credit note cancellations (invoices prefixed with 'C'), zero-priced administrative entries, and extreme wholesale bulk anomalies. The system requires an automated statistical filter (Tukey's IQR fences) to isolate noise without manual intervention.",
        bold_prefix="1. Automated Data Refinery & Cleansing: "
    )
    add_para(
        "A mathematical aggregation engine is needed to condense millions of disparate transactions into three normalized customer behavioral dimensions: Recency (days since last purchase relative to a fixed snapshot T_obs), Frequency (count of distinct checkout visits), and Monetary Value (net cumulative spend).",
        bold_prefix="2. Granular Behavioral Metric Derivation (RFM): "
    )
    add_para(
        "E-commerce spend exhibits extreme right-skew (Pareto distribution). Standard clustering algorithms relying on Euclidean distance fail when applied to raw metrics because monetary variance overpowers frequency. The system requires log-transformation, Z-score standardization, and algorithmic K-Means clustering paired with deterministic 1–5 quantile scoring.",
        bold_prefix="3. Skew-Resistant Algorithmic Clustering: "
    )
    add_para(
        "Executives and CRM campaign specialists need an interactive BI interface that visualizes revenue concentration, highlights at-risk capital, and exports actionable customer lists linked to concrete marketing playbooks.",
        bold_prefix="4. Executive Visualization & CRM Activation: "
    )

    add_h3("1.1.3 Project Scope & Targeted Objectives")
    add_para(
        "The primary objective of this project is to build an end-to-end, production-grade analytics engineering system that bridges raw OLTP transaction data and strategic executive decision-making. The project deliverables include an automated Python refinery pipeline, vectorized RFM feature store, unsupervised machine learning clustering engine, dimensional star-schema data marts (FACT_TRANSACTIONS, FACT_REVERSALS, DIM_CUSTOMERS), and an interactive executive web dashboard."
    )

    doc.add_page_break()

    # =========================================================================
    # PAGE 5: 1.2 REQUIREMENT SPECIFICATIONS
    # =========================================================================
    p_body_hdr2 = doc.add_paragraph()
    p_body_hdr2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_bhdr2 = p_body_hdr2.add_run("Page 2 - 3")
    r_bhdr2.font.name = 'Times New Roman'
    r_bhdr2.font.size = Pt(10)
    r_bhdr2.font.italic = True

    add_h2("1.2 Requirement Specifications (Product/System Tasks)")
    add_para(
        "The requirement specifications define the operational capabilities, behavioural tasks, and performance boundaries that the E-Commerce Customer Lifecycle & RFM Segmentation system must satisfy. These requirements are classified into Functional and Non-Functional specifications."
    )

    add_h3("1.2.1 Functional Requirements Specifications (FR)")
    add_para(
        "The system must load raw CSV/OLTP transactional records, validate schema integrity, enforce typing (casting InvoiceNo to string, InvoiceDate to datetime64[ns], UnitPrice and Quantity to float64), and handle character encoding variances.",
        bold_prefix="• FR-01: Ingestion & Schema Enforcement: "
    )
    add_para(
        "The system must detect records lacking CustomerID (guest/unauthenticated checkouts) and drop them from customer-level longitudinal cohorts to prevent data contamination.",
        bold_prefix="• FR-02: Missing Entity Identification: "
    )
    add_para(
        "The system must detect cancellations and returns (invoices starting with 'C', non-positive quantities, or zero unit prices) and isolate them into a separate audit ledger (fact_reversals_audit.csv) without distorting purchase frequencies.",
        bold_prefix="• FR-03: Reverse Logistics Segregation: "
    )
    add_para(
        "The system must calculate Tukey's Interquartile Range (IQR = Q3 - Q1) on Quantity and UnitPrice, filtering extreme B2B bulk anomalies to bound the data to retail consumer behavior.",
        bold_prefix="• FR-04: Statistical Outlier Normalization: "
    )
    add_para(
        "The system must establish a fixed operational boundary T_obs = max(InvoiceDate) + 1 day and compute Recency (days elapsed), Frequency (count of unique invoices via nunique), Monetary spend (sum of Quantity * UnitPrice), and Average Order Value (AOV = Spend / Frequency).",
        bold_prefix="• FR-05: Vectorized RFM Feature Derivation: "
    )
    add_para(
        "The system must apply log1p transformation to reduce Pareto skewness, standardize features via StandardScaler (mu=0, sigma=1), evaluate cluster quality across k in [2..6] via Elbow Inertia and Silhouette scores, and train K-Means clustering.",
        bold_prefix="• FR-06: Algorithmic Clustering & Diagnostics: "
    )
    add_para(
        "The system must provide a deterministic fallback scoring engine utilizing percentile ranking (pd.qcut) to assign 1–5 scores for R, F, and M, mapping accounts to 8 operational business segments ('Champions', 'Loyal Customers', 'At Risk Whales', etc.).",
        bold_prefix="• FR-07: Quantile Scoring & Segment Mapping: "
    )
    add_para(
        "The system must export clean star-schema datasets and provide an interactive executive dashboard displaying KPI cards, segment revenue donuts, spatial scatterplots, and a searchable customer activation ledger.",
        bold_prefix="• FR-08: Data Mart Export & Dashboarding: "
    )

    add_h3("1.2.2 Non-Functional Requirements Specifications (NFR)")
    add_para("The end-to-end pipeline must process over 15,000 raw transactional records in under 5 seconds on standard x86/ARM hardware.", bold_prefix="• NFR-01: Performance & Latency: ")
    add_para("The feature extraction and aggregation algorithms must be vectorized via NumPy/Pandas C-extensions, maintaining O(N) linear time complexity.", bold_prefix="• NFR-02: Scalability: ")
    add_para("Random seeds must be fixed (random_state=42) to ensure identical cluster centroids across distributed training runs.", bold_prefix="• NFR-03: Reproducibility & Determinism: ")
    add_para("Separating reversals into an audit trail ensures full compliance with financial accounting standards.", bold_prefix="• NFR-04: Data Traceability & Auditability: ")

    add_caption("Table 1.1: Comprehensive System & Hardware/Software Requirements Specification", is_table=True)
    req_table = doc.add_table(rows=6, cols=3)
    req_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    r_hdrs = ["CATEGORY", "MINIMUM REQUIREMENT", "TARGET / DEPLOYED SPECIFICATION"]
    for j, h in enumerate(r_hdrs):
        c = req_table.rows[0].cells[j]
        c.text = h
        c.paragraphs[0].runs[0].font.name = 'Times New Roman'
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.size = Pt(9)
        set_cell_background(c, "0F172A")
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    r_specs = [
        ("Processor / CPU", "Dual-Core x86_64 or ARM64", "Apple M-Series / Intel Core i5/i7 (8-Core+)"),
        ("System Memory (RAM)", "4 GB DDR3/DDR4", "8 GB to 16 GB Unified Memory"),
        ("Storage Subsystem", "500 MB Free Disk Space", "NVMe Solid State Drive (SSD)"),
        ("Operating System", "macOS 12+, Ubuntu 20.04+, Windows 10+", "macOS Darwin (Unix Shell Environment)"),
        ("Software Environment", "Python 3.9+, Modern Web Browser", "Python 3.9.6, Safari / Google Chrome, VS Code")
    ]
    for idx, (cat, min_s, tar_s) in enumerate(r_specs):
        row = req_table.rows[idx + 1]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.text = cat
        c1.text = min_s
        c2.text = tar_s
        for c in [c0, c1, c2]:
            c.paragraphs[0].runs[0].font.name = 'Times New Roman'
            c.paragraphs[0].runs[0].font.size = Pt(8.5)
            set_cell_background(c, "F8FAFC" if idx % 2 == 0 else "FFFFFF")

    doc.add_page_break()

    # =========================================================================
    # PAGE 6: 1.3 TOOLS AND TECHNOLOGY USED
    # =========================================================================
    p_body_hdr3 = doc.add_paragraph()
    p_body_hdr3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_bhdr3 = p_body_hdr3.add_run("Page 4 - 5")
    r_bhdr3.font.name = 'Times New Roman'
    r_bhdr3.font.size = Pt(10)
    r_bhdr3.font.italic = True

    add_h2("1.3 Tools and Technology Used (Front-End/Back-End/Framework/Protocol/API/Services/App. etc.)")
    add_para(
        "The project architecture is built upon a modular, industry-standard data analytics stack. Each technology component has been deliberately chosen to provide optimal throughput, statistical rigor, and responsive visualization."
    )

    add_h3("1.3.1 Core Analytics & Back-End Processing")
    add_para(
        "Python serves as the primary programming runtime. It provides native interoperability with high-performance C-extensions for numeric computation and machine learning.",
        bold_prefix="• Python 3.9+: "
    )
    add_para(
        "Utilized as the primary data manipulation engine. Pandas executes vectorized group-by aggregations, datetime parsing, text coercion, and quantile ranking (pd.qcut) with high memory efficiency.",
        bold_prefix="• Pandas (v2.3.3): "
    )
    add_para(
        "Employed for multidimensional array calculations, logarithmic compression (log1p), and Tukey IQR mathematical fence construction.",
        bold_prefix="• NumPy (v2.0.2): "
    )

    add_h3("1.3.2 Machine Learning & Statistical Modeling")
    add_para(
        "Powers the clustering engine. Specifically utilizes sklearn.preprocessing.StandardScaler for Z-score feature standardization, sklearn.cluster.KMeans with k-means++ centroid initialization, and sklearn.metrics.silhouette_score for mathematical cluster validation.",
        bold_prefix="• Scikit-Learn (v1.6.1): "
    )
    add_para(
        "Provides statistical distribution functions and distance metrics supporting unsupervised spatial boundary derivations.",
        bold_prefix="• SciPy (v1.13.1): "
    )

    add_h3("1.3.3 Front-End Presentation & Business Intelligence Layer")
    add_para(
        "A utility-first CSS framework providing a responsive, dark-mode executive UI with flexible grid zones.",
        bold_prefix="• HTML5 & Tailwind CSS: "
    )
    add_para(
        "A reactive JavaScript charting library used to render the segment capital donut chart, spatial customer migration bubble plot, and ML diagnostic curves with hover tooltips.",
        bold_prefix="• Chart.js (v4.4+): "
    )
    add_para(
        "The processed datasets are structured as a Kimball Star Schema (Fact Table joined to Customer Dimension Table), enabling drag-and-drop ingestion into Tableau Desktop and Microsoft Power BI.",
        bold_prefix="• Tableau / Power BI Integration: "
    )

    add_h3("1.3.4 Engineering Tooling, Version Control & Packaging")
    add_para(
        "Git is utilized for local source code version control, with GitHub hosting the central public repository. Automated build scripts compile the Word project report via python-docx and Jupyter notebooks via nbformat.",
        bold_prefix="• Git, GitHub & Tooling: "
    )

    add_caption("Table 1.2: Technology Stack & Architectural Layer Mapping", is_table=True)
    tech_table = doc.add_table(rows=6, cols=3)
    tech_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_hdrs = ["ARCHITECTURAL LAYER", "FRAMEWORK / TOOL", "PRIMARY SYSTEM FUNCTION"]
    for j, h in enumerate(t_hdrs):
        c = tech_table.rows[0].cells[j]
        c.text = h
        c.paragraphs[0].runs[0].font.name = 'Times New Roman'
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.size = Pt(9)
        set_cell_background(c, "0F172A")
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    t_data = [
        ("Data Ingestion & Refinery", "Python, Pandas, NumPy", "Schema typing, null dropping, reversal isolation, IQR filtering"),
        ("Feature Engineering", "Pandas GroupBy, Datetime", "Derivation of Recency, Frequency, Monetary, and AOV metrics"),
        ("Machine Learning", "Scikit-Learn, SciPy", "Log1p scaling, StandardScaler, K-Means clustering, Silhouette score"),
        ("Data Mart & Storage", "CSV / Parquet Data Marts", "Star-schema fact & dimension tables for BI compatibility"),
        ("Executive BI Dashboard", "HTML5, Tailwind CSS, Chart.js", "Interactive KPI cards, spatial bubble plot, searchable activation grid")
    ]
    for idx, (lay, tool, func) in enumerate(t_data):
        row = tech_table.rows[idx + 1]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        c0.text = lay
        c1.text = tool
        c2.text = func
        for c in [c0, c1, c2]:
            c.paragraphs[0].runs[0].font.name = 'Times New Roman'
            c.paragraphs[0].runs[0].font.size = Pt(8.5)
            set_cell_background(c, "F8FAFC" if idx % 2 == 0 else "FFFFFF")

    doc.add_page_break()

    # =========================================================================
    # PAGE 7: 2. SYSTEM DESIGN USING UML - 2.1 DATA FLOW DIAGRAMS (LEVEL 0)
    # =========================================================================
    p_body_hdr4 = doc.add_paragraph()
    p_body_hdr4.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_bhdr4 = p_body_hdr4.add_run("Page 6 - 7")
    r_bhdr4.font.name = 'Times New Roman'
    r_bhdr4.font.size = Pt(10)
    r_bhdr4.font.italic = True

    add_h1("2. SYSTEM DESIGN USING UML:")
    add_h2("2.1 Data Flow Diagrams (1 Level DFD, 2 Level DFD must).")
    add_para(
        "Data Flow Diagrams (DFDs) provide a graphical representation of the flow of data through the information system. They model process transformations, external data entities, data stores, and directional data pathways without embedding implementation-level branching logic. The diagrams follow standard Gane & Sarson conventions:"
    )
    add_para("Represent sources and sinks of data residing outside the system boundary.", bold_prefix="• External Entities (Rectangles): ")
    add_para("Represent algorithmic transformations that manipulate data.", bold_prefix="• Processes (Circles / Rounded Rectangles): ")
    add_para("Represent persistent repositories or tables holding records.", bold_prefix="• Data Stores (Open Parallel Lines): ")
    add_para("Represent directional pipelines carrying formatted information.", bold_prefix="• Data Flows (Directed Arrows): ")

    add_h3("2.1.1 Level 0 Data Flow Diagram (Context-Level DFD)")
    add_para(
        "The Level 0 Context Diagram establishes the global boundary of the system. It positions Process 0.0 ('E-Commerce Customer Lifecycle & RFM Analytics System') at the center, interacting with four external entities:"
    )
    add_para("Provides raw transactional sales streams, item invoices, quantities, and timestamps.", bold_prefix="1. OLTP Database / POS Stream: ")
    add_para("Supplies algorithmic tuning parameters, IQR fences, and cluster hyper-parameters.", bold_prefix="2. Data Scientist / Analytics Engineer: ")
    add_para("Receives stratified customer cohorts, RFM scores, and automated CRM action playbooks.", bold_prefix="3. CRM & Marketing Automation Engine: ")
    add_para("Consumes portfolio health KPIs, revenue concentration donuts, and churn risk metrics.", bold_prefix="4. Executive Leadership & Commercial Team: ")

    # Embed Figure 2.1
    dfd0_path = img_dir / "fig_dfd_level_0.png"
    if dfd0_path.exists():
        doc.add_picture(str(dfd0_path), width=Inches(6.0))
        add_caption("Figure 2.1: Level 0 Data Flow Diagram (Context-Level DFD)")

    doc.add_page_break()

    # =========================================================================
    # PAGE 8: 2.1.2 LEVEL 1 DATA FLOW DIAGRAM
    # =========================================================================
    p_body_hdr5 = doc.add_paragraph()
    p_body_hdr5.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_bhdr5 = p_body_hdr5.add_run("Page 8")
    r_bhdr5.font.name = 'Times New Roman'
    r_bhdr5.font.size = Pt(10)
    r_bhdr5.font.italic = True

    add_h3("2.1.2 Level 1 Data Flow Diagram (End-to-End Pipeline Architecture)")
    add_para(
        "The Level 1 Data Flow Diagram decomposes the system into five primary sequential operational processes and identifies the underlying persistent data stores:"
    )
    add_para("Ingests raw CSV/stream data, validates data types, and logs records to D1 (Raw Transaction Log).", bold_prefix="• Process 1.0 (Ingestion & Schema Enforcement): ")
    add_para("Purges missing CustomerIDs, routes credit notes and cancellations to D2 (Reversals Audit Ledger), and applies Tukey IQR fences on Quantity and UnitPrice.", bold_prefix="• Process 2.0 (Data Quality & Refinery Pipeline): ")
    add_para("Computes snapshot boundary T_obs, aggregates distinct invoices for Frequency, calculates Recency days, and sums net spend, writing to D3 (RFM Analytical Base Table).", bold_prefix="• Process 3.0 (Vectorized RFM Feature Engineering): ")
    add_para("Applies log1p skewness compression and StandardScaler normalization, fits K-Means, and applies quantile 1–5 scoring rules, writing to D4 (Dimensional Customer Mart).", bold_prefix="• Process 4.0 (ML Clustering & Quantile Scoring): ")
    add_para("Packages the data into a star schema and renders the interactive executive web dashboard.", bold_prefix="• Process 5.0 (Star-Schema & Dashboard Export): ")

    # Embed Figure 2.2
    dfd1_path = img_dir / "fig_dfd_level_1.png"
    if dfd1_path.exists():
        doc.add_picture(str(dfd1_path), width=Inches(6.0))
        add_caption("Figure 2.2: Level 1 Data Flow Diagram (End-to-End Pipeline Architecture)")

    doc.add_page_break()

    # =========================================================================
    # PAGE 9: 2.1.3 LEVEL 2 DATA FLOW DIAGRAM (DECOMPOSITION)
    # =========================================================================
    p_body_hdr6 = doc.add_paragraph()
    p_body_hdr6.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_bhdr6 = p_body_hdr6.add_run("Page 9")
    r_bhdr6.font.name = 'Times New Roman'
    r_bhdr6.font.size = Pt(10)
    r_bhdr6.font.italic = True

    add_h3("2.1.3 Level 2 Data Flow Diagram (Sub-Process Decomposition)")
    add_para(
        "To provide rigorous engineering granularity as mandated by syllabus guidelines, the Level 2 DFD decomposes the two most critical computational nodes: Process 2.0 (Data Refinery Pipeline) and Process 4.0 (Algorithmic & Quantile Engine)."
    )

    add_para(
        "Sub-process 2.1 evaluates CustomerID presence, immediately discarding guest records. Sub-process 2.2 identifies cancellation identifiers (InvoiceNo starting with 'C' or Quantity <= 0), segregating them to an audit ledger. Sub-process 2.3 computes Tukey fences (Q1 - 1.5*IQR, Q3 + 1.5*IQR) across positive items, eliminating wholesale distributor anomalies.",
        bold_prefix="• Sub-Process 2.0 Decomposition (Data Refinery): "
    )
    add_para(
        "Sub-process 4.1 performs log1p transformation to correct heavy right-skewed Pareto tails. Sub-process 4.2 executes Z-score standardization. Sub-process 4.3 runs K-Means clustering across k in [2..6] with Silhouette scoring. Sub-process 4.4 runs quantile binning (pd.qcut) to generate 1–5 scores and map 8 business operating tiers.",
        bold_prefix="• Sub-Process 4.0 Decomposition (ML & Scoring Engine): "
    )

    # Embed Figure 2.3
    dfd2_path = img_dir / "fig_dfd_level_2.png"
    if dfd2_path.exists():
        doc.add_picture(str(dfd2_path), width=Inches(6.0))
        add_caption("Figure 2.3: Level 2 Data Flow Diagram (Refinery & ML Scoring Decomposition)")

    doc.add_page_break()

    # =========================================================================
    # PAGE 10: 2.2 USE CASE DIAGRAM & MILESTONE SPECIFICATION
    # =========================================================================
    p_body_hdr7 = doc.add_paragraph()
    p_body_hdr7.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_bhdr7 = p_body_hdr7.add_run("Page 10")
    r_bhdr7.font.name = 'Times New Roman'
    r_bhdr7.font.size = Pt(10)
    r_bhdr7.font.italic = True

    add_h2("2.2 Use Case Diagram.")
    add_para(
        "The Use Case Model delineates the functional boundaries of the system by illustrating interactions between primary human and automated actors and the system's operational capabilities."
    )

    # Embed Figure 2.4
    uc_path = img_dir / "fig_use_case_diagram.png"
    if uc_path.exists():
        doc.add_picture(str(uc_path), width=Inches(5.8))
        add_caption("Figure 2.4: UML Use Case Diagram (System Actors & Functional Capabilities)")

    add_caption("Table 2.1: Formal Use Case Specification Matrix (UC-01 to UC-06)", is_table=True)
    uc_table = doc.add_table(rows=7, cols=4)
    uc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    u_hdrs = ["UC ID", "USE CASE NAME", "PRIMARY ACTOR", "OPERATIONAL GOAL & PRE-CONDITION"]
    for j, h in enumerate(u_hdrs):
        c = uc_table.rows[0].cells[j]
        c.text = h
        c.paragraphs[0].runs[0].font.name = 'Times New Roman'
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.size = Pt(8.5)
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
        c0.text = uid
        c1.text = uname
        c2.text = uact
        c3.text = ugoal
        for c in [c0, c1, c2, c3]:
            c.paragraphs[0].runs[0].font.name = 'Times New Roman'
            c.paragraphs[0].runs[0].font.size = Pt(8)
            set_cell_background(c, "F8FAFC" if idx % 2 == 0 else "FFFFFF")

    add_h2("MILESTONE EVALUATION SUMMARY (24/09/2026)", space_before=14, space_after=6)
    add_para(
        "This submission formally completes the First Progress Work milestone (10 Marks allotted) scheduled for 24/09/2026. All required index sections (1.1, 1.2, 1.3, 2.1, and 2.2) have been fully developed and validated. The technical artifacts, UML models, and data pipelines are complete and ready for supervisor review."
    )

    add_caption("Table 2.2: Project Progress Work Evaluation & Milestone Tracking", is_table=True)
    prog_table = doc.add_table(rows=6, cols=5)
    prog_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    p_hdrs = ["DATE", "ASSESSMENT PARTICULARS", "INDEX SECTIONS", "MARKS", "CURRENT STATUS"]
    for j, h in enumerate(p_hdrs):
        c = prog_table.rows[0].cells[j]
        c.text = h
        c.paragraphs[0].runs[0].font.name = 'Times New Roman'
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.size = Pt(8)
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
        c0.text = dt
        c1.text = part
        c2.text = sec
        c3.text = mrk
        c4.text = stat
        for c in [c0, c1, c2, c3, c4]:
            c.paragraphs[0].runs[0].font.name = 'Times New Roman'
            c.paragraphs[0].runs[0].font.size = Pt(7.5)
            if idx == 0:
                c.paragraphs[0].runs[0].font.bold = True
                set_cell_background(c, "DCFCE7") # Light Green highlight
            else:
                set_cell_background(c, "F8FAFC" if idx % 2 == 0 else "FFFFFF")

    # Output paths
    out_docx_name = "Deep_Koshiya_ProjectReport_Progress1.docx"
    out_docx_path = base_dir / out_docx_name
    doc.save(str(out_docx_path))
    
    # Also overwrite the primary Deep_Koshiya_ProjectReport.docx so all links point to this comprehensive 10-page report!
    primary_docx_path = base_dir / "Deep_Koshiya_ProjectReport.docx"
    doc.save(str(primary_docx_path))
    
    print(f"[SUCCESS] Generated 10-page First Progress Work Report at:")
    print(f"  • {out_docx_path}")
    print(f"  • {primary_docx_path}")

if __name__ == "__main__":
    create_report()
