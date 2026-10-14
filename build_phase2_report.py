"""
build_phase2_report.py
Generates the comprehensive Second Phase Work (Mid-Term Exam) Project Report:
- Milestone: Second Progress Work / Mid-Term Examination (13/10/2026 - 20 Marks)
- Curriculum Scope: Sections 2.3 (Class Diagram), 2.4 (Sequence Diagrams),
  3.1 (Data Dictionary SQL/NoSQL tables), 4.1 (Screen Layouts with Validations),
  building upon the completed foundations of 1.1, 1.2, 1.3, 2.1, 2.2.
- Candidate: Mr. Deep Koshiya
- Degree: B.Sc. in Artificial Intelligence & Data Science, Semester V (5th Semester)
- Institution: Institute of Advanced Research (IAR), Gandhinagar
- NO DEPARTMENT NAME mentioned anywhere in the document.
- Official IAR banner embedded at the top of the cover page.
- Typography: Times New Roman throughout, 12pt body text, 1.15 line spacing, fully justified.
- All 7 publication-quality architectural diagrams included with formal captions.
- Complete Star-Schema Data Dictionary tables with PK/FK, datatypes, constraints, and descriptions.
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

def set_cell_margins(cell, top=30, bottom=30, left=45, right=45):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_cell_border(cell, **kwargs):
    tcPr = cell._element.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
    for edge in ('top', 'left', 'bottom', 'right'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = f'<w:{edge} {nsdecls("w")} w:val="{edge_data.get("val", "single")}" w:sz="{edge_data.get("sz", "4")}" w:space="0" w:color="{edge_data.get("color", "CBD5E1")}"/>'
        else:
            tag = f'<w:{edge} {nsdecls("w")} w:val="none"/>'
        tcBorders.append(parse_xml(tag))
    tcPr.append(tcBorders)

def apply_table_styles(table, col_widths=None):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(table.rows):
        for c_idx, cell in enumerate(row.cells):
            if col_widths and c_idx < len(col_widths):
                cell.width = col_widths[c_idx]
            set_cell_border(cell,
                            top={"val": "single", "sz": "4", "color": "CBD5E1"},
                            bottom={"val": "single", "sz": "4", "color": "CBD5E1"},
                            left={"val": "single", "sz": "4", "color": "CBD5E1"},
                            right={"val": "single", "sz": "4", "color": "CBD5E1"})

def build_phase2_document():
    base_dir = Path("/Users/deep/.gemini/antigravity/scratch/ecommerce_rfm_analytics")
    img_dir = base_dir / "report_images"
    doc = Document()

    # Configure Margins: 1.1 in left (binding), 0.8 in right, 0.75 in top/bottom
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(1.1)
        section.right_margin = Inches(0.8)

    NAVY = RGBColor(15, 23, 42)
    MAROON = RGBColor(122, 0, 60)      # IAR Brand Maroon #7A003C
    INDIGO = RGBColor(30, 27, 75)      # Deep Navy #1E1B4B
    SLATE = RGBColor(71, 85, 105)

    style_normal = doc.styles['Normal']
    font_normal = style_normal.font
    font_normal.name = 'Times New Roman'
    font_normal.size = Pt(12)
    font_normal.color.rgb = NAVY

    def add_para(text, bold_prefix=None, space_after=3.5, space_before=0, line_spacing=1.15, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        if bold_prefix:
            r0 = p.add_run(bold_prefix)
            r0.font.name = 'Times New Roman'
            r0.font.size = Pt(12)
            r0.font.bold = True
            r0.font.color.rgb = NAVY
        r1 = p.add_run(text)
        r1.font.name = 'Times New Roman'
        r1.font.size = Pt(12)
        r1.font.italic = italic
        r1.font.color.rgb = NAVY
        return p

    def add_h1(text, space_before=7, space_after=2.5):
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

    def add_h2(text, space_before=6, space_after=2):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = NAVY
        return p

    def add_h3(text, space_before=4.5, space_after=1.5):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = NAVY
        return p

    def add_caption(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2.5)
        p.paragraph_format.space_after = Pt(3.5)
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = SLATE
        return p

    def format_row_text(cell, text, is_header=False, font_size=Pt(8), bold=False, color=NAVY, align=WD_ALIGN_PARAGRAPH.LEFT):
        cell.text = text
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.05
        for r in p.runs:
            r.font.name = 'Times New Roman'
            r.font.size = font_size
            r.font.bold = bold or is_header
            r.font.color.rgb = RGBColor(255, 255, 255) if is_header else color

    # =========================================================================
    # PAGE 1: COVER PAGE (IAR BANNER AT TOP, NO DEPARTMENT NAME)
    # =========================================================================
    iar_banner_path = img_dir / "iar_banner_hd.png"
    if not iar_banner_path.exists():
        iar_banner_path = img_dir / "iar_banner.png"

    p_banner = doc.add_paragraph()
    p_banner.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_banner.paragraph_format.space_after = Pt(6)
    p_banner.paragraph_format.space_before = Pt(0)
    if iar_banner_path.exists():
        r_ban = p_banner.add_run()
        r_ban.add_picture(str(iar_banner_path), width=Inches(5.4))

    p_inst = doc.add_paragraph()
    p_inst.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_inst.paragraph_format.space_after = Pt(2)
    p_inst.paragraph_format.space_before = Pt(0)
    r_inst = p_inst.add_run("INSTITUTE OF ADVANCED RESEARCH (IAR)")
    r_inst.font.name = 'Times New Roman'
    r_inst.font.size = Pt(14)
    r_inst.font.bold = True
    r_inst.font.color.rgb = MAROON

    p_tag = doc.add_paragraph()
    p_tag.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tag.paragraph_format.space_after = Pt(3)
    r_tag = p_tag.add_run("The University for Innovation | Gandhinagar, Gujarat")
    r_tag.font.name = 'Times New Roman'
    r_tag.font.size = Pt(9.5)
    r_tag.font.italic = True
    r_tag.font.color.rgb = SLATE

    p_prog = doc.add_paragraph()
    p_prog.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_prog.paragraph_format.space_after = Pt(8)
    r_prog = p_prog.add_run("B.Sc. in Artificial Intelligence & Data Science | Semester V (5th Semester)\nAcademic Session: 2026 – 2027")
    r_prog.font.name = 'Times New Roman'
    r_prog.font.size = Pt(11)
    r_prog.font.bold = True
    r_prog.font.color.rgb = INDIGO

    p_div = doc.add_paragraph()
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_div.paragraph_format.space_after = Pt(10)
    r_div = p_div.add_run("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    r_div.font.name = 'Times New Roman'
    r_div.font.size = Pt(9)
    r_div.font.color.rgb = MAROON

    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t.paragraph_format.space_after = Pt(4)
    r_t = p_t.add_run("E-COMMERCE CUSTOMER LIFECYCLE &\nRFM SEGMENTATION DASHBOARD")
    r_t.font.name = 'Times New Roman'
    r_t.font.size = Pt(16.5)
    r_t.font.bold = True
    r_t.font.color.rgb = INDIGO

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    r_s = p_sub.add_run("An Applied Machine Learning & Analytics Engineering System\nfor Commercial Retention & Behavioral Cohort Optimization")
    r_s.font.name = 'Times New Roman'
    r_s.font.size = Pt(10.5)
    r_s.font.italic = True
    r_s.font.color.rgb = NAVY

    card = doc.add_table(rows=5, cols=2)
    card.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_rows = [
        ("Assessment Milestone:", "SECOND PROGRESS WORK / MID-TERM EXAMINATION EVALUATION"),
        ("Scheduled Evaluation Date:", "13th October 2026 (Time: 10:00 AM to 5:00 PM)"),
        ("Curriculum Scope Completed:", "Sections 2.3 (Class Diagram), 2.4 (Sequence Diagrams), 3.1 (Data Dictionary), 4.1 (Screen Layouts with Validations)"),
        ("Course & Project Code:", "Project - 1 (CE-502 / PRJ-701)"),
        ("Evaluation Marks Allotted:", "20 Marks (Mid-Term Assessment Milestone)")
    ]
    for idx, (lbl, val) in enumerate(c_rows):
        row = card.rows[idx]
        c0, c1 = row.cells[0], row.cells[1]
        format_row_text(c0, lbl, font_size=Pt(9), bold=True)
        format_row_text(c1, val, font_size=Pt(9))
        set_cell_background(c0, "F8FAFC")
        set_cell_background(c1, "FFFFFF")
        set_cell_margins(c0, top=25, bottom=25, left=50, right=50)
        set_cell_margins(c1, top=25, bottom=25, left=50, right=50)
    apply_table_styles(card, [Inches(2.2), Inches(4.2)])

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(10)

    sub_table = doc.add_table(rows=2, cols=2)
    sub_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_s1 = sub_table.rows[0].cells[0]
    cell_s2 = sub_table.rows[0].cells[1]
    cell_s1.text = "PROJECT SCHOLAR:\nMr. Deep Koshiya\nB.Sc. in AI & Data Science\nSemester: V (5th Semester)"
    cell_s2.text = "PROJECT SUPERVISION:\nAssigned Project Supervisor\nInstitute of Advanced Research (IAR)\nEvaluation & Verification Stamp"
    for c in [cell_s1, cell_s2]:
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing = 1.05
            for r in p.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(8.5)
        set_cell_background(c, "F1F5F9")
        set_cell_margins(c, top=30, bottom=30, left=40, right=40)

    cell_sig1 = sub_table.rows[1].cells[0]
    cell_sig2 = sub_table.rows[1].cells[1]
    cell_sig1.text = "_________________________\n(Scholar Signature)"
    cell_sig2.text = "_________________________\n(Supervisor Signature)"
    for c in [cell_sig1, cell_sig2]:
        for p in c.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            for r in p.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(8)
        set_cell_background(c, "FFFFFF")
        set_cell_margins(c, top=25, bottom=25, left=40, right=40)
    apply_table_styles(sub_table, [Inches(3.2), Inches(3.2)])

    doc.add_page_break()

    # =========================================================================
    # PAGE I: CERTIFICATE OF SECOND PROGRESS WORK (MID-TERM) & ACKNOWLEDGMENT
    # =========================================================================
    add_para("Page I", space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h1("PROGRESS WORK EVALUATION CERTIFICATE", space_before=2, space_after=4)
    add_para(
        "This is to formally certify that the project work entitled \"E-Commerce Customer Lifecycle & RFM Segmentation Dashboard\" is an authentic record of academic and applied engineering work carried out by Deep Koshiya, student of B.Sc. in Artificial Intelligence & Data Science, Semester V (5th Semester), at Institute of Advanced Research (IAR), Gandhinagar, Gujarat."
    )
    add_para(
        "The candidate has successfully demonstrated and documented the Second Progress Work (Mid-Term Examination) requirements covering Sections 2.3 (UML Class Diagram), 2.4 (UML Sequence Diagrams), 3.1 (Data Dictionary for SQL/NoSQL & Star-Schema Dimensions), and 4.1 (Implementation Screen Layouts with Operational Validations) for Project-1 (CE-502 / PRJ-701) scheduled for evaluation on 13th October 2026."
    )

    cert_table = doc.add_table(rows=4, cols=4)
    cert_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_hdrs = ["EVALUATION MILESTONE", "CURRICULUM SCOPE", "MARKS", "VERIFICATION STATUS"]
    for j, h in enumerate(c_hdrs):
        c = cert_table.rows[0].cells[j]
        format_row_text(c, h, is_header=True, font_size=Pt(7.5), align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_background(c, "7A003C")

    c_vals = [
        ("First Progress Work (24/09/2026)", "Sections 1.1, 1.2, 1.3, 2.1, 2.2", "10 Marks", "COMPLETED & VERIFIED"),
        ("Second Progress Work (13/10/2026)", "Sections 2.3, 2.4, 3.1, 4.1", "20 Marks", "EVALUATED & DEMONSTRATED"),
        ("Faculty Mentor & Reviewer", "Academic Verification Seal", "Signed Copy", "___________________________")
    ]
    for idx, (param, scope, marks, stat) in enumerate(c_vals):
        row = cert_table.rows[idx + 1]
        c0, c1, c2, c3 = row.cells[0], row.cells[1], row.cells[2], row.cells[3]
        format_row_text(c0, param, font_size=Pt(7.5), bold=True)
        format_row_text(c1, scope, font_size=Pt(7.5))
        format_row_text(c2, marks, font_size=Pt(7.5), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_row_text(c3, stat, font_size=Pt(7.5), align=WD_ALIGN_PARAGRAPH.CENTER, bold=(idx==1))
        bg = "DCFCE7" if idx == 1 else ("EFF6FF" if idx == 0 else "FFFFFF")
        for cell in [c0, c1, c2, c3]:
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=20, bottom=20, left=35, right=35)
    apply_table_styles(cert_table, [Inches(2.4), Inches(2.2), Inches(0.8), Inches(1.8)])

    add_h2("ACKNOWLEDGMENT", space_before=8, space_after=2.5)
    add_para(
        "I express my deepest gratitude to my project supervisor, mentors, and the faculty of the Institute of Advanced Research (IAR) for their consistent academic guidance, architectural critiques, and valuable feedback during the object modeling, data dictionary compilation, and dashboard implementation phases."
    )
    add_para(
        "I also thank the institutional administration and the Head of the School for facilitating the computational environments, high-performance data science toolchains, and resources essential to complete Project-1 (CE-502 / PRJ-701) on schedule."
    )
    add_para("Deep Koshiya | B.Sc. in AI & Data Science (Semester V) | Date: 13/10/2026", space_after=3, space_before=3, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)

    doc.add_page_break()

    # =========================================================================
    # PAGE II: ABSTRACT & PROJECT CURRICULUM INDEX
    # =========================================================================
    add_para("Page II", space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h1("ABSTRACT", space_before=1, space_after=2)
    add_para(
        "Customer retention and longitudinal cohort management represent core determinants of profitability in modern digital commerce. The 'E-Commerce Customer Lifecycle & RFM Segmentation Dashboard' provides an automated data engineering and unsupervised machine learning platform that transforms raw, unorganized transaction streams into structured behavioral intelligence."
    )
    add_para(
        "This Second Phase (Mid-Term) report documents the complete structural design and preliminary implementation deliverables. Specifically, it presents the formal UML Class Diagram (Section 2.3) modeling pipeline domain entities; UML Sequence Diagrams (Section 2.4) detailing message dispatch protocols; a comprehensive Star-Schema Data Dictionary (Section 3.1) specifying FACT_TRANSACTIONS, FACT_REVERSALS, and DIM_CUSTOMERS with primary/foreign keys and validation constraints; and high-fidelity Screen Layouts with Operational Input Validations (Section 4.1). These deliverables satisfy all requirements for the 20-mark Mid-Term evaluation milestone."
    )

    add_h2("INDEX OF THE PROJECT REPORT", space_before=5, space_after=1)
    add_para("(Structured strictly according to Syllabus Guidelines for B.Sc. AI and Data Science SEM V)", space_after=3, italic=True)

    idx_table = doc.add_table(rows=12, cols=4)
    idx_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    idx_headers = ["SR.", "CURRICULUM SECTION / TOPIC", "SCHEDULE", "PAGE & STATUS"]
    for j, h in enumerate(idx_headers):
        c = idx_table.rows[0].cells[j]
        format_row_text(c, h, is_header=True, font_size=Pt(8), align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_background(c, "7A003C")

    index_entries = [
        ("*", "Progress Work Evaluation Certificate & Acknowledgment", "13/10/2026", "Page I [Completed]"),
        ("*", "Abstract & Project Curriculum Index", "13/10/2026", "Page II [Completed]"),
        ("1.1-1.3", "About the System: Problem Definition, Requirements & Tools", "24/09/2026", "Pages 1-3 [Completed]"),
        ("2.1", "Data Flow Diagrams (Level 0 Context, Level 1, Level 2 DFDs)", "24/09/2026", "Pages 4-6 [Completed]"),
        ("2.2", "Use Case Diagram & Formal Specification Matrix", "24/09/2026", "Page 7 [Completed]"),
        ("2.3", "UML Class Diagram (Domain Model & Pipeline Objects)", "13/10/2026", "Page 8 [Completed Mid-Term]"),
        ("2.4", "UML Sequence Diagrams (Execution & Activation Workflow)", "13/10/2026", "Page 9 [Completed Mid-Term]"),
        ("3.1", "Data Dictionary: Star-Schema SQL/NoSQL Database Tables", "13/10/2026", "Pages 10-12 [Completed Mid-Term]"),
        ("4.1", "Screen Layouts with Operational Input Validations", "13/10/2026", "Pages 13-14 [Completed Mid-Term]"),
        ("4.2-4.3", "Sample Coding & Execution Navigation", "27/10/2026", "[Third Progress Milestone]"),
        ("5.1-6.1", "Conclusion & References (Books, Papers, Webs)", "27/10/2026", "[Third Progress Milestone]")
    ]

    for idx, (sr, title, sch, pg) in enumerate(index_entries):
        row = idx_table.rows[idx + 1]
        c0, c1, c2, c3 = row.cells[0], row.cells[1], row.cells[2], row.cells[3]
        format_row_text(c0, sr, font_size=Pt(7.5), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_row_text(c1, title, font_size=Pt(7.5), bold=("[Completed" in pg))
        format_row_text(c2, sch, font_size=Pt(7.5), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_row_text(c3, pg, font_size=Pt(7.5), bold=("[Completed" in pg))
        for cell in [c0, c1, c2, c3]:
            set_cell_margins(cell, top=14, bottom=14, left=30, right=30)
            if "Mid-Term" in pg:
                set_cell_background(cell, "DCFCE7")
            elif "[Completed]" in pg:
                set_cell_background(cell, "EFF6FF")
            else:
                set_cell_background(cell, "FFFFFF" if idx % 2 == 0 else "F8FAFC")
    apply_table_styles(idx_table, [Inches(0.6), Inches(3.2), Inches(1.1), Inches(1.5)])

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 1: ABOUT THE SYSTEM (Sections 1.1, 1.2, 1.3)
    # =========================================================================
    add_para("Page 1", space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h1("1. ABOUT THE SYSTEM:")
    add_h2("1.1 Problem definition (Identification of needs):")
    add_para(
        "Modern digital commerce platforms capture high-velocity transactional event streams across web interfaces, mobile shopping applications, and digital points of sale. Standard Online Transaction Processing (OLTP) repositories excel at logging individual line-item orders. However, these systems record transactions rather than customer journeys. Without an automated analytics engineering layer that synthesizes transactional logs into longitudinal customer profiles, commercial decision-makers face severe visibility gaps."
    )

    add_h3("1.1.1 Industry Background & Commercial Challenges")
    add_para(
        "Digital retailers face rising customer acquisition costs (CAC) and declining organic retention. When marketing campaigns treat the customer base as a single homogenous audience, high-value repeat shoppers receive unnecessary discounts that erode gross margins, while dormant accounts slip into permanent churn without timely, targeted reactivation interventions."
    )

    add_h3("1.1.2 Identification of Core Analytical Needs")
    add_para("The system resolves four foundational operational bottlenecks:", bold_prefix="• Core Analytical Needs: ")
    add_para("Raw logs contain anonymous guest sessions, cancellations ('^C'), and wholesale bulk anomalies. An automated data refinery must enforce schema typing, isolate cancellations into an audit store, and cap volume spikes using Tukey's Interquartile Range (IQR) fences.", bold_prefix="1. Data Refinery & Anomaly Isolation: ")
    add_para("The system must distill transaction streams into three normalized customer dimensions: Recency (days elapsed since last order relative to boundary T_obs = max(InvoiceDate) + 1), Frequency (distinct purchase visits), and Monetary spend (cumulative net value), alongside Average Order Value (AOV).", bold_prefix="2. Behavioral Feature Engineering (RFM): ")
    add_para("Customer spending follows extreme Pareto distributions. Monetary skewness must be compressed via log1p transformation and Z-score standardization before fitting unsupervised K-Means models, supported by deterministic 1–5 quantile fallback scoring.", bold_prefix="3. Skew-Resistant Machine Learning: ")
    add_para("Commercial leadership requires an interactive dashboard to visualize portfolio health, monitor at-risk revenue, and export targeted cohort lists directly into CRM marketing campaigns.", bold_prefix="4. Executive Visualization & Activation: ")

    doc.add_page_break()

    # =========================================================================
    # PAGE 2: 1.2 REQUIREMENT SPECIFICATIONS
    # =========================================================================
    add_para("Page 2", space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h2("1.2 Requirement Specifications (Product/System Tasks)")
    add_para(
        "The requirements formally define system capabilities, data validation boundaries, and operational constraints for the E-Commerce Customer Lifecycle & RFM Segmentation platform."
    )

    add_h3("1.2.1 Functional Requirements Specifications (FR)")
    add_para("Ingests raw CSV logs, enforces typing (InvoiceDate to datetime64[ns], Quantity and UnitPrice to floats), and sanitizes product description text.", bold_prefix="• FR-01 (Ingestion & Schema Typing): ")
    add_para("Detects and drops unauthenticated guest checkouts lacking CustomerID to preserve longitudinal cohort integrity.", bold_prefix="• FR-02 (Missing Entity Resolution): ")
    add_para("Isolates credit notes, cancellations ('^C'), and non-positive quantities into an audit store (fact_reversals_audit.csv).", bold_prefix="• FR-03 (Reverse Logistics Segregation): ")
    add_para("Applies Tukey fences [Q1 - 1.5*IQR, Q3 + 1.5*IQR] to eliminate wholesale bulk distributor anomalies.", bold_prefix="• FR-04 (Tukey IQR Outlier Capping): ")
    add_para("Computes snapshot boundary T_obs = max(InvoiceDate) + 1 day and derives customer Recency, Frequency, Spend, and AOV.", bold_prefix="• FR-05 (Vectorized RFM Derivation): ")
    add_para("Applies log1p transformation and StandardScaler, fits K-Means clustering, and assigns 1–5 quantile scores mapping to 8 business cohorts.", bold_prefix="• FR-06 (ML Clustering & Quantile Scoring): ")

    add_h3("1.2.2 Non-Functional Requirements Specifications (NFR)")
    add_para("Processes 15,000+ transaction rows in under 5.0 seconds using vectorized Pandas C-routines.", bold_prefix="• NFR-01 (Performance & Latency): ")
    add_para("Fixed random seeds (seed=42) guarantee 100% deterministic, reproducible clustering across environments.", bold_prefix="• NFR-02 (Algorithmic Reproducibility): ")
    add_para("Modular Python architecture with clean star-schema marts (FACT_TRANSACTIONS, DIM_CUSTOMERS).", bold_prefix="• NFR-03 (Modularity & Architecture): ")

    add_caption("Table 1.1: System Hardware & Software Requirements Specification")
    req_table = doc.add_table(rows=6, cols=3)
    req_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    r_hdrs = ["SYSTEM SPECIFICATION", "MINIMUM REQUIREMENT", "DEPLOYED PROJECT ENVIRONMENT"]
    for j, h in enumerate(r_hdrs):
        c = req_table.rows[0].cells[j]
        format_row_text(c, h, is_header=True, font_size=Pt(8), align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_background(c, "7A003C")

    r_specs = [
        ("Processor / CPU", "Dual-Core x86_64 or ARM64 (2.0 GHz)", "Apple Silicon M-Series / Intel Core i5/i7 (8-Core+)"),
        ("System Memory (RAM)", "4 GB DDR3/DDR4", "8 GB to 16 GB Unified Memory Architecture"),
        ("Storage Subsystem", "500 MB Free Disk Space", "Solid State Drive (NVMe SSD)"),
        ("Operating System", "macOS 12+, Ubuntu 20.04+, Windows 10+", "macOS Darwin (Unix POSIX Environment)"),
        ("Execution Runtimes", "Python 3.9+, Modern Web Browser", "Python 3.9.6, Safari, Google Chrome, VS Code")
    ]
    for idx, (cat, min_s, tar_s) in enumerate(r_specs):
        row = req_table.rows[idx + 1]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        format_row_text(c0, cat, font_size=Pt(7.5), bold=True)
        format_row_text(c1, min_s, font_size=Pt(7.5))
        format_row_text(c2, tar_s, font_size=Pt(7.5))
        for cell in [c0, c1, c2]:
            set_cell_margins(cell, top=16, bottom=16, left=35, right=35)
            set_cell_background(cell, "F8FAFC" if idx % 2 == 0 else "FFFFFF")
    apply_table_styles(req_table, [Inches(1.8), Inches(2.2), Inches(2.4)])

    doc.add_page_break()

    # =========================================================================
    # PAGE 3: 1.3 TOOLS AND TECHNOLOGY USED
    # =========================================================================
    add_para("Page 3", space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h2("1.3 Tools and Technology Used (Front-End/Back-End/Framework/API/Services)")
    add_para(
        "The project implements a modern analytical engineering stack selected for computational throughput, mathematical precision, and interactive visual analytics."
    )

    add_h3("1.3.1 Data Engineering & Numerical Computing")
    add_para("Core programming runtime executing optimized C-extensions for data science.", bold_prefix="• Python 3.9+: ")
    add_para("Vectorized group aggregations, type coercion, datetime logic, and quantile ranking (pd.qcut).", bold_prefix="• Pandas (v2.3.3): ")
    add_para("Multidimensional array manipulations, logarithmic transformations (log1p), and Tukey fence outlier calculations.", bold_prefix="• NumPy (v2.0.2): ")

    add_h3("1.3.2 Machine Learning & Statistical Computing")
    add_para("StandardScaler feature normalization, KMeans with k-means++ initialization, and Silhouette diagnostics.", bold_prefix="• Scikit-Learn (v1.6.1): ")
    add_para("Statistical distributions and spatial distance metric calculations across multidimensional vectors.", bold_prefix="• SciPy (v1.13.1): ")

    add_h3("1.3.3 Presentation, Dashboards & Dimensional Data Marts")
    add_para("Utility-first responsive framework delivering clean dark/light executive UI layouts.", bold_prefix="• HTML5 & Tailwind CSS: ")
    add_para("Client-side canvas graphing engine driving interactive revenue donuts and spatial RFM bubble charts.", bold_prefix="• Chart.js: ")
    add_para("Star-schema dimensional storage (FACT_TRANSACTIONS, DIM_CUSTOMERS) enabling Tableau/Power BI connectivity.", bold_prefix="• Analytical Data Marts: ")

    add_caption("Table 1.2: Technology Stack & Architectural Layer Mapping")
    tech_table = doc.add_table(rows=6, cols=3)
    tech_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_hdrs = ["ARCHITECTURAL LAYER", "TECHNOLOGY / FRAMEWORK", "PRIMARY OPERATIONAL RESPONSIBILITY"]
    for j, h in enumerate(t_hdrs):
        c = tech_table.rows[0].cells[j]
        format_row_text(c, h, is_header=True, font_size=Pt(8), align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_background(c, "7A003C")

    t_data = [
        ("Ingestion & Refinery", "Python, Pandas, NumPy", "Schema typing, null dropping, cancellation segregation, Tukey IQR fences"),
        ("Feature Engineering", "Pandas GroupBy, Datetime", "Derivation of customer Recency, Frequency, Monetary spend, and AOV"),
        ("Machine Learning", "Scikit-Learn, SciPy", "Log1p scaling, StandardScaler, K-Means clustering, Silhouette diagnostics"),
        ("Data Mart Layer", "CSV / Star Schema", "Clean FACT_TRANSACTIONS and DIM_CUSTOMERS marts for BI compatibility"),
        ("Executive BI Dashboard", "HTML5, Tailwind, Chart.js", "Interactive KPI scorecards, spatial RFM bubble charts, activation tables")
    ]
    for idx, (lay, tool, func) in enumerate(t_data):
        row = tech_table.rows[idx + 1]
        c0, c1, c2 = row.cells[0], row.cells[1], row.cells[2]
        format_row_text(c0, lay, font_size=Pt(7.5), bold=True)
        format_row_text(c1, tool, font_size=Pt(7.5))
        format_row_text(c2, func, font_size=Pt(7.5))
        for cell in [c0, c1, c2]:
            set_cell_margins(cell, top=16, bottom=16, left=35, right=35)
            set_cell_background(cell, "F8FAFC" if idx % 2 == 0 else "FFFFFF")
    apply_table_styles(tech_table, [Inches(1.8), Inches(2.0), Inches(2.6)])

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 2: SYSTEM DESIGN USING UML & DFD
    # =========================================================================
    add_para("Page 4", space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h1("2. SYSTEM DESIGN USING UML & DFD:")
    add_h2("2.1 Data Flow Diagrams (Context, Level 1, Level 2 Architecture)")
    add_para(
        "Data Flow Diagrams (DFDs) visually map information transformations through the system. Using Gane & Sarson conventions, the diagrams represent External Entities (rectangles), Processes (rounded rectangles), Data Stores (open horizontal lines), and Data Flows (directed arrows)."
    )

    add_h3("2.1.1 Level 0 Data Flow Diagram (Context-Level DFD)")
    add_para(
        "The Level 0 Context DFD establishes the macro boundary of Process 0.0 ('E-Commerce Customer Lifecycle & RFM Analytics System') and documents all external entity interactions: (1) OLTP Transaction Stream supplying raw sales event logs; (2) Data Scientist providing IQR thresholds and cluster hyper-parameters; (3) CRM Marketing Engine consuming targeted audience lists; and (4) Commercial Executives receiving high-level portfolio KPIs and revenue concentration scorecards."
    )

    df0_path = img_dir / "fig_dfd_level_0.png"
    if df0_path.exists():
        doc.add_picture(str(df0_path), width=Inches(5.1))
        add_caption("Figure 2.1: Level 0 Data Flow Diagram (Context-Level DFD)")

    add_para(
        "This boundary diagram confirms that no internal process detail is revealed at Level 0, ensuring that all subsequent system interactions are fully accounted for before detailed algorithmic decomposition is undertaken."
    )

    doc.add_page_break()

    # =========================================================================
    # PAGE 5: 2.1.2 LEVEL 1 DATA FLOW DIAGRAM
    # =========================================================================
    add_para("Page 5", space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h3("2.1.2 Level 1 Data Flow Diagram (End-to-End Pipeline Architecture)")
    add_para(
        "The Level 1 DFD decomposes Process 0.0 into five sequential operational processes and identifies the underlying persistent data repositories:"
    )
    add_para("Ingests raw order logs, validates schema datatypes, and records input stream to D1 (Raw Transaction Log).", bold_prefix="• Process 1.0 (Ingestion & Schema Typing): ")
    add_para("Purges missing CustomerIDs, routes return orders to D2 (Reversals Audit Ledger), and applies Tukey IQR outlier fences.", bold_prefix="• Process 2.0 (Data Quality & Refinery Pipeline): ")
    add_para("Aggregates customer transactions into Recency days, Frequency counts, and Monetary spend, saving to D3 (RFM Analytical Table).", bold_prefix="• Process 3.0 (Vectorized RFM Feature Extraction): ")
    add_para("Executes log1p transformation, Z-score scaling, K-Means clustering, and 1–5 quantile scoring, persisting to D4 (Segmented Customer Mart).", bold_prefix="• Process 4.0 (ML Clustering & Quantile Scoring): ")
    add_para("Generates the star schema dimensional marts and powers the interactive executive analytics dashboard.", bold_prefix="• Process 5.0 (Star-Schema & Dashboard Export): ")

    df1_path = img_dir / "fig_dfd_level_1.png"
    if df1_path.exists():
        doc.add_picture(str(df1_path), width=Inches(5.1))
        add_caption("Figure 2.2: Level 1 Data Flow Diagram (End-to-End Pipeline Architecture)")

    add_para(
        "The Level 1 architecture enforces strict unidirectional data flow, ensuring that analytical features are never corrupted by unvalidated or reversed transactional inputs."
    )

    doc.add_page_break()

    # =========================================================================
    # PAGE 6: 2.1.3 LEVEL 2 DATA FLOW DIAGRAM
    # =========================================================================
    add_para("Page 6", space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h3("2.1.3 Level 2 Data Flow Diagram (Sub-Process Decomposition)")
    add_para(
        "To provide rigorous engineering granularity as mandated by university project standards, the Level 2 DFD decomposes the two critical computational processes: Process 2.0 (Data Quality Refinery) and Process 4.0 (Machine Learning Clustering Engine)."
    )
    add_para("Sub-process 2.1 removes unauthenticated records lacking CustomerID. Sub-process 2.2 detects cancellation codes ('^C' or negative quantities) and writes to the audit ledger. Sub-process 2.3 computes Tukey fences [Q1 - 1.5*IQR, Q3 + 1.5*IQR] to eliminate wholesale bulk distributor anomalies.", bold_prefix="• Sub-Process 2.0 Decomposition (Refinery): ")
    add_para("Sub-process 4.1 applies log1p compression (z = ln(1 + x)) to resolve Pareto spend skewness. Sub-process 4.2 executes Z-score normalization. Sub-process 4.3 executes K-Means clustering across k in [2..6] with Silhouette validation. Sub-process 4.4 runs quantile binning (pd.qcut) mapping accounts to 8 business operating cohorts.", bold_prefix="• Sub-Process 4.0 Decomposition (ML & Scoring): ")

    df2_path = img_dir / "fig_dfd_level_2.png"
    if df2_path.exists():
        doc.add_picture(str(df2_path), width=Inches(5.2))
        add_caption("Figure 2.3: Level 2 Data Flow Diagram (Refinery & ML Decomposition)")

    add_para(
        "This granular decomposition formally documents the mathematical transformations and algorithmic filtering required to turn raw transactional records into production-grade machine learning features."
    )

    doc.add_page_break()

    # =========================================================================
    # PAGE 7: 2.2 USE CASE DIAGRAM & SPECIFICATION MATRIX
    # =========================================================================
    add_para("Page 7", space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h2("2.2 Use Case Diagram & Actor Specifications")
    add_para(
        "The Use Case Model defines operational boundaries by mapping system capabilities to primary human and automated actors.",
        space_after=2
    )

    uc_path = img_dir / "fig_use_case_diagram.png"
    if uc_path.exists():
        doc.add_picture(str(uc_path), width=Inches(3.3))
        add_caption("Figure 2.4: UML Use Case Diagram (System Actors & Functional Capabilities)")

    add_caption("Table 2.1: Formal Use Case Specification Matrix (UC-01 to UC-06)")
    uc_table = doc.add_table(rows=7, cols=4)
    uc_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    u_hdrs = ["UC ID", "USE CASE NAME", "PRIMARY ACTOR", "OPERATIONAL GOAL & PRE-CONDITION"]
    for j, h in enumerate(u_hdrs):
        c = uc_table.rows[0].cells[j]
        format_row_text(c, h, is_header=True, font_size=Pt(7.5), align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_background(c, "7A003C")

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
        format_row_text(c0, uid, font_size=Pt(7), bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        format_row_text(c1, uname, font_size=Pt(7), bold=True)
        format_row_text(c2, uact, font_size=Pt(7))
        format_row_text(c3, ugoal, font_size=Pt(7))
        for cell in [c0, c1, c2, c3]:
            set_cell_margins(cell, top=14, bottom=14, left=25, right=25)
            set_cell_background(cell, "F8FAFC" if idx % 2 == 0 else "FFFFFF")
    apply_table_styles(uc_table, [Inches(0.65), Inches(1.85), Inches(1.5), Inches(2.4)])

    doc.add_page_break()

    # =========================================================================
    # PAGE 8: 2.3 UML CLASS DIAGRAM (NEW FOR MID-TERM EVALUATION)
    # =========================================================================
    add_para("Page 8", space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h2("2.3 UML Class Diagram (Domain Model & Object Architecture)")
    add_para(
        "The UML Class Diagram establishes the object-oriented structure of the analytics pipeline, modeling domain entities, transformation services, algorithmic engines, and dimensional export utilities."
    )

    cls_path = img_dir / "fig_class_diagram.png"
    if cls_path.exists():
        doc.add_picture(str(cls_path), width=Inches(5.3))
        add_caption("Figure 2.5: UML Class Diagram (Domain Model & Pipeline Object Architecture)")

    add_h3("2.3.1 Detailed Object Model & Class Specifications")
    add_para("Models single line-item order entities with methods for session validation and gross spend calculation.", bold_prefix="1. TransactionRecord (Domain Entity): ")
    add_para("Controls raw data loading, schema typing, null dropping, cancellation segregation, and Tukey IQR outlier capping.", bold_prefix="2. CleanRefineryPipeline (Refinery Service): ")
    add_para("Vectorized service computing customer Recency relative to T_obs, purchase Frequency, Monetary spend, and Average Order Value.", bold_prefix="3. RFMFeatureEngineer (Feature Engineering): ")
    add_para("Encapsulates feature scaling, log1p compression, StandardScaler, K-Means clustering, and Silhouette diagnostics.", bold_prefix="4. KMeansSegmentationEngine (ML Service): ")
    add_para("Rule-based quantile engine assigning deterministic 1–5 scores and mapping 8 business operational cohorts.", bold_prefix="5. QuantileStratifier (Business Logic Engine): ")
    add_para("Generates relational star-schema tables (FACT_TRANSACTIONS, FACT_REVERSALS, DIM_CUSTOMERS, DIM_PRODUCTS).", bold_prefix="6. StarSchemaMartExporter (Persistence Service): ")

    doc.add_page_break()

    # =========================================================================
    # PAGE 9: 2.4 UML SEQUENCE DIAGRAMS (NEW FOR MID-TERM EVALUATION)
    # =========================================================================
    add_para("Page 9", space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h2("2.4 UML Sequence Diagrams (System Execution & Interaction Flow)")
    add_para(
        "The UML Sequence Diagram models dynamic behavioral interactions and message dispatch sequences across system components during pipeline execution."
    )

    seq_path = img_dir / "fig_sequence_diagram.png"
    if seq_path.exists():
        doc.add_picture(str(seq_path), width=Inches(5.3))
        add_caption("Figure 2.6: UML Sequence Diagram (End-to-End Processing & Activation Workflow)")

    add_h3("2.4.1 Step-by-Step Message Dispatch Protocol")
    add_para("The Data Engineer or automated scheduler issues trigger_pipeline(raw_csv) to initiate ingestion.", bold_prefix="• Step 1 (Pipeline Trigger): ")
    add_para("CleanRefineryPipeline performs self-calls to filter unauthenticated nulls, segregate cancellations, and apply Tukey IQR fences.", bold_prefix="• Steps 2-3 (Data Refinery): ")
    add_para("The cleansed dataframe is passed to RFMFeatureEngineer, which evaluates T_obs and extracts customer vectors.", bold_prefix="• Steps 4-5 (Feature Derivation): ")
    add_para("The RFM base table is dispatched to KMeansEngine, which applies log1p compression and trains k=4 clusters.", bold_prefix="• Steps 6-7 (Machine Learning Execution): ")
    add_para("Cluster predictions and quantile tiers are exported to StarSchemaDataMart, writing dimension and fact tables.", bold_prefix="• Steps 8-9 (Data Mart Export): ")
    add_para("The data mart generates summary JSON payloads, enabling ExecutiveDashboard to render interactive visualizations.", bold_prefix="• Steps 10-11 (Dashboard Activation): ")

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 3: DATA DICTIONARY (NEW FOR MID-TERM EVALUATION)
    # =========================================================================
    add_para("Page 10", space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h1("3. DATA DICTIONARY:")
    add_h2("3.1 Database tables SQL/NoSQL (Dimensional Schema Specification)")
    add_para(
        "The Data Dictionary provides complete schema specifications for the star-schema dimensional data mart architecture. The mart decouples high-volume transactional event facts from customer and product dimensions, ensuring maximum SQL query performance and native compatibility with enterprise BI tools like Tableau and Power BI."
    )

    add_caption("Table 3.1: FACT_TRANSACTIONS (Order-Level Fact Mart)")
    f_table = doc.add_table(rows=10, cols=5)
    f_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    f_hdrs = ["COLUMN NAME", "DATA TYPE", "KEY", "NULL", "OPERATIONAL DESCRIPTION & CONSTRAINTS"]
    for j, h in enumerate(f_hdrs):
        c = f_table.rows[0].cells[j]
        format_row_text(c, h, is_header=True, font_size=Pt(7.5), align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_background(c, "7A003C")

    f_data = [
        ("transaction_id", "BIGINT", "PK", "NO", "Surrogate primary key generated via monotonic sequence."),
        ("invoice_no", "VARCHAR(20)", "None", "NO", "Alphanumeric invoice identifier; 6-digit order code."),
        ("stock_code", "VARCHAR(20)", "FK", "NO", "Stock keeping unit (SKU) linking to DIM_PRODUCTS."),
        ("quantity", "INTEGER", "None", "NO", "Purchased item quantity; strictly Quantity > 0 after refinery."),
        ("invoice_date", "TIMESTAMP", "FK", "NO", "Transaction timestamp; parsed to datetime64[ns]."),
        ("unit_price", "DECIMAL(10,2)", "None", "NO", "Unit price per product item; strictly UnitPrice > 0.00."),
        ("customer_id", "VARCHAR(20)", "FK", "NO", "Customer identifier linking to DIM_CUSTOMERS."),
        ("country", "VARCHAR(50)", "None", "YES", "Geographic country destination of invoice."),
        ("gross_amount", "DECIMAL(12,2)", "None", "NO", "Calculated line spend: Quantity * UnitPrice.")
    ]
    for idx, (col, dtype, key, nl, desc) in enumerate(f_data):
        row = f_table.rows[idx + 1]
        c0, c1, c2, c3, c4 = row.cells[0], row.cells[1], row.cells[2], row.cells[3], row.cells[4]
        format_row_text(c0, col, font_size=Pt(7), bold=True)
        format_row_text(c1, dtype, font_size=Pt(7))
        format_row_text(c2, key, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER, bold=(key=="PK"))
        format_row_text(c3, nl, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_row_text(c4, desc, font_size=Pt(7))
        for cell in [c0, c1, c2, c3, c4]:
            set_cell_margins(cell, top=14, bottom=14, left=25, right=25)
            set_cell_background(cell, "F8FAFC" if idx % 2 == 0 else "FFFFFF")
    apply_table_styles(f_table, [Inches(1.2), Inches(1.1), Inches(0.6), Inches(0.6), Inches(2.9)])

    doc.add_page_break()

    # =========================================================================
    # PAGE 11: DATA DICTIONARY - FACT_REVERSALS & DIM_CUSTOMERS
    # =========================================================================
    add_para("Page 11", space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_caption("Table 3.2: FACT_REVERSALS (Reverse Logistics Audit Fact Mart)")
    r_table = doc.add_table(rows=8, cols=5)
    r_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(f_hdrs):
        c = r_table.rows[0].cells[j]
        format_row_text(c, h, is_header=True, font_size=Pt(7.5), align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_background(c, "7A003C")

    r_data = [
        ("reversal_id", "BIGINT", "PK", "NO", "Surrogate primary key for return audit records."),
        ("invoice_no", "VARCHAR(20)", "None", "NO", "Cancellation invoice code matching regex '^C[0-9]+'."),
        ("stock_code", "VARCHAR(20)", "FK", "NO", "Product item identifier returned by customer."),
        ("return_quantity", "INTEGER", "None", "NO", "Negative quantity returned (Quantity < 0)."),
        ("reversal_date", "TIMESTAMP", "FK", "NO", "Timestamp when cancellation was processed."),
        ("unit_price", "DECIMAL(10,2)", "None", "NO", "Original sale unit price of returned item."),
        ("customer_id", "VARCHAR(20)", "FK", "NO", "Customer ID associated with the reversal transaction.")
    ]
    for idx, (col, dtype, key, nl, desc) in enumerate(r_data):
        row = r_table.rows[idx + 1]
        c0, c1, c2, c3, c4 = row.cells[0], row.cells[1], row.cells[2], row.cells[3], row.cells[4]
        format_row_text(c0, col, font_size=Pt(7), bold=True)
        format_row_text(c1, dtype, font_size=Pt(7))
        format_row_text(c2, key, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER, bold=(key=="PK"))
        format_row_text(c3, nl, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_row_text(c4, desc, font_size=Pt(7))
        for cell in [c0, c1, c2, c3, c4]:
            set_cell_margins(cell, top=14, bottom=14, left=25, right=25)
            set_cell_background(cell, "F8FAFC" if idx % 2 == 0 else "FFFFFF")
    apply_table_styles(r_table, [Inches(1.2), Inches(1.1), Inches(0.6), Inches(0.6), Inches(2.9)])

    add_caption("Table 3.3: DIM_CUSTOMERS (Customer RFM & Behavioral Dimension Mart)")
    c_table = doc.add_table(rows=11, cols=5)
    c_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(f_hdrs):
        c = c_table.rows[0].cells[j]
        format_row_text(c, h, is_header=True, font_size=Pt(7.5), align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_background(c, "7A003C")

    c_data = [
        ("customer_key", "BIGINT", "PK", "NO", "Surrogate primary key for customer dimension."),
        ("customer_id", "VARCHAR(20)", "None", "NO", "Natural unique customer identifier (non-null)."),
        ("recency_days", "INTEGER", "None", "NO", "Days elapsed since last order: T_obs - max(InvoiceDate)."),
        ("frequency_orders", "INTEGER", "None", "NO", "Distinct purchase trips: count(distinct InvoiceNo)."),
        ("monetary_total", "DECIMAL(12,2)", "None", "NO", "Cumulative net customer spend across all valid orders."),
        ("avg_order_value", "DECIMAL(10,2)", "None", "NO", "Average Order Value: Monetary / Frequency."),
        ("r_score", "SMALLINT", "None", "NO", "Quantile Recency score (1 to 5; 5 = most recent)."),
        ("f_score", "SMALLINT", "None", "NO", "Quantile Frequency score (1 to 5; 5 = highest frequency)."),
        ("m_score", "SMALLINT", "None", "NO", "Quantile Monetary score (1 to 5; 5 = highest spend)."),
        ("segment_tier", "VARCHAR(40)", "None", "NO", "Mapped operating tier ('Champions', 'Loyal', 'At Risk').")
    ]
    for idx, (col, dtype, key, nl, desc) in enumerate(c_data):
        row = c_table.rows[idx + 1]
        c0, c1, c2, c3, c4 = row.cells[0], row.cells[1], row.cells[2], row.cells[3], row.cells[4]
        format_row_text(c0, col, font_size=Pt(7), bold=True)
        format_row_text(c1, dtype, font_size=Pt(7))
        format_row_text(c2, key, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER, bold=(key=="PK"))
        format_row_text(c3, nl, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_row_text(c4, desc, font_size=Pt(7))
        for cell in [c0, c1, c2, c3, c4]:
            set_cell_margins(cell, top=14, bottom=14, left=25, right=25)
            set_cell_background(cell, "F8FAFC" if idx % 2 == 0 else "FFFFFF")
    apply_table_styles(c_table, [Inches(1.2), Inches(1.1), Inches(0.6), Inches(0.6), Inches(2.9)])

    doc.add_page_break()

    # =========================================================================
    # PAGE 12: DATA DICTIONARY - DIM_PRODUCTS & DIM_DATE
    # =========================================================================
    add_para("Page 12", space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_caption("Table 3.4: DIM_PRODUCTS (Catalog Item Dimension Mart)")
    p_table = doc.add_table(rows=6, cols=5)
    p_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(f_hdrs):
        c = p_table.rows[0].cells[j]
        format_row_text(c, h, is_header=True, font_size=Pt(7.5), align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_background(c, "7A003C")

    p_data = [
        ("product_key", "BIGINT", "PK", "NO", "Surrogate primary key for product dimension."),
        ("stock_code", "VARCHAR(20)", "None", "NO", "Unique SKU code assigned to the retail inventory item."),
        ("description", "VARCHAR(255)", "None", "YES", "Cleaned product title text; sanitized of special characters."),
        ("unit_price_ref", "DECIMAL(10,2)", "None", "NO", "Baseline catalog unit price for retail inventory item."),
        ("is_active", "BOOLEAN", "None", "NO", "Product active flag (TRUE = currently sold).")
    ]
    for idx, (col, dtype, key, nl, desc) in enumerate(p_data):
        row = p_table.rows[idx + 1]
        c0, c1, c2, c3, c4 = row.cells[0], row.cells[1], row.cells[2], row.cells[3], row.cells[4]
        format_row_text(c0, col, font_size=Pt(7), bold=True)
        format_row_text(c1, dtype, font_size=Pt(7))
        format_row_text(c2, key, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER, bold=(key=="PK"))
        format_row_text(c3, nl, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_row_text(c4, desc, font_size=Pt(7))
        for cell in [c0, c1, c2, c3, c4]:
            set_cell_margins(cell, top=14, bottom=14, left=25, right=25)
            set_cell_background(cell, "F8FAFC" if idx % 2 == 0 else "FFFFFF")
    apply_table_styles(p_table, [Inches(1.2), Inches(1.1), Inches(0.6), Inches(0.6), Inches(2.9)])

    add_caption("Table 3.5: DIM_DATE (Temporal Dimension Mart)")
    d_table = doc.add_table(rows=7, cols=5)
    d_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(f_hdrs):
        c = d_table.rows[0].cells[j]
        format_row_text(c, h, is_header=True, font_size=Pt(7.5), align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_background(c, "7A003C")

    d_data = [
        ("date_key", "INTEGER", "PK", "NO", "Integer key formatted as YYYYMMDD (e.g. 20261013)."),
        ("full_date", "DATE", "None", "NO", "Calendar calendar date."),
        ("day_of_week", "VARCHAR(10)", "None", "NO", "Day name: 'Monday', 'Tuesday', ..., 'Sunday'."),
        ("calendar_month", "VARCHAR(15)", "None", "NO", "Month name: 'January', 'October', etc."),
        ("calendar_quarter", "SMALLINT", "None", "NO", "Calendar quarter (1 to 4)."),
        ("is_weekend", "BOOLEAN", "None", "NO", "Weekend flag (TRUE if Saturday or Sunday).")
    ]
    for idx, (col, dtype, key, nl, desc) in enumerate(d_data):
        row = d_table.rows[idx + 1]
        c0, c1, c2, c3, c4 = row.cells[0], row.cells[1], row.cells[2], row.cells[3], row.cells[4]
        format_row_text(c0, col, font_size=Pt(7), bold=True)
        format_row_text(c1, dtype, font_size=Pt(7))
        format_row_text(c2, key, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER, bold=(key=="PK"))
        format_row_text(c3, nl, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_row_text(c4, desc, font_size=Pt(7))
        for cell in [c0, c1, c2, c3, c4]:
            set_cell_margins(cell, top=14, bottom=14, left=25, right=25)
            set_cell_background(cell, "F8FAFC" if idx % 2 == 0 else "FFFFFF")
    apply_table_styles(d_table, [Inches(1.2), Inches(1.1), Inches(0.6), Inches(0.6), Inches(2.9)])

    doc.add_page_break()

    # =========================================================================
    # CHAPTER 4: IMPLEMENTATION (NEW FOR MID-TERM EVALUATION)
    # =========================================================================
    add_para("Page 13", space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h1("4. IMPLEMENTATION:")
    add_h2("4.1 Screen Layouts with validations")
    add_para(
        "The web-based BI dashboard provides four synchronized analytical views designed with utility-first Tailwind CSS and Chart.js. Each interface incorporates strict client-side and server-side data validations to ensure reliable commercial operation."
    )

    scr_path = img_dir / "fig_screen_layouts.png"
    if scr_path.exists():
        doc.add_picture(str(scr_path), width=Inches(5.3))
        add_caption("Figure 4.1: Implementation Screen Layouts with Operational Input Validations")

    add_h3("4.1.1 Screen Layouts Architectural Walkthrough")
    add_para("Renders four dynamic KPI scorecards displaying portfolio spend ($65,049.20), active accounts (256), average order value ($254.10), and at-risk revenue share (7.3%).", bold_prefix="• Screen 1 (Executive Overview & Scorecards): ")
    add_para("Canvas scatterplot plotting Recency (X-axis) against Monetary Spend (Y-axis) with bubble radii scaled to Frequency.", bold_prefix="• Screen 2 (RFM Behavioral Scatterplot): ")
    add_para("Searchable ledger with instantaneous regex filtering, pagination controls, and targeted CSV export capabilities.", bold_prefix="• Screen 3 (Customer Cohort Activation Grid): ")
    add_para("Isolated audit table displaying quarantined cancellations and credit notes for compliance review.", bold_prefix="• Screen 4 (Reverse Logistics Audit Ledger): ")

    doc.add_page_break()

    # =========================================================================
    # PAGE 14: IMPLEMENTATION VALIDATIONS & ASSESSMENT SCHEDULE
    # =========================================================================
    add_para("Page 14", space_after=1, align=WD_ALIGN_PARAGRAPH.RIGHT, italic=True)
    add_h3("4.1.2 Formal Input & Operational Validation Architecture")
    add_para("The dashboard executes four discrete validation rules to prevent corrupt states:", bold_prefix="• System Validation Rules: ")
    add_para("Monetary totals are constrained to non-negative floats (Spend >= 0.00). Active customer counts verify distinct non-null CustomerIDs.", bold_prefix="1. VAL-01 (Scorecard Numerical Integrity): ")
    add_para("Scatterplot coordinates enforce bounded ranges: Recency in [1..365], Frequency in [1..50], Spend in [0..10,000]. Hover tooltips validate customer existence.", bold_prefix="2. VAL-02 (Spatial Canvas Boundary Constraints): ")
    add_para("Client search input is sanitized against script injection. Table pagination enforces integer bounds (1 <= Page <= MaxPage).", bold_prefix="3. VAL-03 (Search Sanitization & Pagination Bounds): ")
    add_para("Reversals strictly enforce cancellation regex ('^C') and negative quantities (Quantity < 0) before insertion into audit stores.", bold_prefix="4. VAL-04 (Audit Ledger Quarantine Logic): ")

    add_h2("PROJECT PROGRESS WORK EVALUATION SCHEDULE (MID-TERM)", space_before=4, space_after=1)
    add_caption("Table 4.1: Project Progress Work Evaluation Schedule & Milestone Tracking")
    prog_table = doc.add_table(rows=6, cols=5)
    prog_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    p_hdrs = ["DATE", "ASSESSMENT PARTICULARS", "INDEX SECTIONS", "MARKS", "CURRENT STATUS"]
    for j, h in enumerate(p_hdrs):
        c = prog_table.rows[0].cells[j]
        format_row_text(c, h, is_header=True, font_size=Pt(7), align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_background(c, "7A003C")

    p_data = [
        ("24/09/2026", "First Progress Work", "1.1, 1.2, 1.3, 2.1, 2.2", "10 Marks", "COMPLETED & VERIFIED"),
        ("13/10/2026", "Mid-Term Exam (Second Progress)", "2.3, 2.4, 3.1, 4.1", "20 Marks", "DEMONSTRATED (Current Submission)"),
        ("27/10/2026", "Third Progress Work", "4.2, 4.3, 5.1, 5.2, 6.1", "10 Marks", "Scheduled for Third Milestone"),
        ("06/11/2026", "Report Submission (Spiral Binding)", "Final Signed Spiral Copy", "10 Marks", "Scheduled for Final Submission"),
        ("18-20/11/2026", "Final Exam (SEE)", "Exam & Viva Voce", "50 Marks", "Scheduled for Semester Examination")
    ]
    for idx, (dt, part, sec, mrk, stat) in enumerate(p_data):
        row = prog_table.rows[idx + 1]
        c0, c1, c2, c3, c4 = row.cells[0], row.cells[1], row.cells[2], row.cells[3], row.cells[4]
        format_row_text(c0, dt, font_size=Pt(6.5), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_row_text(c1, part, font_size=Pt(6.5), bold=(idx==1))
        format_row_text(c2, sec, font_size=Pt(6.5), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_row_text(c3, mrk, font_size=Pt(6.5), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_row_text(c4, stat, font_size=Pt(6.5), bold=(idx==1))
        for cell in [c0, c1, c2, c3, c4]:
            set_cell_margins(cell, top=14, bottom=14, left=25, right=25)
            if idx == 1:
                set_cell_background(cell, "DCFCE7")
            elif idx == 0:
                set_cell_background(cell, "EFF6FF")
            else:
                set_cell_background(cell, "F8FAFC" if idx % 2 == 0 else "FFFFFF")
    apply_table_styles(prog_table, [Inches(0.9), Inches(1.7), Inches(1.3), Inches(0.7), Inches(1.8)])

    out_docx_path = base_dir / "Deep_Koshiya_ProjectReport_Progress2.docx"
    doc.save(str(out_docx_path))
    doc.save(str(base_dir / "Deep_Koshiya_ProjectReport.docx"))
    print(f"[SUCCESS] Second Phase Report saved to {out_docx_path}")

if __name__ == "__main__":
    build_phase2_document()
