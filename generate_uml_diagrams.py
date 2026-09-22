"""
generate_uml_diagrams.py
Generates publication-quality diagrams for Chapter 2:
- Figure 2.1: Level 0 DFD (Context Level)
- Figure 2.2: Level 1 DFD (Decomposed System Architecture)
- Figure 2.3: Level 2 DFD (Refinery & ML Scoring Decomposition)
- Figure 2.4: UML Use Case Diagram
"""

import matplotlib.pyplot as plt
import matplotlib.patches as patches
from pathlib import Path

img_dir = Path("/Users/deep/.gemini/antigravity/scratch/ecommerce_rfm_analytics/report_images")
img_dir.mkdir(parents=True, exist_ok=True)

# -------------------------------------------------------------
# 1. Level 0 Context DFD
# -------------------------------------------------------------
def create_level_0_dfd():
    fig, ax = plt.subplots(figsize=(10, 5), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")

    # Central Process Circle
    central_circle = patches.Circle((50, 50), 18, facecolor="#4f46e5", edgecolor="#312e81", linewidth=2.5, zorder=3)
    ax.add_patch(central_circle)
    ax.text(50, 53, "0.0\nE-COMMERCE RFM &\nCUSTOMER LIFECYCLE", ha="center", va="center", color="white", fontsize=10, fontweight="bold", zorder=4)
    ax.text(50, 42, "ANALYTICS SYSTEM", ha="center", va="center", color="#c7d2fe", fontsize=8.5, fontweight="bold", zorder=4)

    # External Entities (Rectangles)
    entities = [
        ("OLTP Transaction\nDatabase / POS", (5, 65), (22, 16), "#0f172a"),
        ("Analytics Engineer /\nData Scientist", (5, 18), (22, 16), "#0f172a"),
        ("CRM & Marketing\nAutomation Engine", (73, 65), (22, 16), "#0f172a"),
        ("Executive Leadership /\nCommercial Team", (73, 18), (22, 16), "#0f172a")
    ]

    for name, (x, y), (w, h), color in entities:
        rect = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=1.5", facecolor=color, edgecolor="#475569", linewidth=2, zorder=2)
        ax.add_patch(rect)
        ax.text(x + w/2, y + h/2, name, ha="center", va="center", color="white", fontsize=8.5, fontweight="bold", zorder=4)

    # Arrows & Data Flows
    arrow_props = dict(arrowstyle="->", color="#334155", lw=1.8, mutation_scale=14)
    
    # Inflow from OLTP to Central
    ax.annotate("", xy=(33, 58), xytext=(27, 70), arrowprops=arrow_props)
    ax.text(28, 67, "Raw Sales Orders\n(Invoice Logs)", fontsize=7.5, color="#1e293b", fontweight="bold", rotation=25)

    # Inflow from Engineer to Central
    ax.annotate("", xy=(34, 43), xytext=(27, 28), arrowprops=arrow_props)
    ax.text(26, 33, "Tukey Multiplier &\nk-Value Parameters", fontsize=7.5, color="#1e293b", fontweight="bold", rotation=-25)

    # Outflow to CRM
    ax.annotate("", xy=(73, 70), xytext=(67, 58), arrowprops=arrow_props)
    ax.text(66, 68, "Targeted Cohorts &\nCRM Action Playbooks", fontsize=7.5, color="#1e293b", fontweight="bold", rotation=-25)

    # Outflow to Executive
    ax.annotate("", xy=(73, 28), xytext=(66, 43), arrowprops=arrow_props)
    ax.text(66, 32, "Executive BI Dashboards\n& Portfolio Metrics", fontsize=7.5, color="#1e293b", fontweight="bold", rotation=25)

    plt.title("Figure 2.1: Level 0 Data Flow Diagram (Context-Level DFD)", fontsize=11, fontweight="bold", pad=12, color="#0f172a")
    plt.tight_layout()
    plt.savefig(img_dir / "fig_dfd_level_0.png")
    plt.close()

# -------------------------------------------------------------
# 2. Level 1 DFD
# -------------------------------------------------------------
def create_level_1_dfd():
    fig, ax = plt.subplots(figsize=(11, 5.5), dpi=300)
    ax.set_xlim(0, 110)
    ax.set_ylim(0, 80)
    ax.axis("off")

    # Processes (Numbered Rounded Boxes)
    processes = [
        ("1.0 Ingest Raw\nTransactional Logs", (8, 48), "#4f46e5"),
        ("2.0 Data Quality &\nRefinery Pipeline", (29, 48), "#4f46e5"),
        ("3.0 Vectorized RFM\nFeature Engineering", (50, 48), "#4f46e5"),
        ("4.0 ML Clustering &\nQuantile Scoring", (71, 48), "#4f46e5"),
        ("5.0 Star-Schema &\nDashboard Export", (92, 48), "#4f46e5"),
    ]

    for label, (x, y), color in processes:
        box = patches.FancyBboxPatch((x-7, y-7), 14, 14, boxstyle="round,pad=1.2", facecolor=color, edgecolor="#312e81", lw=2, zorder=3)
        ax.add_patch(box)
        ax.text(x, y, label, ha="center", va="center", color="white", fontsize=7.5, fontweight="bold", zorder=4)

    # Data Stores (Parallel Horizontal Lines)
    stores = [
        ("D1: Raw Transaction Log", (18, 15)),
        ("D2: Reversals Audit Ledger", (39, 15)),
        ("D3: RFM Analytical Base Table", (60, 15)),
        ("D4: Dimensional Customer Mart", (82, 15)),
    ]

    for name, (x, y) in stores:
        # Top and bottom line of open store
        ax.plot([x-8, x+8], [y+4, y+4], color="#0f172a", lw=2)
        ax.plot([x-8, x+8], [y-4, y-4], color="#0f172a", lw=2)
        ax.text(x, y, name, ha="center", va="center", color="#0f172a", fontsize=7, fontweight="bold")

    arrow_props = dict(arrowstyle="->", color="#334155", lw=1.6, mutation_scale=12)

    # Process to Process Flows
    ax.annotate("", xy=(22, 48), xytext=(15, 48), arrowprops=arrow_props)
    ax.annotate("", xy=(43, 48), xytext=(36, 48), arrowprops=arrow_props)
    ax.annotate("", xy=(64, 48), xytext=(57, 48), arrowprops=arrow_props)
    ax.annotate("", xy=(85, 48), xytext=(78, 48), arrowprops=arrow_props)

    # Store Interactions
    # 1.0 reads/writes to D1
    ax.annotate("", xy=(18, 20), xytext=(10, 40), arrowprops=arrow_props)
    # 2.0 writes reversals to D2
    ax.annotate("", xy=(39, 20), xytext=(32, 40), arrowprops=arrow_props)
    ax.text(37, 28, "Reversals\n(^C, Q<=0)", fontsize=6.5, color="#dc2626", fontweight="bold")
    # 3.0 writes ABT to D3
    ax.annotate("", xy=(60, 20), xytext=(53, 40), arrowprops=arrow_props)
    # 4.0 writes to D4
    ax.annotate("", xy=(82, 20), xytext=(74, 40), arrowprops=arrow_props)
    # 5.0 reads from D4 to Dashboard
    ax.annotate("", xy=(92, 40), xytext=(85, 20), arrowprops=arrow_props)

    # External Entity: Marketing / Dashboard
    ent = patches.Rectangle((83, 68), 18, 9, facecolor="#0f172a", edgecolor="#475569", lw=1.5, zorder=3)
    ax.add_patch(ent)
    ax.text(92, 72.5, "Executive Dashboard\n& CRM Destinations", ha="center", va="center", color="white", fontsize=7, fontweight="bold", zorder=4)
    ax.annotate("", xy=(92, 68), xytext=(92, 56), arrowprops=arrow_props)

    plt.title("Figure 2.2: Level 1 Data Flow Diagram (End-to-End Analytics Pipeline)", fontsize=11, fontweight="bold", pad=12, color="#0f172a")
    plt.tight_layout()
    plt.savefig(img_dir / "fig_dfd_level_1.png")
    plt.close()

# -------------------------------------------------------------
# 3. Level 2 DFD (Decomposition of Process 2.0 and 4.0)
# -------------------------------------------------------------
def create_level_2_dfd():
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5), dpi=300)

    # Panel 1: Decomposition of Process 2.0 (Refinery)
    ax1.set_xlim(0, 100)
    ax1.set_ylim(0, 100)
    ax1.axis("off")
    ax1.set_title("Sub-Process 2.0: Data Refinery Pipeline", fontsize=10, fontweight="bold", pad=10, color="#4f46e5")

    p2_boxes = [
        ("2.1 Authenticate Entity\n(Drop Null CustomerID)", (50, 80), "#4338ca"),
        ("2.2 Isolate Reversals\n(Credit Notes 'C' & Q<=0)", (50, 50), "#4338ca"),
        ("2.3 Tukey IQR Outlier Capping\n(Quantity & Price Fences)", (50, 20), "#4338ca")
    ]
    for lbl, (x, y), col in p2_boxes:
        box = patches.FancyBboxPatch((x-22, y-8), 44, 16, boxstyle="round,pad=1.2", facecolor=col, edgecolor="#312e81", lw=1.5)
        ax1.add_patch(box)
        ax1.text(x, y, lbl, ha="center", va="center", color="white", fontsize=8, fontweight="bold")

    arrow_props = dict(arrowstyle="->", color="#334155", lw=1.8, mutation_scale=12)
    ax1.annotate("", xy=(50, 59), xytext=(50, 71), arrowprops=arrow_props)
    ax1.text(53, 65, "Identified Rows", fontsize=7, color="#1e293b")
    ax1.annotate("", xy=(50, 29), xytext=(50, 41), arrowprops=arrow_props)
    ax1.text(53, 35, "Valid Invoices", fontsize=7, color="#1e293b")

    # Side export for Reversals
    ax1.annotate("", xy=(88, 50), xytext=(73, 50), arrowprops=dict(arrowstyle="->", color="#dc2626", lw=1.8))
    ax1.text(90, 50, "Audit Ledger\n(713 rows)", fontsize=7, color="#dc2626", fontweight="bold", va="center")

    # Panel 2: Decomposition of Process 4.0 (ML Clustering & Scoring)
    ax2.set_xlim(0, 100)
    ax2.set_ylim(0, 100)
    ax2.axis("off")
    ax2.set_title("Sub-Process 4.0: Algorithmic & Quantile Engine", fontsize=10, fontweight="bold", pad=10, color="#4f46e5")

    p4_boxes = [
        ("4.1 Log1p Compression\n(Skew Mitigation)", (50, 85), "#059669"),
        ("4.2 StandardScaler\n(Z-Score Normalization)", (50, 60), "#059669"),
        ("4.3 K-Means Clustering\n(Elbow / Silhouette Tuning)", (50, 35), "#059669"),
        ("4.4 Quantile 1-5 Stratification\n(8 Operational Business Tiers)", (50, 10), "#059669"),
    ]
    for lbl, (x, y), col in p4_boxes:
        box = patches.FancyBboxPatch((x-22, y-7), 44, 14, boxstyle="round,pad=1.2", facecolor=col, edgecolor="#064e3b", lw=1.5)
        ax2.add_patch(box)
        ax2.text(x, y, lbl, ha="center", va="center", color="white", fontsize=8, fontweight="bold")

    ax2.annotate("", xy=(50, 68), xytext=(50, 77), arrowprops=arrow_props)
    ax2.annotate("", xy=(50, 43), xytext=(50, 52), arrowprops=arrow_props)
    ax2.annotate("", xy=(50, 18), xytext=(50, 27), arrowprops=arrow_props)

    plt.suptitle("Figure 2.3: Level 2 Data Flow Diagram (Refinery & ML Decomposition)", fontsize=11, fontweight="bold", y=0.98, color="#0f172a")
    plt.tight_layout()
    plt.savefig(img_dir / "fig_dfd_level_2.png")
    plt.close()

# -------------------------------------------------------------
# 4. UML Use Case Diagram
# -------------------------------------------------------------
def create_use_case_diagram():
    fig, ax = plt.subplots(figsize=(10, 6.5), dpi=300)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")

    # Boundary Box (System Boundary)
    sys_box = patches.Rectangle((25, 5), 50, 88, facecolor="#f8fafc", edgecolor="#475569", lw=2, linestyle="--")
    ax.add_patch(sys_box)
    ax.text(50, 89, "E-Commerce Customer Lifecycle & RFM System", ha="center", va="center", fontsize=9.5, fontweight="bold", color="#0f172a")

    # Actors (Stick Figure representations or stylized labels)
    def draw_actor(x, y, name):
        circle = patches.Circle((x, y+4), 2.5, facecolor="#0f172a", edgecolor="#0f172a")
        ax.add_patch(circle)
        ax.plot([x, x], [y+1.5, y-3.5], color="#0f172a", lw=2) # body
        ax.plot([x-3, x+3], [y, y], color="#0f172a", lw=2)     # arms
        ax.plot([x, x-2.5], [y-3.5, y-8], color="#0f172a", lw=2) # leg L
        ax.plot([x, x+2.5], [y-3.5, y-8], color="#0f172a", lw=2) # leg R
        ax.text(x, y-11, name, ha="center", va="top", fontsize=8, fontweight="bold", color="#0f172a")

    draw_actor(12, 75, "Data Engineer /\nETL Admin")
    draw_actor(12, 30, "Data Scientist /\nAnalytics Eng.")
    draw_actor(88, 75, "Marketing Manager /\nCRM Specialist")
    draw_actor(88, 30, "Commercial Executive /\nDirector")

    # Use Cases (Ellipses)
    use_cases = [
        (50, 80, "UC1: Ingest Transaction Logs\n& Enforce Typing"),
        (50, 66, "UC2: Isolate Reversals & Cap\nTukey IQR Outliers"),
        (50, 52, "UC3: Compute Recency,\nFrequency & Spend (ABT)"),
        (50, 38, "UC4: Train K-Means &\nAssign Quantile RFM Tiers"),
        (50, 24, "UC5: Inspect Revenue Share\n& High-Value At-Risk Whales"),
        (50, 11, "UC6: Export Target Audiences\nfor CRM Campaign Workflows")
    ]

    for x, y, label in use_cases:
        ellipse = patches.Ellipse((x, y), 24, 8.5, facecolor="#e0e7ff", edgecolor="#4f46e5", lw=1.6, zorder=3)
        ax.add_patch(ellipse)
        ax.text(x, y, label, ha="center", va="center", fontsize=7.5, color="#1e1b4b", fontweight="bold", zorder=4)

    # Association Lines
    line_props = dict(color="#64748b", lw=1.4)
    # Actor 1 connections (Data Engineer)
    ax.plot([14, 38], [72, 80], **line_props)
    ax.plot([14, 38], [70, 66], **line_props)

    # Actor 2 connections (Data Scientist)
    ax.plot([14, 38], [30, 52], **line_props)
    ax.plot([14, 38], [28, 38], **line_props)

    # Actor 3 connections (Marketing)
    ax.plot([86, 62], [72, 38], **line_props)
    ax.plot([86, 62], [70, 24], **line_props)
    ax.plot([86, 62], [68, 11], **line_props)

    # Actor 4 connections (Executive)
    ax.plot([86, 62], [30, 24], **line_props)

    plt.title("Figure 2.4: UML Use Case Diagram (System Actors & Functional Capabilities)", fontsize=11, fontweight="bold", pad=12, color="#0f172a")
    plt.tight_layout()
    plt.savefig(img_dir / "fig_use_case_diagram.png")
    plt.close()

if __name__ == "__main__":
    create_level_0_dfd()
    create_level_1_dfd()
    create_level_2_dfd()
    create_use_case_diagram()
    print("[SUCCESS] All 4 UML & DFD diagrams created successfully.")
