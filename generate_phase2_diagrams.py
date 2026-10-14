"""
generate_phase2_diagrams.py
Generates publication-quality architectural diagrams for Second Phase Work (Mid-Term):
- Figure 2.5: UML Class Diagram (Domain Model & Pipeline Object Architecture)
- Figure 2.6: UML Sequence Diagram (End-to-End Processing & Activation Workflow)
- Figure 4.1: Executive Screen Layouts with Validation Architecture
"""

import matplotlib.pyplot as plt
import matplotlib.patches as patches
from pathlib import Path

img_dir = Path("/Users/deep/.gemini/antigravity/scratch/ecommerce_rfm_analytics/report_images")
img_dir.mkdir(parents=True, exist_ok=True)

# -------------------------------------------------------------
# 1. UML Class Diagram (Section 2.3)
# -------------------------------------------------------------
def create_class_diagram():
    fig, ax = plt.subplots(figsize=(12, 7.5), dpi=300)
    ax.set_xlim(0, 120)
    ax.set_ylim(0, 100)
    ax.axis("off")

    def draw_class(x, y, w, h, name, attrs, methods, header_col="#1e1b4b"):
        # Header Box
        hdr_h = 4.5
        hdr = patches.Rectangle((x, y + h - hdr_h), w, hdr_h, facecolor=header_col, edgecolor="#0f172a", lw=1.5, zorder=3)
        ax.add_patch(hdr)
        ax.text(x + w/2, y + h - hdr_h/2, name, ha="center", va="center", color="white", fontsize=8, fontweight="bold", zorder=4)

        # Body Box
        body = patches.Rectangle((x, y), w, h - hdr_h, facecolor="#f8fafc", edgecolor="#0f172a", lw=1.5, zorder=2)
        ax.add_patch(body)

        # Attributes
        curr_y = y + h - hdr_h - 2.5
        for attr in attrs:
            ax.text(x + 1.2, curr_y, attr, ha="left", va="center", color="#0f172a", fontsize=6.8, fontfamily="monospace", zorder=4)
            curr_y -= 2.2

        # Divider
        ax.plot([x, x + w], [curr_y - 0.5, curr_y - 0.5], color="#94a3b8", lw=1, zorder=3)
        curr_y -= 2.5

        # Methods
        for m in methods:
            ax.text(x + 1.2, curr_y, m, ha="left", va="center", color="#0f172a", fontsize=6.8, fontfamily="monospace", zorder=4)
            curr_y -= 2.2

    # Classes definition
    # 1. TransactionLog (Model Entity)
    draw_class(
        x=5, y=62, w=32, h=32,
        name="TransactionRecord",
        attrs=[
            "+ invoice_no: String",
            "+ stock_code: String",
            "+ description: String",
            "+ quantity: Integer",
            "+ invoice_date: DateTime",
            "+ unit_price: Float",
            "+ customer_id: String"
        ],
        methods=[
            "+ is_valid_session(): Boolean",
            "+ is_cancellation(): Boolean",
            "+ get_gross_amount(): Float"
        ],
        header_col="#7a003c" # IAR Maroon
    )

    # 2. CleanRefineryPipeline
    draw_class(
        x=44, y=62, w=34, h=32,
        name="CleanRefineryPipeline",
        attrs=[
            "- raw_path: Path",
            "- output_dir: Path",
            "- iqr_multiplier: Float = 1.5",
            "+ valid_df: DataFrame",
            "+ audit_reversals: DataFrame"
        ],
        methods=[
            "+ load_raw_transactions(): DataFrame",
            "+ enforce_schema_typing(): DataFrame",
            "+ filter_unauthenticated_nulls(): Void",
            "+ isolate_reverse_logistics(): Void",
            "+ cap_tukey_outliers(): DataFrame",
            "+ execute_refinery(): Tuple"
        ],
        header_col="#1e1b4b"
    )

    # 3. RFMFeatureEngineering
    draw_class(
        x=83, y=62, w=32, h=32,
        name="RFMFeatureEngineer",
        attrs=[
            "- snapshot_date: DateTime",
            "+ rfm_table: DataFrame",
            "+ customer_count: Integer"
        ],
        methods=[
            "+ compute_snapshot_boundary(): Date",
            "+ aggregate_recency(): Series",
            "+ aggregate_frequency(): Series",
            "+ aggregate_monetary(): Series",
            "+ derive_aov_metric(): Series",
            "+ build_analytical_base_table(): DF"
        ],
        header_col="#1e1b4b"
    )

    # 4. KMeansSegmentationEngine
    draw_class(
        x=5, y=10, w=34, h=44,
        name="KMeansSegmentationEngine",
        attrs=[
            "- k_clusters: Integer = 4",
            "- random_state: Integer = 42",
            "- scaler: StandardScaler",
            "- model: KMeans",
            "+ inertia_scores: List<Float>",
            "+ silhouette_scores: List<Float>",
            "+ cluster_labels: Series"
        ],
        methods=[
            "+ apply_log1p_compression(): Array",
            "+ standardize_features(): Array",
            "+ evaluate_elbow_diagnostics(): Dict",
            "+ fit_predict_clusters(): Series",
            "+ compute_cluster_centroids(): DF",
            "+ export_model_diagnostics(): Void"
        ],
        header_col="#047857" # Green
    )

    # 5. QuantileStratifier
    draw_class(
        x=44, y=10, w=34, h=44,
        name="QuantileStratifier",
        attrs=[
            "- num_quantiles: Integer = 5",
            "- segment_map: Dict<Tuple, String>",
            "+ rfm_scores: DataFrame",
            "+ business_segments: Series"
        ],
        methods=[
            "+ calculate_recency_score(): Series",
            "+ calculate_frequency_score(): Series",
            "+ calculate_monetary_score(): Series",
            "+ synthesize_rfm_string(): Series",
            "+ map_8_operating_cohorts(): Series",
            "+ compute_segment_summary(): DF"
        ],
        header_col="#047857"
    )

    # 6. StarSchemaMartExporter
    draw_class(
        x=83, y=10, w=32, h=44,
        name="StarSchemaMartExporter",
        attrs=[
            "- export_dir: Path",
            "+ fact_transactions: DataFrame",
            "+ fact_reversals: DataFrame",
            "+ dim_customers: DataFrame",
            "+ dim_products: DataFrame"
        ],
        methods=[
            "+ build_fact_transactions(): Void",
            "+ build_fact_reversals(): Void",
            "+ build_dim_customers(): Void",
            "+ build_dim_products(): Void",
            "+ export_csv_marts(): Void",
            "+ generate_summary_json(): Void"
        ],
        header_col="#7a003c"
    )

    # Relationship Arrows
    arrow_props = dict(arrowstyle="->", color="#334155", lw=1.6)
    
    # Pipeline depends on TransactionRecord
    ax.annotate("", xy=(37, 78), xytext=(44, 78), arrowprops=arrow_props)
    ax.text(40.5, 80, "uses", fontsize=7, color="#475569", ha="center")

    # Pipeline feeds RFM Engineer
    ax.annotate("", xy=(83, 78), xytext=(78, 78), arrowprops=arrow_props)
    ax.text(80.5, 80, "feeds", fontsize=7, color="#475569", ha="center")

    # RFM Engineer feeds KMeans & Quantile
    ax.annotate("", xy=(22, 54), xytext=(90, 62), arrowprops=dict(arrowstyle="->", color="#4338ca", lw=1.5, linestyle="--"))
    ax.annotate("", xy=(61, 54), xytext=(99, 62), arrowprops=dict(arrowstyle="->", color="#4338ca", lw=1.5, linestyle="--"))

    # KMeans & Quantile feed StarSchema
    ax.annotate("", xy=(83, 30), xytext=(78, 30), arrowprops=arrow_props)
    ax.text(80.5, 32, "exports", fontsize=7, color="#475569", ha="center")

    plt.title("Figure 2.5: UML Class Diagram (Domain Model & Analytical Pipeline Classes)", fontsize=11, fontweight="bold", pad=12, color="#0f172a")
    plt.tight_layout()
    plt.savefig(img_dir / "fig_class_diagram.png")
    plt.close()

# -------------------------------------------------------------
# 2. UML Sequence Diagram (Section 2.4)
# -------------------------------------------------------------
def create_sequence_diagram():
    fig, ax = plt.subplots(figsize=(12, 7.5), dpi=300)
    ax.set_xlim(0, 110)
    ax.set_ylim(0, 100)
    ax.axis("off")

    lifelines = [
        ("Data Engineer /\nScheduler", 10),
        ("CleanRefinery\nPipeline", 30),
        ("RFMFeature\nEngineer", 50),
        ("KMeans\nEngine", 70),
        ("StarSchema\nDataMart", 90),
        ("Executive\nDashboard", 104)
    ]

    # Draw Lifeline Headers & Vertical Dashed Lines
    for name, x in lifelines:
        box = patches.FancyBboxPatch((x-8, 90), 16, 7.5, boxstyle="round,pad=0.8", facecolor="#1e1b4b", edgecolor="#0f172a", lw=1.5, zorder=3)
        ax.add_patch(box)
        ax.text(x, 93.75, name, ha="center", va="center", color="white", fontsize=7, fontweight="bold", zorder=4)
        ax.plot([x, x], [89, 8], color="#cbd5e1", lw=1.4, linestyle="--", zorder=1)

    # Sequence Messages
    msgs = [
        (85, 10, 30, "1: trigger_pipeline(raw_csv)", "#1e293b", False),
        (78, 30, 30, "2: filter_nulls_and_reversals()", "#7a003c", True), # self-call
        (71, 30, 30, "3: apply_tukey_iqr_capping()", "#7a003c", True),
        (64, 30, 50, "4: generate_rfm_metrics(clean_df)", "#1e293b", False),
        (57, 50, 50, "5: compute_snapshot_boundary()", "#059669", True),
        (50, 50, 70, "6: fit_predict_clusters(rfm_abt)", "#1e293b", False),
        (43, 70, 70, "7: log1p_scaling & k-means(k=4)", "#059669", True),
        (36, 70, 90, "8: export_dim_customers(segments)", "#1e293b", False),
        (29, 30, 90, "9: export_fact_tables(transactions, reversals)", "#1e293b", False),
        (22, 90, 104, "10: refresh_summary_mart()", "#2563eb", False),
        (15, 104, 10, "11: render_executive_kpis & charts", "#2563eb", False)
    ]

    for y, x_src, x_dst, label, col, is_self in msgs:
        if is_self:
            # Self Call Loop
            ax.plot([x_src, x_src + 6, x_src + 6, x_src], [y, y, y - 4, y - 4], color=col, lw=1.5, zorder=3)
            ax.annotate("", xy=(x_src, y - 4), xytext=(x_src + 2, y - 4), arrowprops=dict(arrowstyle="->", color=col, lw=1.5))
            ax.text(x_src + 7.5, y - 2, label, ha="left", va="center", color=col, fontsize=7.2, fontweight="bold", zorder=4)
        else:
            # Standard Message Arrow
            ax.annotate("", xy=(x_dst, y), xytext=(x_src, y), arrowprops=dict(arrowstyle="->", color=col, lw=1.6))
            mid_x = (x_src + x_dst) / 2
            ax.text(mid_x, y + 1.8, label, ha="center", va="bottom", color=col, fontsize=7.2, fontweight="bold", zorder=4)

    plt.title("Figure 2.6: UML Sequence Diagram (End-to-End Processing & Activation Workflow)", fontsize=11, fontweight="bold", pad=12, color="#0f172a")
    plt.tight_layout()
    plt.savefig(img_dir / "fig_sequence_diagram.png")
    plt.close()

# -------------------------------------------------------------
# 3. Screen Layouts with Validations (Section 4.1)
# -------------------------------------------------------------
def create_screen_layouts_visual():
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(12, 8), dpi=300)

    # Panel 1: Executive KPI Scorecards Screen
    ax1.set_xlim(0, 100)
    ax1.set_ylim(0, 100)
    ax1.axis("off")
    ax1.set_title("Screen 1: Executive KPI Overview & Health Scorecards", fontsize=9, fontweight="bold", color="#1e1b4b", pad=8)
    
    # Container
    ax1.add_patch(patches.Rectangle((2, 2), 96, 94, facecolor="#0f172a", edgecolor="#334155", lw=1.5))
    # KPI Cards
    kpis = [
        ("Gross Revenue", "$65,049.20", (6, 58), "#1e293b", "#10b981"),
        ("Active Cohort", "256 Accts", (29, 58), "#1e293b", "#3b82f6"),
        ("Avg Order Value", "$254.10", (52, 58), "#1e293b", "#f59e0b"),
        ("At-Risk Revenue", "7.3%", (75, 58), "#1e293b", "#ef4444"),
    ]
    for lbl, val, (x, y), bg, tc in kpis:
        ax1.add_patch(patches.FancyBboxPatch((x, y), 20, 26, boxstyle="round,pad=1", facecolor=bg, edgecolor="#475569", lw=1))
        ax1.text(x+10, y+18, lbl, ha="center", va="center", color="#94a3b8", fontsize=6.5, fontweight="bold")
        ax1.text(x+10, y+8, val, ha="center", va="center", color=tc, fontsize=9.5, fontweight="bold")

    # Validation Callout Box
    ax1.add_patch(patches.FancyBboxPatch((6, 8), 88, 38, boxstyle="round,pad=1", facecolor="#1e293b", edgecolor="#7a003c", lw=1.5))
    ax1.text(10, 38, "VALIDATION RULES (VAL-01):", color="#f43f5e", fontsize=7.5, fontweight="bold")
    ax1.text(10, 28, "• Monetary totals validated against non-negative constraint (Monetary >= 0.00)", color="#e2e8f0", fontsize=6.5)
    ax1.text(10, 20, "• Active customer count strictly cross-checked with unique non-null CustomerIDs", color="#e2e8f0", fontsize=6.5)
    ax1.text(10, 12, "• At-Risk percentage evaluated strictly as: (Sum(Spend_AtRisk) / Total_Spend) * 100", color="#e2e8f0", fontsize=6.5)

    # Panel 2: Interactive RFM Scatterplot
    ax2.set_xlim(0, 100)
    ax2.set_ylim(0, 100)
    ax2.axis("off")
    ax2.set_title("Screen 2: Multi-Dimensional RFM Behavioral Scatterplot", fontsize=9, fontweight="bold", color="#1e1b4b", pad=8)
    ax2.add_patch(patches.Rectangle((2, 2), 96, 94, facecolor="#0f172a", edgecolor="#334155", lw=1.5))

    # Scatter simulation
    import numpy as np
    np.random.seed(42)
    xs = np.random.uniform(15, 85, 30)
    ys = np.random.uniform(45, 88, 30)
    ax2.scatter(xs, ys, s=50, c="#38bdf8", alpha=0.8, edgecolors="white", lw=0.5)

    ax2.add_patch(patches.FancyBboxPatch((6, 8), 88, 30, boxstyle="round,pad=1", facecolor="#1e293b", edgecolor="#7a003c", lw=1.5))
    ax2.text(10, 31, "VALIDATION RULES (VAL-02):", color="#f43f5e", fontsize=7.5, fontweight="bold")
    ax2.text(10, 22, "• Axis coordinates bounded to: Recency [1..365], Frequency [1..50], Spend [0..10k]", color="#e2e8f0", fontsize=6.5)
    ax2.text(10, 13, "• Tooltip rendering validates customer ID existence prior to canvas hover event", color="#e2e8f0", fontsize=6.5)

    # Panel 3: Customer Cohort Activation Ledger
    ax3.set_xlim(0, 100)
    ax3.set_ylim(0, 100)
    ax3.axis("off")
    ax3.set_title("Screen 3: Customer Cohort Drilldown & Activation Grid", fontsize=9, fontweight="bold", color="#1e1b4b", pad=8)
    ax3.add_patch(patches.Rectangle((2, 2), 96, 94, facecolor="#0f172a", edgecolor="#334155", lw=1.5))

    # Table Simulation
    ax3.add_patch(patches.Rectangle((6, 48), 88, 12, facecolor="#7a003c"))
    ax3.text(12, 54, "CustomerID", color="white", fontsize=6.5, fontweight="bold")
    ax3.text(32, 54, "Recency", color="white", fontsize=6.5, fontweight="bold")
    ax3.text(48, 54, "Frequency", color="white", fontsize=6.5, fontweight="bold")
    ax3.text(64, 54, "Monetary", color="white", fontsize=6.5, fontweight="bold")
    ax3.text(82, 54, "Segment Tier", color="white", fontsize=6.5, fontweight="bold")

    rows = [
        ("CUST-14911", "1 day", "201 trips", "$4,210.50", "Champions", 40),
        ("CUST-17841", "3 days", "112 trips", "$1,890.20", "Champions", 32),
        ("CUST-12346", "325 days", "1 trip", "$7,718.30", "At Risk Whales", 24)
    ]
    for cid, rec, freq, mon, seg, y_pos in rows:
        ax3.text(12, y_pos, cid, color="#cbd5e1", fontsize=6)
        ax3.text(32, y_pos, rec, color="#cbd5e1", fontsize=6)
        ax3.text(48, y_pos, freq, color="#cbd5e1", fontsize=6)
        ax3.text(64, y_pos, mon, color="#cbd5e1", fontsize=6)
        ax3.text(82, y_pos, seg, color="#34d399", fontsize=6, fontweight="bold")

    ax3.add_patch(patches.FancyBboxPatch((6, 5), 88, 16, boxstyle="round,pad=0.8", facecolor="#1e293b", edgecolor="#7a003c", lw=1))
    ax3.text(8, 15, "VAL-03: Strict client-side search query sanitation & regex checking", color="#f43f5e", fontsize=6.5, fontweight="bold")
    ax3.text(8, 8, "• Pagination bound enforcement (Page >= 1 and Page <= MaxPage)", color="#e2e8f0", fontsize=6)

    # Panel 4: Reversals Audit Ledger Screen
    ax4.set_xlim(0, 100)
    ax4.set_ylim(0, 100)
    ax4.axis("off")
    ax4.set_title("Screen 4: Reverse Logistics & Return Audit Ledger", fontsize=9, fontweight="bold", color="#1e1b4b", pad=8)
    ax4.add_patch(patches.Rectangle((2, 2), 96, 94, facecolor="#0f172a", edgecolor="#334155", lw=1.5))

    ax4.add_patch(patches.Rectangle((6, 48), 88, 12, facecolor="#1e293b", edgecolor="#dc2626", lw=1.2))
    ax4.text(12, 54, "InvoiceNo", color="white", fontsize=6.5, fontweight="bold")
    ax4.text(32, 54, "StockCode", color="white", fontsize=6.5, fontweight="bold")
    ax4.text(48, 54, "Qty (< 0)", color="#f87171", fontsize=6.5, fontweight="bold")
    ax4.text(66, 54, "UnitPrice", color="white", fontsize=6.5, fontweight="bold")
    ax4.text(82, 54, "CustomerID", color="white", fontsize=6.5, fontweight="bold")

    rev_rows = [
        ("C541433", "23166", "-74,215", "$1.04", "12346", 40),
        ("C536391", "22556", "-12", "$1.65", "17548", 32),
        ("C536383", "35004C", "-1", "$4.65", "15311", 24)
    ]
    for inv, sc, q, up, cid, y_pos in rev_rows:
        ax4.text(12, y_pos, inv, color="#fca5a5", fontsize=6)
        ax4.text(32, y_pos, sc, color="#cbd5e1", fontsize=6)
        ax4.text(48, y_pos, q, color="#f87171", fontsize=6, fontweight="bold")
        ax4.text(66, y_pos, up, color="#cbd5e1", fontsize=6)
        ax4.text(82, y_pos, cid, color="#cbd5e1", fontsize=6)

    ax4.add_patch(patches.FancyBboxPatch((6, 5), 88, 16, boxstyle="round,pad=0.8", facecolor="#1e293b", edgecolor="#7a003c", lw=1))
    ax4.text(8, 15, "VAL-04: Cancellation code regex constraint: InvoiceNo.startswith('C')", color="#f43f5e", fontsize=6.5, fontweight="bold")
    ax4.text(8, 8, "• Negative quantity validation (Quantity < 0) before insertion into audit ledger", color="#e2e8f0", fontsize=6)

    plt.suptitle("Figure 4.1: Implementation Screen Layouts with Operational Input Validations", fontsize=11, fontweight="bold", y=0.98, color="#0f172a")
    plt.tight_layout()
    plt.savefig(img_dir / "fig_screen_layouts.png")
    plt.close()

if __name__ == "__main__":
    create_class_diagram()
    create_sequence_diagram()
    create_screen_layouts_visual()
    print("[SUCCESS] All Phase 2 diagrams created successfully.")
