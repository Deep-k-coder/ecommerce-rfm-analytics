"""
generate_report_assets.py
Generates high-resolution chart images to embed within the Word project report.
"""

from pathlib import Path
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

def generate_assets():
    base_dir = Path(__file__).resolve().parent
    img_dir = base_dir / "report_images"
    img_dir.mkdir(parents=True, exist_ok=True)
    
    dim_cust_path = base_dir / "data" / "processed" / "dim_customer_rfm.csv"
    df = pd.read_csv(dim_cust_path)
    
    # 1. Revenue Share & Spend Chart
    seg_summary = df.groupby("Business_Segment").agg(
        Spend=("Monetary", "sum"),
        Users=("Recency", "count"),
        Avg_Spend=("Monetary", "mean")
    ).reset_index().sort_values(by="Spend", ascending=False)
    
    plt.figure(figsize=(10, 5), dpi=300)
    plt.pie(
        seg_summary["Spend"],
        labels=seg_summary["Business_Segment"],
        autopct="%1.1f%%",
        startangle=140,
        colors=["#10b981", "#6366f1", "#64748b", "#38bdf8", "#ef4444", "#475569", "#f59e0b", "#a855f7"]
    )
    plt.title("Revenue Contribution Share by Customer Segment", fontsize=13, fontweight="bold", pad=15)
    plt.tight_layout()
    pie_path = img_dir / "fig_revenue_share.png"
    plt.savefig(pie_path)
    plt.close()
    
    # 2. Elbow & Silhouette Curves
    k_vals = [2, 3, 4, 5, 6]
    inertia_vals = [757.9, 545.3, 430.5, 338.4, 288.8]
    sil_vals = [0.4575, 0.3625, 0.3817, 0.3505, 0.3536]
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 4.5), dpi=300)
    ax1.plot(k_vals, inertia_vals, marker="o", color="#4f46e5", linewidth=2.5, markersize=7)
    ax1.set_title("Elbow Method: Inertia (WCSS)", fontsize=11, fontweight="bold")
    ax1.set_xlabel("Clusters (k)")
    ax1.set_ylabel("Within-Cluster Sum of Squares")
    ax1.grid(True, linestyle="--", alpha=0.5)
    
    colors = ["#10b981" if k == 4 else "#94a3b8" for k in k_vals]
    ax2.bar(k_vals, sil_vals, color=colors, width=0.55)
    ax2.set_title("Silhouette Coefficient (Optimal k=4)", fontsize=11, fontweight="bold")
    ax2.set_xlabel("Clusters (k)")
    ax2.set_ylabel("Silhouette Score")
    ax2.set_ylim(0.2, 0.55)
    ax2.grid(True, linestyle="--", alpha=0.5)
    
    plt.tight_layout()
    diag_path = img_dir / "fig_kmeans_diagnostics.png"
    plt.savefig(diag_path)
    plt.close()

    # 3. Spatial Scatterplot
    plt.figure(figsize=(10, 5.5), dpi=300)
    segment_colors = {
        "Champions": "#10b981",
        "Loyal Customers": "#6366f1",
        "Hibernating": "#64748b",
        "Mid-Market Regulars": "#38bdf8",
        "At Risk (High Value)": "#ef4444",
        "Lost / Inactive": "#475569",
        "Needs Attention": "#f59e0b",
        "Recent Converts / Promising": "#a855f7"
    }
    for seg, grp in df.groupby("Business_Segment"):
        plt.scatter(
            grp["Recency"],
            grp["Monetary"],
            label=seg,
            alpha=0.7,
            s=grp["Frequency"] * 5 + 15,
            c=segment_colors.get(seg, "#3b82f6"),
            edgecolors="none"
        )
    plt.title("Spatial Customer Migration Space (Recency vs. Spend)", fontsize=12, fontweight="bold")
    plt.xlabel("Recency (Days Inactive)")
    plt.ylabel("Net Cumulative Spend ($)")
    plt.legend(bbox_to_anchor=(1.02, 1), loc="upper left", frameon=True, fontsize=8)
    plt.grid(True, linestyle="--", alpha=0.5)
    plt.tight_layout()
    scatter_path = img_dir / "fig_spatial_scatter.png"
    plt.savefig(scatter_path)
    plt.close()
    
    print("[ASSETS] Generated high-res report chart assets:")
    print(f"  • {pie_path}")
    print(f"  • {diag_path}")
    print(f"  • {scatter_path}")

if __name__ == "__main__":
    generate_assets()
