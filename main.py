"""
main.py
End-to-end execution orchestrator for the E-Commerce Customer Lifecycle & RFM Segmentation project.
"""

import json
from pathlib import Path
from src.clean_refinery_pipeline import RetailDataPipeline
from src.rfm_feature_engineering import generate_rfm_features
from src.segmentation_engine import CustomerSegmentationEngine


def run_pipeline():
    base_dir = Path(__file__).resolve().parent
    raw_data_path = base_dir / "data" / "raw" / "online_retail.csv"
    processed_dir = base_dir / "data" / "processed"
    processed_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 70)
    print("PHASE 1: DATA CLEANING & REFINERY PIPELINE")
    print("=" * 70)
    pipeline = RetailDataPipeline(str(raw_data_path))
    clean_df, reversals_df = pipeline.execute()

    fact_tx_path = processed_dir / "fact_transactions.csv"
    reversals_path = processed_dir / "fact_reversals_audit.csv"
    clean_df.to_csv(fact_tx_path, index=False)
    reversals_df.to_csv(reversals_path, index=False)
    print(f"[EXPORT] Saved clean transactions to {fact_tx_path}")
    print(f"[EXPORT] Saved audit reversals to {reversals_path}")

    print("\n" + "=" * 70)
    print("PHASE 2: FEATURE ENGINEERING (RFM METRICS)")
    print("=" * 70)
    rfm_table = generate_rfm_features(clean_df)

    print("\n" + "=" * 70)
    print("PHASE 3: ALGORITHMIC SEGMENTATION & QUANTILE SCORING")
    print("=" * 70)
    engine = CustomerSegmentationEngine(rfm_table)
    engine.transform_features()

    print("\n--- Evaluating Cluster Quality Across k=[2..6] ---")
    eval_metrics = engine.evaluate_optimal_clusters(k_range=range(2, 7))

    # Fit K-Means with k=4
    engine.fit_kmeans(k=4)

    # Assign quantile rules & business segments
    final_dim_customer = engine.assign_rule_based_segments()

    dim_cust_path = processed_dir / "dim_customer_rfm.csv"
    final_dim_customer.to_csv(dim_cust_path, index=True)
    print(f"[EXPORT] Saved dimensional customer table to {dim_cust_path}")

    # Generate executive summary statistics for reporting & dashboard
    total_rev = float(final_dim_customer["Monetary"].sum())
    total_cust = len(final_dim_customer)
    total_orders = int(final_dim_customer["Frequency"].sum())
    overall_aov = round(total_rev / total_orders, 2)

    segment_summary = (
        final_dim_customer.groupby("Business_Segment")
        .agg(
            Customer_Count=("Recency", "count"),
            Total_Spend=("Monetary", "sum"),
            Avg_Recency=("Recency", "mean"),
            Avg_Frequency=("Frequency", "mean"),
            Avg_Spend=("Monetary", "mean"),
            Avg_AOV=("AOV", "mean")
        )
        .reset_index()
    )
    segment_summary["Rev_Contribution_Pct"] = (
        (segment_summary["Total_Spend"] / total_rev) * 100
    ).round(2)
    segment_summary["Customer_Pct"] = (
        (segment_summary["Customer_Count"] / total_cust) * 100
    ).round(2)
    segment_summary = segment_summary.sort_values(by="Total_Spend", ascending=False)

    print("\n" + "=" * 70)
    print("EXECUTIVE SEGMENT BREAKDOWN SUMMARY")
    print("=" * 70)
    for _, row in segment_summary.iterrows():
        print(
            f"• {row['Business_Segment']:<28} | "
            f"Users: {row['Customer_Count']:>4} ({row['Customer_Pct']:>5.1f}%) | "
            f"Spend: ${row['Total_Spend']:>10,.2f} ({row['Rev_Contribution_Pct']:>5.1f}%) | "
            f"Avg R: {row['Avg_Recency']:>5.1f}d | "
            f"Avg F: {row['Avg_Frequency']:>4.1f} | "
            f"Avg M: ${row['Avg_Spend']:>7,.2f}"
        )

    print("\n" + "=" * 70)
    print("KEY PORTFOLIO HEALTH METRICS")
    print("=" * 70)
    print(f"Total Portfolio Revenue     : ${total_rev:,.2f}")
    print(f"Active Customer Base        : {total_cust:,}")
    print(f"Total Customer Invoices     : {total_orders:,}")
    print(f"Blended Portfolio AOV       : ${overall_aov:,.2f}")
    
    champions_rev = segment_summary[segment_summary["Business_Segment"] == "Champions"]["Rev_Contribution_Pct"].values
    champions_pct = champions_rev[0] if len(champions_rev) > 0 else 0
    at_risk_spend = segment_summary[segment_summary["Business_Segment"].str.contains("At Risk")]["Total_Spend"].values
    at_risk_val = at_risk_spend[0] if len(at_risk_spend) > 0 else 0
    
    print(f"Champions Revenue Share     : {champions_pct:.1f}%")
    print(f"At-Risk Capital Exposure    : ${at_risk_val:,.2f}")
    print("=" * 70)

    # Save summary metadata for the interactive HTML dashboard
    dashboard_payload = {
        "kpis": {
            "total_revenue": total_rev,
            "active_customers": total_cust,
            "total_orders": total_orders,
            "blended_aov": overall_aov,
            "champions_rev_share": champions_pct,
            "at_risk_capital": at_risk_val,
        },
        "cluster_eval": eval_metrics,
        "segments": segment_summary.to_dict(orient="records"),
        "customers_sample": final_dim_customer.reset_index().head(100).to_dict(orient="records")
    }
    with open(base_dir / "dashboard" / "summary_data.json", "w") as f:
        json.dump(dashboard_payload, f, indent=2)

    print("\n[SUCCESS] Pipeline orchestration successfully completed!")


if __name__ == "__main__":
    run_pipeline()
