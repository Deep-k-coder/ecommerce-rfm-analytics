"""
Deep_Koshiya_Ecommerce_RFM_Segmentation.py
Complete Project Code: E-Commerce Customer Lifecycle & RFM Segmentation Dashboard
Author: Deep Koshiya
Date: September 2026
"""

from pathlib import Path
from src.clean_refinery_pipeline import RetailDataPipeline
from src.rfm_feature_engineering import generate_rfm_features
from src.segmentation_engine import CustomerSegmentationEngine


def main():
    base_dir = Path(__file__).resolve().parent
    raw_data_path = base_dir / "data" / "raw" / "online_retail.csv"
    processed_dir = base_dir / "data" / "processed"
    processed_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 75)
    print("E-COMMERCE CUSTOMER LIFECYCLE & RFM SEGMENTATION PIPELINE")
    print("Author: Deep Koshiya")
    print("=" * 75)

    # 1. Ingestion & Data Refinery
    print("\n[STEP 1] Running Data Cleaning & Quality Refinery...")
    pipeline = RetailDataPipeline(str(raw_data_path))
    clean_df, reversals_df = pipeline.execute()

    clean_df.to_csv(processed_dir / "fact_transactions.csv", index=False)
    reversals_df.to_csv(processed_dir / "fact_reversals_audit.csv", index=False)

    # 2. RFM Feature Engineering
    print("\n[STEP 2] Computing Vectorized Customer-Level RFM Metrics...")
    rfm_table = generate_rfm_features(clean_df)

    # 3. Machine Learning & Quantile Scoring
    print("\n[STEP 3] Performing Skew Correction, K-Means Clustering, & RFM Scoring...")
    engine = CustomerSegmentationEngine(rfm_table)
    engine.transform_features()

    print("Evaluating cluster inertia and silhouette diagnostics across k=[2..6]...")
    engine.evaluate_optimal_clusters(k_range=range(2, 7))

    print("Fitting production K-Means model with k=4...")
    engine.fit_kmeans(k=4)

    print("Applying quantile 1-5 scoring rules and mapping business segments...")
    final_dim_customer = engine.assign_rule_based_segments()

    final_dim_customer.to_csv(processed_dir / "dim_customer_rfm.csv", index=True)

    # 4. Summary Reporting
    total_rev = final_dim_customer["Monetary"].sum()
    total_cust = len(final_dim_customer)
    total_orders = final_dim_customer["Frequency"].sum()
    blended_aov = round(total_rev / total_orders, 2)

    print("\n" + "=" * 75)
    print("EXECUTIVE SUMMARY METRICS")
    print("=" * 75)
    print(f"Total Portfolio Revenue     : ${total_rev:,.2f}")
    print(f"Active Customer Accounts    : {total_cust:,}")
    print(f"Total Completed Invoices    : {total_orders:,}")
    print(f"Blended Portfolio AOV       : ${blended_aov:,.2f}")

    segment_summary = (
        final_dim_customer.groupby("Business_Segment")
        .agg(
            Users=("Recency", "count"),
            Spend=("Monetary", "sum"),
            Avg_R=("Recency", "mean"),
            Avg_F=("Frequency", "mean"),
            Avg_M=("Monetary", "mean")
        )
        .reset_index()
        .sort_values(by="Spend", ascending=False)
    )

    print("\n--- Business Segment Breakdown ---")
    for _, row in segment_summary.iterrows():
        rev_share = (row["Spend"] / total_rev) * 100
        user_share = (row["Users"] / total_cust) * 100
        print(
            f"• {row['Business_Segment']:<28} | Users: {row['Users']:>3} ({user_share:>4.1f}%) | "
            f"Spend: ${row['Spend']:>10,.2f} ({rev_share:>5.1f}%) | "
            f"Avg R: {row['Avg_R']:>5.1f}d | Avg F: {row['Avg_F']:>4.1f} | Avg M: ${row['Avg_M']:>6,.2f}"
        )

    print("\n[SUCCESS] Pipeline execution finished successfully.")


if __name__ == "__main__":
    main()
