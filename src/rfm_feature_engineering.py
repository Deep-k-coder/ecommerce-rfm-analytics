"""
rfm_feature_engineering.py
Vectorized aggregation of transactional streams into an RFM analytical base table (ABT).
"""

import pandas as pd


def generate_rfm_features(clean_df: pd.DataFrame) -> pd.DataFrame:
    """
    Transforms row-level transactional records into a customer-level analytical feature store.

    Parameters:
        clean_df: Filtered transactional DataFrame with CustomerID, InvoiceDate,
                  InvoiceNo, and LineItemTotal.

    Returns:
        pd.DataFrame indexed by CustomerID with columns: [Recency, Frequency, Monetary, AOV]
    """
    # Establish snapshot boundary (1 day past the maximum observed transaction)
    snapshot_date = clean_df["InvoiceDate"].max() + pd.Timedelta(days=1)

    # Perform a single-pass optimized group aggregation
    rfm_abt = clean_df.groupby("CustomerID").agg(
        Recency=(
            "InvoiceDate",
            lambda dates: int((snapshot_date - dates.max()).days),
        ),
        Frequency=("InvoiceNo", "nunique"),
        Monetary=("LineItemTotal", "sum"),
    )

    # Secondary derived behavioral metrics
    rfm_abt["AOV"] = (rfm_abt["Monetary"] / rfm_abt["Frequency"]).round(2)
    rfm_abt["Monetary"] = rfm_abt["Monetary"].round(2)

    # Filter any potential zero-spend artifacts
    rfm_abt = rfm_abt[rfm_abt["Monetary"] > 0].copy()

    print(f"[FEATURE ENG] Built RFM ABT for {len(rfm_abt):,} distinct active customers.")
    return rfm_abt
