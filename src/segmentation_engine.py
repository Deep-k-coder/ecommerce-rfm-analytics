"""
segmentation_engine.py
Algorithmic K-Means clustering alongside rule-based percentile RFM stratification.
"""

from typing import Dict, List
import numpy as np
import pandas as pd
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score
from sklearn.preprocessing import StandardScaler


class CustomerSegmentationEngine:
    """
    Executes transformation, unsupervised clustering, diagnostic evaluations,
    and rule-based quantile fallback mappings.
    """

    def __init__(self, rfm_df: pd.DataFrame):
        self.rfm_df = rfm_df.copy()
        self.scaled_features: np.ndarray = np.empty((0, 3))
        self.scaler = StandardScaler()
        self.kmeans_model: KMeans = None
        self.optimal_k: int = 4

    def transform_features(self) -> np.ndarray:
        """Applies Log1p compression followed by standard Z-score normalization."""
        rfm_features = self.rfm_df[["Recency", "Frequency", "Monetary"]].values

        # Log transformation (log1p handles small values smoothly)
        log_transformed = np.log1p(rfm_features)
        self.scaled_features = self.scaler.fit_transform(log_transformed)
        return self.scaled_features

    def evaluate_optimal_clusters(
        self, k_range: range = range(2, 7)
    ) -> Dict[str, List[float]]:
        """
        Computes Within-Cluster Sum of Squares (WCSS / Inertia) and Silhouette Scores
        across cluster counts to guide selection of optimal k.
        """
        metrics = {"k": [], "inertia": [], "silhouette": []}

        for k in k_range:
            km = KMeans(n_clusters=k, init="k-means++", n_init=10, random_state=42)
            labels = km.fit_predict(self.scaled_features)
            score = silhouette_score(self.scaled_features, labels)

            metrics["k"].append(k)
            metrics["inertia"].append(float(km.inertia_))
            metrics["silhouette"].append(float(score))
            print(f"[ML EVAL] k={k} -> Inertia: {km.inertia_:.1f}, Silhouette Score: {score:.4f}")

        return metrics

    def fit_kmeans(self, k: int = 4) -> pd.DataFrame:
        """Fits production K-Means model with fixed seed for determinism."""
        self.optimal_k = k
        self.kmeans_model = KMeans(
            n_clusters=k, init="k-means++", n_init=20, max_iter=300, random_state=42
        )
        self.rfm_df["Cluster_ID"] = self.kmeans_model.fit_predict(self.scaled_features)

        # Profile clusters to order them from lowest to highest commercial value
        cluster_order = (
            self.rfm_df.groupby("Cluster_ID")["Monetary"]
            .mean()
            .sort_values()
            .index
        )
        cluster_label_map = {
            old_id: f"Cluster_{new_id + 1}"
            for new_id, old_id in enumerate(cluster_order)
        }
        self.rfm_df["Cluster_Tier"] = self.rfm_df["Cluster_ID"].map(cluster_label_map)

        print(f"[ML CLUSTERING] Fitted K-Means with k={k}.")
        return self.rfm_df

    def assign_rule_based_segments(self) -> pd.DataFrame:
        """
        Generates industry-standard 1-5 RFM scores via quantile binning (pd.qcut)
        and maps them to business operating definitions.
        """
        df = self.rfm_df.copy()

        # Recency: Lower values are better (5 = most recent, 1 = longest absence)
        # Using rank(method='first') to handle identical bin boundary collisions
        df["R_Score"] = pd.qcut(
            df["Recency"].rank(method="first"), q=5, labels=[5, 4, 3, 2, 1]
        ).astype(int)

        # Frequency: Higher values are better
        df["F_Score"] = pd.qcut(
            df["Frequency"].rank(method="first"), q=5, labels=[1, 2, 3, 4, 5]
        ).astype(int)

        # Monetary: Higher values are better
        df["M_Score"] = pd.qcut(
            df["Monetary"].rank(method="first"), q=5, labels=[1, 2, 3, 4, 5]
        ).astype(int)

        # Composite RFM Score string
        df["RFM_Composite"] = (
            df["R_Score"].astype(str)
            + df["F_Score"].astype(str)
            + df["M_Score"].astype(str)
        )

        # Operational Segment Logic Engine
        def _map_segment(row: pd.Series) -> str:
            r, f = row["R_Score"], row["F_Score"]

            if r in [4, 5] and f in [4, 5]:
                return "Champions"
            elif r in [3, 4, 5] and f in [3, 4]:
                return "Loyal Customers"
            elif r in [4, 5] and f in [1, 2]:
                return "Recent Converts / Promising"
            elif r == 3 and f in [1, 2, 3]:
                return "Needs Attention"
            elif r in [1, 2] and f in [4, 5]:
                return "At Risk (High Value)"
            elif r in [1, 2] and f in [2, 3]:
                return "Hibernating"
            elif r in [1, 2] and f == 1:
                return "Lost / Inactive"
            return "Mid-Market Regulars"

        df["Business_Segment"] = df.apply(_map_segment, axis=1)

        # Actionable CRM Playbook mapping
        action_map = {
            "Champions": "VIP concierge access, early product drops, loyalty rewards",
            "Loyal Customers": "Category expansion, upsell incentives, referral programs",
            "Recent Converts / Promising": "Onboarding nurture stream, time-limited 2nd-order voucher",
            "Needs Attention": "Personalized re-activation email with product recommendations",
            "At Risk (High Value)": "Direct retention outreach, aggressive win-back discount",
            "Hibernating": "Low-cost email engagement test; purge if no response in 60d",
            "Lost / Inactive": "Exclude from paid acquisition remarketing to conserve budget",
            "Mid-Market Regulars": "Standard promo calendar, threshold-based free shipping"
        }
        df["Recommended_Action"] = df["Business_Segment"].map(action_map)

        self.rfm_df = df
        print(f"[SCORING] Assigned 1-5 RFM scores and business segments across {len(df):,} accounts.")
        return self.rfm_df
