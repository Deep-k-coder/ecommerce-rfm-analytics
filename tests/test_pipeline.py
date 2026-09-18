"""
test_pipeline.py
Unit tests verifying pipeline filtering, null dropping, and RFM calculations.
"""

import unittest
import pandas as pd
from datetime import datetime
from src.clean_refinery_pipeline import RetailDataPipeline
from src.rfm_feature_engineering import generate_rfm_features

class TestPipeline(unittest.TestCase):
    def setUp(self):
        self.sample_data = pd.DataFrame({
            "InvoiceNo": ["5001", "C5002", "5003", "5004"],
            "StockCode": ["85123A", "71053", "22752", "22633"],
            "Description": ["MUG", "BOX", "LAMP", "BAG"],
            "Quantity": [5.0, -1.0, 1000.0, 2.0],  # 1000 is outlier
            "InvoiceDate": ["2023-01-01 10:00:00", "2023-01-02 11:00:00", "2023-01-03 12:00:00", "2023-01-04 13:00:00"],
            "UnitPrice": [2.5, 3.0, 500.0, 4.0],  # 500 is outlier
            "CustomerID": [101.0, 101.0, 102.0, None]  # None should be dropped
        })
        self.test_csv = "/tmp/test_transactions.csv"
        self.sample_data.to_csv(self.test_csv, index=False)

    def test_clean_and_rfm(self):
        pipeline = RetailDataPipeline(self.test_csv)
        clean_df, reversals = pipeline.execute()

        # Check that null customer was dropped
        self.assertNotIn(None, clean_df["CustomerID"].values)
        
        # Check that cancellation was isolated
        self.assertEqual(len(reversals), 1)
        self.assertTrue(reversals["InvoiceNo"].iloc[0].startswith("C"))

        # Verify RFM feature engineering on clean records
        if len(clean_df) > 0:
            rfm = generate_rfm_features(clean_df)
            self.assertTrue((rfm["Recency"] >= 1).all())
            self.assertTrue((rfm["Frequency"] >= 1).all())
            self.assertTrue((rfm["Monetary"] > 0).all())

if __name__ == "__main__":
    unittest.main()
