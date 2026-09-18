"""
clean_refinery_pipeline.py
Production-grade data ingestion, validation, and refinery module.
"""

from typing import Tuple
import numpy as np
import pandas as pd


class RetailDataPipeline:
    """
    Refinery pipeline for transactional e-commerce event streams.
    Enforces schema typing, purges unauthenticated checkouts, isolates reversals,
    and caps non-representative statistical outliers via Tukey's IQR fences.
    """

    def __init__(self, raw_filepath: str):
        self.raw_filepath = raw_filepath
        self.df: pd.DataFrame = pd.DataFrame()
        self.reversals: pd.DataFrame = pd.DataFrame()

    def load_data(self) -> "RetailDataPipeline":
        """Loads transactional data handling encoding anomalies and date parsing."""
        dtype_spec = {
            "InvoiceNo": "str",
            "StockCode": "str",
            "Description": "str",
            "Quantity": "float64",
            "UnitPrice": "float64",
            "CustomerID": "float64",
        }

        try:
            self.df = pd.read_csv(
                self.raw_filepath,
                dtype=dtype_spec,
                parse_dates=["InvoiceDate"],
                encoding="utf-8",
            )
        except UnicodeDecodeError:
            self.df = pd.read_csv(
                self.raw_filepath,
                dtype=dtype_spec,
                parse_dates=["InvoiceDate"],
                encoding="ISO-8859-1",
            )

        print(f"[REFINERY] Ingested {len(self.df):,} raw transactional records.")
        return self

    def clean_text_and_nulls(self) -> "RetailDataPipeline":
        """Standardizes text fields and purges missing entity identifiers."""
        self.df["InvoiceNo"] = self.df["InvoiceNo"].astype(str).str.strip()
        self.df["StockCode"] = self.df["StockCode"].astype(str).str.strip()
        self.df["Description"] = (
            self.df["Description"].astype(str).str.strip().str.upper()
        )

        initial_count = len(self.df)
        # Drop records lacking CustomerID (guest/unregistered checkouts)
        self.df = self.df.dropna(subset=["CustomerID"]).copy()
        self.df["CustomerID"] = self.df["CustomerID"].astype(int).astype(str)

        dropped_nulls = initial_count - len(self.df)
        print(f"[REFINERY] Dropped {dropped_nulls:,} rows lacking valid CustomerID.")
        return self

    def isolate_reversals_and_cancellations(self) -> "RetailDataPipeline":
        """
        Splits credit notes, cancellations, and stock write-offs into a segregated ledger.
        Invoices prefixed with 'C' or carrying non-positive quantities/prices are separated.
        """
        cancellation_mask = (
            self.df["InvoiceNo"].str.startswith("C", na=False)
            | (self.df["Quantity"] <= 0)
            | (self.df["UnitPrice"] <= 0)
        )

        self.reversals = self.df[cancellation_mask].copy()
        self.df = self.df[~cancellation_mask].copy()

        print(f"[REFINERY] Isolated {len(self.reversals):,} reversal/cancellation entries.")
        print(f"[REFINERY] Retained {len(self.df):,} verified positive transactions.")
        return self

    @staticmethod
    def _compute_iqr_bounds(
        series: pd.Series, multiplier: float = 1.5
    ) -> Tuple[float, float]:
        """Calculates lower and upper Tukey fences using IQR."""
        q25 = series.quantile(0.25)
        q75 = series.quantile(0.75)
        iqr = q75 - q25
        lower_bound = q25 - (multiplier * iqr)
        upper_bound = q75 + (multiplier * iqr)
        return lower_bound, upper_bound

    def filter_statistical_outliers(
        self, multiplier: float = 1.5
    ) -> "RetailDataPipeline":
        """
        Applies Tukey's fences on Quantity and UnitPrice to eliminate extreme B2B bulk purchases
        or pricing glitches that distort consumer behavioral profiles.
        """
        qty_lower, qty_upper = self._compute_iqr_bounds(
            self.df["Quantity"], multiplier
        )
        price_lower, price_upper = self._compute_iqr_bounds(
            self.df["UnitPrice"], multiplier
        )

        # Retain transactions within valid bounds
        valid_mask = (
            (self.df["Quantity"] >= max(1, qty_lower))
            & (self.df["Quantity"] <= qty_upper)
            & (self.df["UnitPrice"] > 0)
            & (self.df["UnitPrice"] <= price_upper)
        )

        outlier_count = len(self.df) - valid_mask.sum()
        self.df = self.df[valid_mask].copy()

        # Compute LineItemTotal as net monetary spend
        self.df["LineItemTotal"] = (self.df["Quantity"] * self.df["UnitPrice"]).round(2)

        print(
            f"[REFINERY] Removed {outlier_count:,} statistical outlier rows. "
            f"Quantity bounded to [1, {qty_upper:.1f}], "
            f"UnitPrice bounded to (0, {price_upper:.2f}]."
        )
        return self

    def execute(self) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Executes the pipeline sequentially and returns clean transactions and audit reversals."""
        (
            self.load_data()
            .clean_text_and_nulls()
            .isolate_reversals_and_cancellations()
            .filter_statistical_outliers()
        )
        return self.df, self.reversals
