"""
sheets/enriched/sheet_duration.py
─────────────────────────────────
شیت: توزیع_مدت_مکالمه
"""

import pandas as pd
from .utils import col_exists_and_has_data, safe_mean


def write_duration_sheet(writer: pd.ExcelWriter, df: pd.DataFrame) -> None:
    """توزیع مدت مکالمه Rate (فقط ردیف‌های مچ‌شده با Rate)."""
    if not col_exists_and_has_data(df, "duration_bucket"):
        return

    rate_matched = df[df["has_rate_match"]].copy()
    if rate_matched.empty:
        return

    dur_dist = (
        rate_matched.groupby("duration_bucket", observed=True)
        .agg(
            count=("customer_10", "count"),
            avg_score=(
                "rate_score",
                lambda x: round(x.dropna().mean(), 2) if x.notna().any() else 0,
            ),
            avg_confidence=(
                "rate_confidence",
                lambda x: round(x.dropna().mean(), 1) if x.notna().any() else 0,
            ),
            avg_duration=(
                "rate_duration_seconds",
                lambda x: round(x.dropna().mean(), 0) if x.notna().any() else 0,
            ),
        )
        .reset_index()
    )

    total = dur_dist["count"].sum()
    dur_dist["percent"] = (dur_dist["count"] / total * 100).round(2) if total else 0

    dur_dist.to_excel(writer, index=False, sheet_name="توزیع_مدت_مکالمه")
