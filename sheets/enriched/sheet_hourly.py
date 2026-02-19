"""
sheets/enriched/sheet_hourly.py
───────────────────────────────
شیت: توزیع_ساعتی
"""

import pandas as pd
from .utils import col_exists_and_has_data, STATUS_CONNECTED


def write_hourly_sheet(writer: pd.ExcelWriter, df: pd.DataFrame) -> None:
    """توزیع تماس‌ها به تفکیک ساعت."""
    if not col_exists_and_has_data(df, "call_hour"):
        return

    hourly = (
        df.groupby("call_hour")
        .agg(
            total_calls=("customer_10", "count"),
            connected=("status", lambda x: (x == STATUS_CONNECTED).sum()),
            abandoned=("status", lambda x: (x != STATUS_CONNECTED).sum()),
            avg_wait=("wait_seconds", "mean"),
            support_matched=("has_support_match", "sum"),
            rate_matched=("has_rate_match", "sum"),
        )
        .reset_index()
    )

    hourly["avg_wait"] = hourly["avg_wait"].round(1)
    hourly.to_excel(writer, index=False, sheet_name="توزیع_ساعتی")
