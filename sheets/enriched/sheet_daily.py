"""
sheets/enriched/sheet_daily.py
──────────────────────────────
شیت: توزیع_روزانه
"""

import pandas as pd
from .utils import col_exists_and_has_data, STATUS_CONNECTED


def write_daily_sheet(writer: pd.ExcelWriter, df: pd.DataFrame) -> None:
    """توزیع تماس‌ها به تفکیک روز."""
    if not col_exists_and_has_data(df, "call_date"):
        return

    daily = (
        df.groupby(["call_date", "day_of_week"])
        .agg(
            total_calls=("customer_10", "count"),
            connected=("status", lambda x: (x == STATUS_CONNECTED).sum()),
            abandoned=("status", lambda x: (x != STATUS_CONNECTED).sum()),
            avg_wait=("wait_seconds", "mean"),
            support_matched=("has_support_match", "sum"),
            rate_matched=("has_rate_match", "sum"),
        )
        .reset_index()
        .sort_values("call_date")
    )

    daily["avg_wait"] = daily["avg_wait"].round(1)
    daily.to_excel(writer, index=False, sheet_name="توزیع_روزانه")
