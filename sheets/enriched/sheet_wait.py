"""
sheets/enriched/sheet_wait.py
─────────────────────────────
شیت: توزیع_انتظار
"""

import pandas as pd
from .utils import col_exists_and_has_data, STATUS_CONNECTED


def write_wait_sheet(writer: pd.ExcelWriter, df: pd.DataFrame) -> None:
    """توزیع تماس‌ها بر اساس دسته‌بندی زمان انتظار."""
    if not col_exists_and_has_data(df, "wait_bucket"):
        return

    wait_dist = (
        df.groupby("wait_bucket", observed=True)
        .agg(
            count=("customer_10", "count"),
            connected=("status", lambda x: (x == STATUS_CONNECTED).sum()),
            abandoned=("status", lambda x: (x != STATUS_CONNECTED).sum()),
        )
        .reset_index()
    )

    total = wait_dist["count"].sum()
    wait_dist["percent"] = (wait_dist["count"] / total * 100).round(2) if total else 0

    wait_dist.to_excel(writer, index=False, sheet_name="توزیع_انتظار")
