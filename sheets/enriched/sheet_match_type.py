"""
sheets/enriched/sheet_match_type.py
───────────────────────────────────
شیت: توزیع_نوع_مچ
"""

import pandas as pd


def write_match_type_sheet(writer: pd.ExcelWriter, df: pd.DataFrame) -> None:
    """توزیع نوع مچ (both / support_only / rate_only / none)."""
    if "match_type" not in df.columns:
        return

    match_dist = (
        df.groupby("match_type")
        .agg(
            count=("customer_10", "count"),
            avg_wait=("wait_seconds", lambda x: round(x.mean(), 1)),
        )
        .reset_index()
    )

    total = match_dist["count"].sum()
    match_dist["percent"] = (match_dist["count"] / total * 100).round(2) if total else 0

    # ترتیب معنادار
    order = ["both", "support_only", "rate_only", "none"]
    match_dist["_sort"] = match_dist["match_type"].map(
        {v: i for i, v in enumerate(order)}
    )
    match_dist.sort_values("_sort", inplace=True)
    match_dist.drop(columns="_sort", inplace=True)

    match_dist.to_excel(writer, index=False, sheet_name="توزیع_نوع_مچ")
