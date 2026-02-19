"""
sheets/enriched/sheet_agent.py
──────────────────────────────
شیت: آمار_داخلی‌ها (عملکرد هر اپراتور)
"""

import pandas as pd
from .utils import connected_mask


def write_agent_sheet(writer: pd.ExcelWriter, df: pd.DataFrame) -> None:
    """آمار عملکرد به تفکیک شماره داخلی."""
    if "agent_ext" not in df.columns:
        return

    agent_data = df[connected_mask(df)].copy()
    if agent_data.empty:
        return

    agent_stats = (
        agent_data
        .groupby("agent_ext")
        .agg(
            total_calls=("customer_10", "count"),
            avg_wait=("wait_seconds", lambda x: round(x.mean(), 1)),
            support_matched=("has_support_match", "sum"),
            rate_matched=("has_rate_match", "sum"),
            avg_rate_duration=(
                "rate_duration_seconds",
                lambda x: round(x.dropna().mean(), 1) if x.notna().any() else 0,
            ),
            avg_rate_score=(
                "rate_score",
                lambda x: round(x.dropna().mean(), 2) if x.notna().any() else 0,
            ),
            avg_sup_confidence=(
                "sup_confidence",
                lambda x: round(x.dropna().mean(), 1) if x.notna().any() else 0,
            ),
            avg_rate_confidence=(
                "rate_confidence",
                lambda x: round(x.dropna().mean(), 1) if x.notna().any() else 0,
            ),
        )
        .reset_index()
        .sort_values("total_calls", ascending=False)
    )

    # نرخ مچ
    agent_stats["sup_match_rate"] = (
        (agent_stats["support_matched"] / agent_stats["total_calls"] * 100).round(1)
    )
    agent_stats["rate_match_rate"] = (
        (agent_stats["rate_matched"] / agent_stats["total_calls"] * 100).round(1)
    )

    agent_stats.to_excel(writer, index=False, sheet_name="آمار_داخلی‌ها")
