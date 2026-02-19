"""
sheets/enriched/lookups.py
──────────────────────────
ساخت lookup DataFrameها برای LEFT JOIN
از خروجی مچینگ Support و Rate.
"""

import pandas as pd
from .utils import safe_cols

# ──────────────────────────────────────────────
#  Support Lookup
# ──────────────────────────────────────────────

_SUP_COLS = [
    "customer_10", "monitoring_time",
    "sup_time", "sup_agent_ext",
    "sup_delta_minutes", "sup_confidence",
    "sup_match_status", "sup_jalali_date",
]

_SUP_RENAME = {
    "monitoring_event_time": "monitoring_time",
    "support_time": "sup_time",
    "support_agent_ext": "sup_agent_ext",
    "delta_minutes": "sup_delta_minutes",
    "confidence": "sup_confidence",
    "match_status": "sup_match_status",
    "support_jalali_date": "sup_jalali_date",
}


def build_support_lookup(support_matches_df: pd.DataFrame) -> pd.DataFrame:
    """
    از خروجی مچینگ Support یک lookup برای LEFT JOIN می‌سازد.
    کلید JOIN: (customer_10, monitoring_time)
    """
    if support_matches_df is None or support_matches_df.empty:
        return pd.DataFrame(columns=_SUP_COLS)

    df = support_matches_df.copy()

    if "match_status" in df.columns:
        df = df[df["match_status"] == "matched"].copy()

    if df.empty:
        return pd.DataFrame(columns=_SUP_COLS)

    df.rename(columns=_SUP_RENAME, inplace=True)

    return df[safe_cols(df, _SUP_COLS)].copy()


# ──────────────────────────────────────────────
#  Rate Lookup
# ──────────────────────────────────────────────

_RATE_COLS = [
    "customer_10", "monitoring_time",
    "rate_time", "rate_agent_ext", "rate_duration_seconds",
    "rate_score",
    "rate_delta_minutes", "rate_confidence",
    "rate_match_status", "rate_jalali_date",
]

_RATE_RENAME = {
    "monitoring_event_time": "monitoring_time",
    "delta_minutes": "rate_delta_minutes",
    "confidence": "rate_confidence",
    "match_status": "rate_match_status",
    "rate_jalali_date": "rate_jalali_date",
}


def build_rate_lookup(rate_matches_df: pd.DataFrame) -> pd.DataFrame:
    """
    از خروجی مچینگ Rate یک lookup برای LEFT JOIN می‌سازد.
    کلید JOIN: (customer_10, monitoring_time)
    شامل ستون rate_score (امتیاز نظرسنجی مشتری).
    """
    if rate_matches_df is None or rate_matches_df.empty:
        return pd.DataFrame(columns=_RATE_COLS)

    df = rate_matches_df.copy()

    if "match_status" in df.columns:
        df = df[df["match_status"] == "matched"].copy()

    if df.empty:
        return pd.DataFrame(columns=_RATE_COLS)

    df.rename(columns=_RATE_RENAME, inplace=True)

    # امتیاز نظرسنجی
    if "score" in df.columns:
        df.rename(columns={"score": "rate_score"}, inplace=True)
    elif "rate_score" not in df.columns:
        df["rate_score"] = None

    return df[safe_cols(df, _RATE_COLS)].copy()
