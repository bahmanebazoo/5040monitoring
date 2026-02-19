"""
sheets/enriched/computed_columns.py
───────────────────────────────────
اضافه کردن ستون‌های محاسبه‌ای به DataFrame غنی‌شده.
"""

import pandas as pd
import numpy as np
from .utils import WEEKDAY_FA, col_exists_and_has_data


def add_match_flags(df: pd.DataFrame) -> pd.DataFrame:
    """ستون‌های has_support_match, has_rate_match, match_type."""

    # Support match flag
    if "sup_match_status" in df.columns:
        df["has_support_match"] = df["sup_match_status"].eq("matched")
    else:
        df["has_support_match"] = False

    # Rate match flag
    if "rate_match_status" in df.columns:
        df["has_rate_match"] = df["rate_match_status"].eq("matched")
    else:
        df["has_rate_match"] = False

    # Match type
    df["match_type"] = np.where(
        df["has_support_match"] & df["has_rate_match"], "both",
        np.where(
            df["has_support_match"], "support_only",
            np.where(
                df["has_rate_match"], "rate_only",
                "none"
            )
        )
    )

    return df


def add_connect_time(df: pd.DataFrame) -> pd.DataFrame:
    """ستون connect_time = monitoring_time + wait_seconds."""
    if "monitoring_time" in df.columns and "wait_seconds" in df.columns:
        df["wait_seconds"] = pd.to_numeric(
            df["wait_seconds"], errors="coerce"
        ).fillna(0)
        df["connect_time"] = (
            pd.to_datetime(df["monitoring_time"], errors="coerce")
            + pd.to_timedelta(df["wait_seconds"], unit="s")
        )
    else:
        df["connect_time"] = None
    return df


def add_wait_bucket(df: pd.DataFrame) -> pd.DataFrame:
    """ستون wait_bucket: دسته‌بندی زمان انتظار."""
    if col_exists_and_has_data(df, "wait_seconds"):
        df["wait_bucket"] = pd.cut(
            df["wait_seconds"].fillna(0),
            bins=[0, 30, 60, 120, 180, 300, float("inf")],
            labels=["0-30s", "30-60s", "1-2min", "2-3min", "3-5min", "5min+"],
            right=True,
            include_lowest=True,
        )
    else:
        df["wait_bucket"] = None
    return df


def add_duration_bucket(df: pd.DataFrame) -> pd.DataFrame:
    """ستون duration_bucket: دسته‌بندی مدت مکالمه Rate."""
    if col_exists_and_has_data(df, "rate_duration_seconds"):
        df["duration_bucket"] = pd.cut(
            df["rate_duration_seconds"].fillna(0),
            bins=[0, 60, 180, 300, 600, 900, float("inf")],
            labels=["0-1min", "1-3min", "3-5min", "5-10min", "10-15min", "15min+"],
            right=True,
            include_lowest=True,
        )
    else:
        df["duration_bucket"] = None
    return df


def add_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    """ستون‌های call_date, call_hour, day_of_week."""
    if "monitoring_time" in df.columns:
        mt = pd.to_datetime(df["monitoring_time"], errors="coerce")
        df["call_date"] = mt.dt.date
        df["call_hour"] = mt.dt.hour
        df["day_of_week"] = mt.dt.dayofweek.map(WEEKDAY_FA)
    else:
        df["call_date"] = None
        df["call_hour"] = None
        df["day_of_week"] = None
    return df


def add_all_computed(df: pd.DataFrame) -> pd.DataFrame:
    """تمام ستون‌های محاسبه‌ای را یکجا اضافه می‌کند."""
    df = add_match_flags(df)
    df = add_connect_time(df)
    df = add_wait_bucket(df)
    df = add_duration_bucket(df)
    df = add_date_columns(df)
    return df
