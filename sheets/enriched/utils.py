"""
sheets/enriched/utils.py
────────────────────────
توابع و ثابت‌های مشترک بین تمام شیت‌ها.
"""

import pandas as pd

# ──────────────────────────────────────────────
#  ثابت‌ها
# ──────────────────────────────────────────────

WEEKDAY_FA = {
    0: "دوشنبه",
    1: "سه‌شنبه",
    2: "چهارشنبه",
    3: "پنج‌شنبه",
    4: "جمعه",
    5: "شنبه",
    6: "یک‌شنبه",
}

STATUS_CONNECTED = "وصل شده"


# ──────────────────────────────────────────────
#  توابع مشترک
# ──────────────────────────────────────────────

def safe_cols(df: pd.DataFrame, wanted: list[str]) -> list[str]:
    """فقط ستون‌هایی را برگردان که واقعاً در DataFrame وجود دارند."""
    return [c for c in wanted if c in df.columns]


def safe_mean(series: pd.Series, decimals: int = 1) -> float:
    """میانگین امن: اگر همه NaN بود، 0 برمی‌گرداند."""
    if series is not None and series.notna().any():
        return round(series.dropna().mean(), decimals)
    return 0


def col_exists_and_has_data(df: pd.DataFrame, col: str) -> bool:
    """آیا ستون وجود دارد و حداقل یک مقدار غیر‌NaN دارد؟"""
    return col in df.columns and df[col].notna().any()


def connected_mask(df: pd.DataFrame) -> pd.Series:
    """ماسک تماس‌های وصل‌شده."""
    if "status" in df.columns:
        return df["status"] == STATUS_CONNECTED
    return pd.Series(False, index=df.index)


def abandoned_mask(df: pd.DataFrame) -> pd.Series:
    """ماسک تماس‌های رها‌شده."""
    return ~connected_mask(df)
