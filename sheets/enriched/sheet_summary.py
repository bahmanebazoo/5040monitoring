"""
sheets/enriched/sheet_summary.py
────────────────────────────────
شیت: خلاصه_آماری
"""

import pandas as pd
from .utils import safe_mean, connected_mask, abandoned_mask


def write_summary_sheet(writer: pd.ExcelWriter, df: pd.DataFrame) -> dict:
    """
    شیت خلاصه آماری + برگرداندن dict آمار برای گزارش کنسولی.
    """
    n = len(df)
    n_connected = int(connected_mask(df).sum())
    n_abandoned = int(abandoned_mask(df).sum())
    n_sup = int(df["has_support_match"].sum())
    n_rate = int(df["has_rate_match"].sum())
    n_both = int((df["has_support_match"] & df["has_rate_match"]).sum())
    n_sup_only = int((df["has_support_match"] & ~df["has_rate_match"]).sum())
    n_rate_only = int((~df["has_support_match"] & df["has_rate_match"]).sum())
    n_none = int((~df["has_support_match"] & ~df["has_rate_match"]).sum())

    pct = lambda x: round(x / n * 100, 2) if n else 0

    rows = [
        ("کل تماس‌های مونیتورینگ", n),
        ("تماس‌های وصل‌شده", n_connected),
        ("تماس‌های رها‌شده", n_abandoned),
        ("─── مچینگ ───", ""),
        ("مچ‌شده با Support", n_sup),
        ("مچ‌شده با Rate", n_rate),
        ("مچ‌شده با هر دو", n_both),
        ("فقط Support", n_sup_only),
        ("فقط Rate", n_rate_only),
        ("بدون مچ", n_none),
        ("─── نرخ‌ها ───", ""),
        ("نرخ مچ Support (%)", pct(n_sup)),
        ("نرخ مچ Rate (%)", pct(n_rate)),
        ("نرخ مچ حداقل یکی (%)", pct(n_sup + n_rate - n_both)),
        ("─── میانگین‌ها ───", ""),
        ("میانگین زمان انتظار (ثانیه)", safe_mean(df.get("wait_seconds"), 1)),
        ("میانگین مدت مکالمه Rate (ثانیه)", safe_mean(df.get("rate_duration_seconds"), 1)),
        ("میانگین امتیاز اعتماد Support", safe_mean(df.get("sup_confidence"), 1)),
        ("میانگین امتیاز اعتماد Rate", safe_mean(df.get("rate_confidence"), 1)),
        ("میانگین امتیاز نظرسنجی Rate", safe_mean(df.get("rate_score"), 2)),
    ]

    summary_df = pd.DataFrame(rows, columns=["متریک", "مقدار"])
    summary_df.to_excel(writer, index=False, sheet_name="خلاصه_آماری")

    # dict برای گزارش کنسولی
    return {
        "total": n,
        "connected": n_connected,
        "abandoned": n_abandoned,
        "sup_match": n_sup,
        "rate_match": n_rate,
        "both_match": n_both,
        "no_match": n_none,
    }
