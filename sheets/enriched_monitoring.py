"""
sheets/enriched_monitoring.py
─────────────────────────────
Orchestrator: تولید فایل enriched_monitoring.xlsx

• تمام ردیف‌های Monitoring حفظ می‌شوند
• ستون‌های مچ Support و Rate به صورت LEFT JOIN اضافه می‌شوند
• ستون‌های محاسبه‌ای برای داشبورد
• ۱۰ شیت تحلیلی تولید می‌شود
"""

import pandas as pd
from pathlib import Path

from .enriched.utils import safe_cols
from .enriched.lookups import build_support_lookup, build_rate_lookup
from .enriched.computed_columns import add_all_computed
from .enriched.sheet_raw import write_raw_sheets
from .enriched.sheet_summary import write_summary_sheet
from .enriched.sheet_hourly import write_hourly_sheet
from .enriched.sheet_daily import write_daily_sheet
from .enriched.sheet_wait import write_wait_sheet
from .enriched.sheet_duration import write_duration_sheet
from .enriched.sheet_agent import write_agent_sheet
from .enriched.sheet_match_type import write_match_type_sheet


# ──────────────────────────────────────────────
#  ترتیب نهایی ستون‌ها
# ──────────────────────────────────────────────

_FINAL_COL_ORDER = [
    # شناسایی
    "customer_10", "customer_raw",
    # زمان مونیتورینگ
    "monitoring_time", "event_time", "connect_time",
    "agent_ext", "status",
    "wait_time", "wait_seconds", "wait_bucket",
    # Support Match
    "sup_time", "sup_agent_ext",
    "sup_delta_minutes", "sup_confidence", "sup_match_status",
    # Rate Match
    "rate_time", "rate_agent_ext", "rate_duration_seconds",
    "rate_score",
    "rate_delta_minutes", "rate_confidence", "rate_match_status",
    # محاسبه‌ای
    "has_support_match", "has_rate_match", "match_type",
    "duration_bucket",
    "call_date", "call_hour", "day_of_week",
]


# ──────────────────────────────────────────────
#  MAIN
# ──────────────────────────────────────────────

def generate_enriched_excel(
    monitoring_df: pd.DataFrame,
    support_matches_df: pd.DataFrame,
    rate_matches_df: pd.DataFrame,
    output_path: str = "sheets/enriched_monitoring.xlsx",
) -> str:
    """
    تولید فایل enriched_monitoring.xlsx.

    Returns: مسیر فایل ذخیره‌شده
    """

    # ─── STEP 1: آماده‌سازی پایه ───
    base = monitoring_df.copy()
    base_wanted = [
        "customer_10", "customer_raw",
        "event_time_normalized", "event_time",
        "agent_ext", "status",
        "wait_time", "wait_seconds",
    ]
    base = base[safe_cols(base, base_wanted)].copy()

    if "event_time_normalized" in base.columns:
        base.rename(columns={"event_time_normalized": "monitoring_time"}, inplace=True)
    elif "event_time" in base.columns:
        base.rename(columns={"event_time": "monitoring_time"}, inplace=True)

    # ─── STEP 2: ساخت Lookups ───
    sup_lookup = build_support_lookup(support_matches_df)
    rate_lookup = build_rate_lookup(rate_matches_df)

    print(f"   [DEBUG] base shape:          {base.shape}")
    print(f"   [DEBUG] sup_lookup shape:    {sup_lookup.shape}")
    print(f"   [DEBUG] rate_lookup shape:   {rate_lookup.shape}")

    # ─── STEP 3: LEFT JOIN — Support ───
    if not sup_lookup.empty and "monitoring_time" in sup_lookup.columns:
        enriched = base.merge(sup_lookup, on=["customer_10", "monitoring_time"], how="left")
    else:
        enriched = base.copy()
        for col in ["sup_time", "sup_agent_ext", "sup_delta_minutes",
                     "sup_confidence", "sup_match_status"]:
            enriched[col] = None

    # ─── STEP 4: LEFT JOIN — Rate ───
    if not rate_lookup.empty and "monitoring_time" in rate_lookup.columns:
        enriched = enriched.merge(rate_lookup, on=["customer_10", "monitoring_time"], how="left")
    else:
        for col in ["rate_time", "rate_agent_ext", "rate_duration_seconds",
                     "rate_score", "rate_delta_minutes", "rate_confidence",
                     "rate_match_status"]:
            enriched[col] = None

    # ─── STEP 5: ستون‌های محاسبه‌ای ───
    enriched = add_all_computed(enriched)

    # ─── STEP 6: مرتب‌سازی ستون‌ها ───
    final_cols = safe_cols(enriched, _FINAL_COL_ORDER)
    extra_cols = [c for c in enriched.columns if c not in final_cols]
    enriched = enriched[final_cols + extra_cols].copy()

    enriched.sort_values("monitoring_time", inplace=True, na_position="last")
    enriched.reset_index(drop=True, inplace=True)

    # ─── STEP 7: ذخیره اکسل ───
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        write_raw_sheets(writer, enriched)
        stats = write_summary_sheet(writer, enriched)
        write_hourly_sheet(writer, enriched)
        write_daily_sheet(writer, enriched)
        write_wait_sheet(writer, enriched)
        write_duration_sheet(writer, enriched)
        write_agent_sheet(writer, enriched)
        write_match_type_sheet(writer, enriched)

    # ─── STEP 8: گزارش کنسولی ───
    print(f"\n   ✅ Enriched monitoring saved: {output_path}")
    print(f"      → Total rows:      {stats['total']}")
    print(f"      → Connected:       {stats['connected']}")
    print(f"      → Abandoned:       {stats['abandoned']}")
    print(f"      → Support matched: {stats['sup_match']}")
    print(f"      → Rate matched:    {stats['rate_match']}")
    print(f"      → Both matched:    {stats['both_match']}")
    print(f"      → No match:        {stats['no_match']}")
    print(f"      → Sheets: 10 sheets generated")

    return output_path
