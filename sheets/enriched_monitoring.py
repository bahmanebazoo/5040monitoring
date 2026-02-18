"""
sheets/enriched_monitoring.py
─────────────────────────────
تولید اکسل داده خام غنی‌شده:
  • تمام ردیف‌های Monitoring حفظ می‌شوند
  • ستون‌های مچ Support و Rate به صورت LEFT JOIN اضافه می‌شوند
  • خروجی مناسب برای ساخت نمودار / داشبورد
"""

import pandas as pd
from pathlib import Path


# ──────────────────────────────────────────────
#  HELPERS
# ──────────────────────────────────────────────

def _safe_cols(df: pd.DataFrame, wanted: list[str]) -> list[str]:
    """فقط ستون‌هایی را برگردان که واقعاً در DataFrame وجود دارند."""
    return [c for c in wanted if c in df.columns]


def _build_support_lookup(support_matches_df: pd.DataFrame) -> pd.DataFrame:
    """
    از خروجی مچینگ Support یک lookup برای LEFT JOIN می‌سازد.
    کلید JOIN: (customer_10, monitoring_time)
    """
    empty_cols = [
        "customer_10", "monitoring_time",
        "sup_time", "sup_agent_ext",
        "sup_delta_minutes", "sup_confidence", "sup_match_status",
    ]

    if support_matches_df is None or support_matches_df.empty:
        return pd.DataFrame(columns=empty_cols)

    df = support_matches_df.copy()

    # فقط مچ‌شده‌ها را نگه‌دار (unmatched را نمی‌خواهیم JOIN کنیم)
    if "match_status" in df.columns:
        df = df[df["match_status"] == "matched"].copy()

    if df.empty:
        return pd.DataFrame(columns=empty_cols)

    rename_map = {
        "support_time": "sup_time",
        "support_agent_ext": "sup_agent_ext",
        "delta_minutes": "sup_delta_minutes",
        "confidence": "sup_confidence",
        "match_status": "sup_match_status",
    }
    df.rename(columns=rename_map, inplace=True)

    keep = _safe_cols(df, empty_cols)
    return df[keep].copy()


def _build_rate_lookup(rate_matches_df: pd.DataFrame) -> pd.DataFrame:
    """
    از خروجی مچینگ Rate یک lookup برای LEFT JOIN می‌سازد.
    کلید JOIN: (customer_10, monitoring_time)
    """
    empty_cols = [
        "customer_10", "monitoring_time",
        "rate_time", "rate_agent_ext", "rate_duration_seconds",
        "rate_delta_minutes", "rate_confidence", "rate_match_status",
    ]

    if rate_matches_df is None or rate_matches_df.empty:
        return pd.DataFrame(columns=empty_cols)

    df = rate_matches_df.copy()

    # فقط مچ‌شده‌ها
    if "match_status" in df.columns:
        df = df[df["match_status"] == "matched"].copy()

    if df.empty:
        return pd.DataFrame(columns=empty_cols)

    rename_map = {
        "delta_minutes": "rate_delta_minutes",
        "confidence": "rate_confidence",
        "match_status": "rate_match_status",
    }
    df.rename(columns=rename_map, inplace=True)

    keep = _safe_cols(df, empty_cols)
    return df[keep].copy()


# ──────────────────────────────────────────────
#  MAIN FUNCTION
# ──────────────────────────────────────────────

def generate_enriched_excel(
    monitoring_df: pd.DataFrame,
    support_matches_df: pd.DataFrame,
    rate_matches_df: pd.DataFrame,
    output_path: str = "sheets/enriched_monitoring.xlsx",
) -> str:
    """
    تمام ردیف‌های Monitoring + ستون‌های مچ Support و Rate
    را در یک فایل اکسل ذخیره می‌کند.

    Returns
    -------
    str : مسیر فایل ذخیره‌شده
    """

    # ─── 1. آماده‌سازی Monitoring (پایه) ───
    base = monitoring_df.copy()

    base_cols = _safe_cols(base, [
        "customer_10", "customer_raw",
        "event_time_normalized", "event_time",
        "agent_ext", "status",
        "wait_time", "wait_seconds",
    ])
    base = base[base_cols].copy()

    # ستون کلید JOIN — تغییر نام event_time_normalized → monitoring_time
    if "event_time_normalized" in base.columns:
        base.rename(columns={"event_time_normalized": "monitoring_time"}, inplace=True)
    elif "event_time" in base.columns:
        # fallback: اگر normalized نبود از event_time استفاده کن
        base.rename(columns={"event_time": "monitoring_time"}, inplace=True)

    # ─── 2. ساخت Lookups ───
    sup_lookup = _build_support_lookup(support_matches_df)
    rate_lookup = _build_rate_lookup(rate_matches_df)

    # ─── 3. بررسی ستون‌ها قبل از JOIN ───
    print(f"   [DEBUG] base columns:        {list(base.columns)}")
    print(f"   [DEBUG] base shape:          {base.shape}")
    print(f"   [DEBUG] sup_lookup columns:  {list(sup_lookup.columns)}")
    print(f"   [DEBUG] sup_lookup shape:    {sup_lookup.shape}")
    print(f"   [DEBUG] rate_lookup columns: {list(rate_lookup.columns)}")
    print(f"   [DEBUG] rate_lookup shape:   {rate_lookup.shape}")

    # ─── 4. LEFT JOIN: Monitoring ← Support ───
    sup_empty_cols = ["sup_time", "sup_agent_ext", "sup_delta_minutes",
                      "sup_confidence", "sup_match_status"]

    if not sup_lookup.empty and "monitoring_time" in sup_lookup.columns:
        enriched = base.merge(
            sup_lookup,
            on=["customer_10", "monitoring_time"],
            how="left",
        )
    else:
        enriched = base.copy()
        for col in sup_empty_cols:
            enriched[col] = None

    # ─── 5. LEFT JOIN: Monitoring ← Rate ───
    rate_empty_cols = ["rate_time", "rate_agent_ext", "rate_duration_seconds",
                       "rate_delta_minutes", "rate_confidence", "rate_match_status"]

    if not rate_lookup.empty and "monitoring_time" in rate_lookup.columns:
        enriched = enriched.merge(
            rate_lookup,
            on=["customer_10", "monitoring_time"],
            how="left",
        )
    else:
        for col in rate_empty_cols:
            enriched[col] = None

    # ─── 6. ستون‌های کمکی برای نمودار ───

    # آیا Support مچ شده؟
    if "sup_match_status" in enriched.columns:
        enriched["has_support_match"] = enriched["sup_match_status"].eq("matched")
    else:
        enriched["has_support_match"] = False

    # آیا Rate مچ شده؟
    if "rate_match_status" in enriched.columns:
        enriched["has_rate_match"] = enriched["rate_match_status"].eq("matched")
    else:
        enriched["has_rate_match"] = False

    # دسته‌بندی زمان انتظار
    if "wait_seconds" in enriched.columns and enriched["wait_seconds"].notna().any():
        enriched["wait_bucket"] = pd.cut(
            enriched["wait_seconds"].fillna(0),
            bins=[0, 30, 60, 120, 180, 300, float("inf")],
            labels=["0-30s", "30-60s", "1-2min", "2-3min", "3-5min", "5min+"],
            right=True,
            include_lowest=True,
        )
    else:
        enriched["wait_bucket"] = None

    # ساعت تماس
    if "monitoring_time" in enriched.columns:
        enriched["call_hour"] = pd.to_datetime(
            enriched["monitoring_time"], errors="coerce"
        ).dt.hour

    # مرتب‌سازی نهایی
    enriched.sort_values("monitoring_time", inplace=True, na_position="last")
    enriched.reset_index(drop=True, inplace=True)

    # ─── 7. ذخیره ───
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        # شیت ۱: داده خام کامل
        enriched.to_excel(writer, index=False, sheet_name="داده_خام_غنی‌شده")

        # شیت ۲: فقط وصل‌شده‌ها
        connected = enriched[enriched["status"] == "وصل شده"]
        connected.to_excel(writer, index=False, sheet_name="وصل_شده")

        # شیت ۳: فقط رها‌شده‌ها
        abandoned = enriched[enriched["status"] != "وصل شده"]
        abandoned.to_excel(writer, index=False, sheet_name="رها_شده")

        # شیت ۴: خلاصه آماری
        n = len(enriched)
        n_connected = len(connected)
        n_abandoned = len(abandoned)
        n_sup_match = int(enriched["has_support_match"].sum())
        n_rate_match = int(enriched["has_rate_match"].sum())
        n_both = int((enriched["has_support_match"] & enriched["has_rate_match"]).sum())

        summary_data = {
            "متریک": [
                "کل تماس‌های مونیتورینگ",
                "تماس‌های وصل‌شده",
                "تماس‌های رها‌شده",
                "مچ‌شده با Support",
                "مچ‌شده با Rate",
                "مچ‌شده با هر دو",
                "نرخ مچ Support (%)",
                "نرخ مچ Rate (%)",
                "میانگین زمان انتظار (ثانیه)",
                "میانگین مدت مکالمه Rate (ثانیه)",
                "میانگین امتیاز Support",
                "میانگین امتیاز Rate",
            ],
            "مقدار": [
                n,
                n_connected,
                n_abandoned,
                n_sup_match,
                n_rate_match,
                n_both,
                round(n_sup_match / n * 100, 2) if n else 0,
                round(n_rate_match / n * 100, 2) if n else 0,
                round(enriched["wait_seconds"].mean(), 1)
                    if "wait_seconds" in enriched.columns and enriched["wait_seconds"].notna().any() else 0,
                round(enriched["rate_duration_seconds"].dropna().mean(), 1)
                    if "rate_duration_seconds" in enriched.columns and enriched["rate_duration_seconds"].notna().any() else 0,
                round(enriched["sup_confidence"].dropna().mean(), 1)
                    if "sup_confidence" in enriched.columns and enriched["sup_confidence"].notna().any() else 0,
                round(enriched["rate_confidence"].dropna().mean(), 1)
                    if "rate_confidence" in enriched.columns and enriched["rate_confidence"].notna().any() else 0,
            ],
        }
        pd.DataFrame(summary_data).to_excel(writer, index=False, sheet_name="خلاصه_آماری")

        # شیت ۵: توزیع ساعتی
        if "call_hour" in enriched.columns and enriched["call_hour"].notna().any():
            hourly = (
                enriched.groupby("call_hour")
                .agg(
                    total_calls=("customer_10", "count"),
                    connected=("status", lambda x: (x == "وصل شده").sum()),
                    abandoned=("status", lambda x: (x != "وصل شده").sum()),
                    avg_wait=("wait_seconds", "mean"),
                    support_matched=("has_support_match", "sum"),
                    rate_matched=("has_rate_match", "sum"),
                )
                .reset_index()
            )
            hourly.to_excel(writer, index=False, sheet_name="توزیع_ساعتی")

        # شیت ۶: توزیع زمان انتظار
        if "wait_bucket" in enriched.columns and enriched["wait_bucket"].notna().any():
            wait_dist = (
                enriched.groupby("wait_bucket", observed=True)
                .agg(
                    count=("customer_10", "count"),
                    connected=("status", lambda x: (x == "وصل شده").sum()),
                    abandoned=("status", lambda x: (x != "وصل شده").sum()),
                )
                .reset_index()
            )
            wait_dist.to_excel(writer, index=False, sheet_name="توزیع_انتظار")

        # شیت ۷: آمار به تفکیک داخلی
        if "agent_ext" in enriched.columns:
            agent_data = enriched[enriched["status"] == "وصل شده"]
            if not agent_data.empty:
                agent_stats = (
                    agent_data
                    .groupby("agent_ext")
                    .agg(
                        total_calls=("customer_10", "count"),
                        avg_wait=("wait_seconds", "mean"),
                        support_matched=("has_support_match", "sum"),
                        rate_matched=("has_rate_match", "sum"),
                        avg_rate_duration=(
                            "rate_duration_seconds",
                            lambda x: x.dropna().mean() if x.notna().any() else 0,
                        ),
                        avg_sup_confidence=(
                            "sup_confidence",
                            lambda x: x.dropna().mean() if x.notna().any() else 0,
                        ),
                        avg_rate_confidence=(
                            "rate_confidence",
                            lambda x: x.dropna().mean() if x.notna().any() else 0,
                        ),
                    )
                    .reset_index()
                    .sort_values("total_calls", ascending=False)
                )
                agent_stats.to_excel(writer, index=False, sheet_name="آمار_داخلی‌ها")

    print(f"   ✅ Enriched monitoring saved: {output_path}")
    print(f"      → Sheets: داده_خام_غنی‌شده, وصل_شده, رها_شده, خلاصه_آماری, توزیع_ساعتی, توزیع_انتظار, آمار_داخلی‌ها")
    print(f"      → Total rows: {len(enriched)}")
    print(f"      → Support matched: {n_sup_match}")
    print(f"      → Rate matched: {n_rate_match}")

    return output_path
