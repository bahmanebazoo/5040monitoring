import pandas as pd
import numpy as np
from pathlib import Path


def _parse_wait_to_seconds(wait_col: pd.Series) -> pd.Series:
    """
    تبدیل ستون زمان انتظار به ثانیه.
    فرمت‌های پشتیبانی‌شده: عدد (ثانیه)، HH:MM:SS، MM:SS، Timedelta
    """
    def _convert(val):
        if pd.isna(val):
            return np.nan
        if isinstance(val, (int, float)):
            return float(val)
        if isinstance(val, pd.Timedelta):
            return val.total_seconds()
        val_str = str(val).strip()
        parts = val_str.split(":")
        try:
            if len(parts) == 3:
                return int(parts[0]) * 3600 + int(parts[1]) * 60 + float(parts[2])
            elif len(parts) == 2:
                return int(parts[0]) * 60 + float(parts[1])
            else:
                return float(val_str)
        except (ValueError, TypeError):
            return np.nan

    return wait_col.apply(_convert)


def generate_analytics_excel(
    monitoring_df: pd.DataFrame,
    rate_df: pd.DataFrame,
    support_matches_df: pd.DataFrame,
    rate_matches_df: pd.DataFrame,
    output_path: str = "sheets/analytics_report.xlsx",
) -> str:
    """
    تولید فایل اکسل تحلیلی با شیت‌های مختلف.

    Parameters
    ----------
    monitoring_df : داده خام مونیتورینگ (نرمال‌شده)
    rate_df : داده خام Rate (نرمال‌شده)
    support_matches_df : خروجی مچینگ Monitoring ↔ Support
    rate_matches_df : خروجی مچینگ Monitoring ↔ Rate
    output_path : مسیر فایل خروجی

    Returns
    -------
    str : مسیر فایل ذخیره‌شده
    """

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    # ═══════════════════════════════════════════════════════
    # آماده‌سازی داده‌ها
    # ═══════════════════════════════════════════════════════

    # ---- Monitoring ----
    mon = monitoring_df.copy()
    mon["wait_seconds_num"] = _parse_wait_to_seconds(mon["wait_seconds"])

    # ---- Rate Matches (فقط مچ‌شده‌ها) ----
    rm = rate_matches_df.copy()
    rm_matched = rm[rm["match_status"] == "matched"].copy()
    rm_matched["monitoring_wait_seconds_num"] = _parse_wait_to_seconds(
        rm_matched["monitoring_wait_seconds"]
    )

    # ═══════════════════════════════════════════════════════
    # شیت ۱: خلاصه کلی (Summary)
    # ═══════════════════════════════════════════════════════

    total_monitoring = len(mon)
    total_connected = len(mon[mon["status"] == "وصل شده"])
    total_abandoned = len(mon[mon["status"] == "رها شده"])

    # تماس وصل‌شده با انتظار بیش از 2 دقیقه (120 ثانیه)
    connected_wait_over_2min = len(
        mon[(mon["status"] == "وصل شده") & (mon["wait_seconds_num"] > 120)]
    )

    # تماس رها شده با انتظار زیر 30 ثانیه
    abandoned_wait_under_30s = len(
        mon[(mon["status"] == "رها شده") & (mon["wait_seconds_num"] < 30)]
    )

    total_rate = len(rate_df)
    total_rate_matched = len(rm_matched)
    rate_match_pct = (
        (total_rate_matched / total_connected * 100) if total_connected else 0
    )

    total_support_matched = len(
        support_matches_df[support_matches_df["match_status"] == "matched"]
    ) if support_matches_df is not None and len(support_matches_df) > 0 else 0

    avg_wait_all = mon["wait_seconds_num"].mean()
    avg_wait_connected = mon.loc[
        mon["status"] == "وصل شده", "wait_seconds_num"
    ].mean()
    avg_wait_abandoned = mon.loc[
        mon["status"] == "رها شده", "wait_seconds_num"
    ].mean()

    summary_data = {
        "شاخص": [
            "کل تماس‌های مونیتورینگ",
            "تماس‌های وصل‌شده",
            "تماس‌های رها شده",
            "وصل‌شده با انتظار بیش از ۲ دقیقه",
            "رها شده با انتظار زیر ۳۰ ثانیه",
            "کل رکوردهای Rate",
            "مچ‌شده Monitoring ↔ Rate",
            "درصد مچ Rate (از وصل‌شده‌ها)",
            "مچ‌شده Monitoring ↔ Support",
            "میانگین زمان انتظار (همه - ثانیه)",
            "میانگین زمان انتظار (وصل‌شده - ثانیه)",
            "میانگین زمان انتظار (رها شده - ثانیه)",
        ],
        "مقدار": [
            total_monitoring,
            total_connected,
            total_abandoned,
            connected_wait_over_2min,
            abandoned_wait_under_30s,
            total_rate,
            total_rate_matched,
            round(rate_match_pct, 2),
            total_support_matched,
            round(avg_wait_all, 1) if not pd.isna(avg_wait_all) else 0,
            round(avg_wait_connected, 1) if not pd.isna(avg_wait_connected) else 0,
            round(avg_wait_abandoned, 1) if not pd.isna(avg_wait_abandoned) else 0,
        ],
    }
    summary_df = pd.DataFrame(summary_data)

    # ═══════════════════════════════════════════════════════
    # شیت ۲: آمار به تفکیک داخلی (Per Agent)
    # ═══════════════════════════════════════════════════════

    # --- از Monitoring ---
    mon_connected = mon[mon["status"] == "وصل شده"].copy()

    mon_agent_stats = (
        mon_connected.groupby("agent_ext")
        .agg(
            تعداد_تماس_مونیتورینگ=("agent_ext", "size"),
            میانگین_انتظار_ثانیه=("wait_seconds_num", "mean"),
        )
        .reset_index()
        .rename(columns={"agent_ext": "داخلی"})
    )
    mon_agent_stats["میانگین_انتظار_ثانیه"] = mon_agent_stats[
        "میانگین_انتظار_ثانیه"
    ].round(1)

    # --- از Rate (خام) ---
    rate_work = rate_df.copy()
    rate_work["duration_seconds_num"] = pd.to_numeric(
        rate_work.get("duration_seconds", pd.Series(dtype=float)),
        errors="coerce",
    ).fillna(0)

    rate_agent_stats = (
        rate_work.groupby("agent_ext")
        .agg(
            تعداد_تماس_Rate=("agent_ext", "size"),
            میانگین_مکالمه_ثانیه=("duration_seconds_num", "mean"),
            مجموع_مکالمه_ثانیه=("duration_seconds_num", "sum"),
        )
        .reset_index()
        .rename(columns={"agent_ext": "داخلی"})
    )
    rate_agent_stats["میانگین_مکالمه_ثانیه"] = rate_agent_stats[
        "میانگین_مکالمه_ثانیه"
    ].round(1)
    rate_agent_stats["مجموع_مکالمه_ثانیه"] = rate_agent_stats[
        "مجموع_مکالمه_ثانیه"
    ].round(0)

    # --- از Rate Matches (مچ‌شده) ---
    if not rm_matched.empty and "rate_agent_ext" in rm_matched.columns:
        rm_matched["rate_duration_seconds_num"] = pd.to_numeric(
            rm_matched.get("rate_duration_seconds", pd.Series(dtype=float)),
            errors="coerce",
        ).fillna(0)

        matched_agent_stats = (
            rm_matched.groupby("rate_agent_ext")
            .agg(
                تعداد_مچ_شده=("rate_agent_ext", "size"),
                میانگین_مکالمه_مچ_ثانیه=("rate_duration_seconds_num", "mean"),
                مجموع_مکالمه_مچ_ثانیه=("rate_duration_seconds_num", "sum"),
                میانگین_امتیاز_اطمینان=("confidence", "mean"),
            )
            .reset_index()
            .rename(columns={"rate_agent_ext": "داخلی"})
        )
        matched_agent_stats["میانگین_مکالمه_مچ_ثانیه"] = matched_agent_stats[
            "میانگین_مکالمه_مچ_ثانیه"
        ].round(1)
        matched_agent_stats["مجموع_مکالمه_مچ_ثانیه"] = matched_agent_stats[
            "مجموع_مکالمه_مچ_ثانیه"
        ].round(0)
        matched_agent_stats["میانگین_امتیاز_اطمینان"] = matched_agent_stats[
            "میانگین_امتیاز_اطمینان"
        ].round(1)
    else:
        matched_agent_stats = pd.DataFrame(
            columns=[
                "داخلی",
                "تعداد_مچ_شده",
                "میانگین_مکالمه_مچ_ثانیه",
                "مجموع_مکالمه_مچ_ثانیه",
                "میانگین_امتیاز_اطمینان",
            ]
        )

    # --- ترکیب همه آمار داخلی‌ها ---
    agent_merged = (
        mon_agent_stats.merge(rate_agent_stats, on="داخلی", how="outer")
        .merge(matched_agent_stats, on="داخلی", how="outer")
        .fillna(0)
    )

    # ستون‌های عددی int
    for col in ["تعداد_تماس_مونیتورینگ", "تعداد_تماس_Rate", "تعداد_مچ_شده"]:
        if col in agent_merged.columns:
            agent_merged[col] = agent_merged[col].astype(int)

    agent_merged.sort_values("تعداد_تماس_مونیتورینگ", ascending=False, inplace=True)
    agent_merged.reset_index(drop=True, inplace=True)

    # ═══════════════════════════════════════════════════════
    # شیت ۳: تماس‌های وصل‌شده با انتظار بیش از ۲ دقیقه
    # ═══════════════════════════════════════════════════════

    connected_long_wait = mon[
        (mon["status"] == "وصل شده") & (mon["wait_seconds_num"] > 120)
    ].copy()

    connected_long_wait["wait_formatted"] = connected_long_wait[
        "wait_seconds_num"
    ].apply(
        lambda s: (
            f"{int(s // 3600):02d}:{int((s % 3600) // 60):02d}:{int(s % 60):02d}"
            if not pd.isna(s)
            else ""
        )
    )

    long_wait_export = connected_long_wait[
        [
            "customer_10",
            "agent_ext",
            "event_time_normalized",
            "wait_formatted",
            "wait_seconds_num",
        ]
    ].rename(
        columns={
            "customer_10": "شماره_مشتری",
            "agent_ext": "داخلی",
            "event_time_normalized": "زمان_تماس",
            "wait_formatted": "زمان_انتظار",
            "wait_seconds_num": "انتظار_ثانیه",
        }
    )
    long_wait_export.sort_values("انتظار_ثانیه", ascending=False, inplace=True)
    long_wait_export.reset_index(drop=True, inplace=True)

    # ═══════════════════════════════════════════════════════
    # شیت ۴: تماس‌های رها شده با انتظار زیر ۳۰ ثانیه
    # ═══════════════════════════════════════════════════════

    abandoned_short_wait = mon[
        (mon["status"] == "رها شده") & (mon["wait_seconds_num"] < 30)
    ].copy()

    abandoned_short_wait["wait_formatted"] = abandoned_short_wait[
        "wait_seconds_num"
    ].apply(
        lambda s: (
            f"{int(s // 3600):02d}:{int((s % 3600) // 60):02d}:{int(s % 60):02d}"
            if not pd.isna(s)
            else ""
        )
    )

    short_wait_export = abandoned_short_wait[
        [
            "customer_10",
            "event_time_normalized",
            "wait_formatted",
            "wait_seconds_num",
        ]
    ].rename(
        columns={
            "customer_10": "شماره_مشتری",
            "event_time_normalized": "زمان_تماس",
            "wait_formatted": "زمان_انتظار",
            "wait_seconds_num": "انتظار_ثانیه",
        }
    )
    short_wait_export.sort_values("انتظار_ثانیه", ascending=True, inplace=True)
    short_wait_export.reset_index(drop=True, inplace=True)

    # ═══════════════════════════════════════════════════════
    # شیت ۵: امتیاز Rate به تفکیک داخلی
    # ═══════════════════════════════════════════════════════

    if not rm_matched.empty and "rate_agent_ext" in rm_matched.columns:
        # نمره (score) از ستون rate اگر وجود دارد
        rate_score_col = None
        for candidate_col in ["rate_score", "score", "rate_rate", "امتیاز"]:
            if candidate_col in rm_matched.columns:
                rate_score_col = candidate_col
                break

        # اگر ستون امتیاز Rate در rate_df خام هست
        if rate_score_col is None:
            for candidate_col in ["score", "rate", "امتیاز"]:
                if candidate_col in rate_df.columns:
                    rate_score_col = candidate_col
                    break

        if rate_score_col and rate_score_col in rate_df.columns:
            rate_with_score = rate_df[["agent_ext", rate_score_col]].copy()
            rate_with_score[rate_score_col] = pd.to_numeric(
                rate_with_score[rate_score_col], errors="coerce"
            )

            rate_score_stats = (
                rate_with_score.groupby("agent_ext")
                .agg(
                    تعداد_نمره=( rate_score_col, "count"),
                    میانگین_نمره=(rate_score_col, "mean"),
                    حداقل_نمره=(rate_score_col, "min"),
                    حداکثر_نمره=(rate_score_col, "max"),
                )
                .reset_index()
                .rename(columns={"agent_ext": "داخلی"})
            )
            rate_score_stats["میانگین_نمره"] = rate_score_stats["میانگین_نمره"].round(2)
        else:
            # اگر ستون امتیاز وجود ندارد، از confidence استفاده کنیم
            rate_score_stats = (
                rm_matched.groupby("rate_agent_ext")
                .agg(
                    تعداد_مچ=("confidence", "count"),
                    میانگین_اطمینان=("confidence", "mean"),
                    حداقل_اطمینان=("confidence", "min"),
                    حداکثر_اطمینان=("confidence", "max"),
                )
                .reset_index()
                .rename(columns={"rate_agent_ext": "داخلی"})
            )
            rate_score_stats["میانگین_اطمینان"] = rate_score_stats[
                "میانگین_اطمینان"
            ].round(2)
    else:
        rate_score_stats = pd.DataFrame({"داخلی": [], "توضیح": ["داده‌ای موجود نیست"]})

    # ═══════════════════════════════════════════════════════
    # شیت ۶: توزیع زمان انتظار (Wait Distribution)
    # ═══════════════════════════════════════════════════════

    bins = [0, 10, 20, 30, 60, 120, 300, 600, 1800, 3600, float("inf")]
    labels = [
        "۰-۱۰ ثانیه",
        "۱۰-۲۰ ثانیه",
        "۲۰-۳۰ ثانیه",
        "۳۰ ثانیه-۱ دقیقه",
        "۱-۲ دقیقه",
        "۲-۵ دقیقه",
        "۵-۱۰ دقیقه",
        "۱۰-۳۰ دقیقه",
        "۳۰-۶۰ دقیقه",
        "بیش از ۶۰ دقیقه",
    ]

    mon["wait_bin"] = pd.cut(
        mon["wait_seconds_num"], bins=bins, labels=labels, right=False
    )

    wait_dist = (
        mon.groupby(["wait_bin", "status"], observed=False)
        .size()
        .unstack(fill_value=0)
        .reset_index()
        .rename(columns={"wait_bin": "بازه_انتظار"})
    )

    if "وصل شده" in wait_dist.columns and "رها شده" in wait_dist.columns:
        wait_dist["مجموع"] = wait_dist["وصل شده"] + wait_dist["رها شده"]
    else:
        wait_dist["مجموع"] = wait_dist.iloc[:, 1:].sum(axis=1)

    # ═══════════════════════════════════════════════════════
    # شیت ۷: داده خام مچ‌شده Monitoring ↔ Rate
    # ═══════════════════════════════════════════════════════

    raw_rate_matched = rm_matched.copy()

    # ═══════════════════════════════════════════════════════
    # نوشتن فایل اکسل
    # ═══════════════════════════════════════════════════════

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(
            writer, sheet_name="خلاصه_کلی", index=False
        )
        agent_merged.to_excel(
            writer, sheet_name="آمار_داخلی‌ها", index=False
        )
        long_wait_export.to_excel(
            writer, sheet_name="وصل_انتظار_بالای_۲دقیقه", index=False
        )
        short_wait_export.to_excel(
            writer, sheet_name="رها_انتظار_زیر_۳۰ثانیه", index=False
        )
        rate_score_stats.to_excel(
            writer, sheet_name="امتیاز_Rate_داخلی", index=False
        )
        wait_dist.to_excel(
            writer, sheet_name="توزیع_زمان_انتظار", index=False
        )
        raw_rate_matched.to_excel(
            writer, sheet_name="داده_خام_مچ_Rate", index=False
        )

        # ---- فرمت‌دهی عرض ستون‌ها ----
        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            for col_cells in ws.columns:
                max_len = 0
                col_letter = col_cells[0].column_letter
                for cell in col_cells:
                    try:
                        cell_len = len(str(cell.value)) if cell.value else 0
                        if cell_len > max_len:
                            max_len = cell_len
                    except Exception:
                        pass
                adjusted = min(max_len + 4, 50)
                ws.column_dimensions[col_letter].width = adjusted

    print(f"\n✅ فایل تحلیل ذخیره شد: {output_path}")
    print(f"   شیت‌ها: {list(pd.ExcelFile(output_path).sheet_names)}")

    return output_path
