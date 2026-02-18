import pandas as pd
from pathlib import Path


def export_matches(
    matches_df: pd.DataFrame,
    metrics,
    output_path: Path,
    as_excel: bool = True,
    sheet_prefix: str = "",
) -> None:
    """
    خروجی Excel/CSV با شیت‌های:
      - matched: ردیف‌های مچ‌شده
      - unmatched: ردیف‌هایی که مچ نشدند
      - summary: خلاصه آمار
      - confidence_breakdown: توزیع امتیازها
      - delta_stats: آمار فاصله زمانی
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # جدا کردن مچ‌شده و مچ‌نشده
    status_col = "match_status" if "match_status" in matches_df.columns else "status"
    matched_only = matches_df[matches_df[status_col] == "matched"].copy()
    unmatched_only = matches_df[matches_df[status_col] != "matched"].copy()

    if as_excel:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            # شیت مچ‌شده‌ها
            matched_only.to_excel(
                writer, index=False, sheet_name=f"{sheet_prefix}matched"
            )

            # شیت مچ‌نشده‌ها
            if not unmatched_only.empty:
                unmatched_only.to_excel(
                    writer, index=False, sheet_name=f"{sheet_prefix}unmatched"
                )

            # شیت خلاصه
            total_mon = metrics.total_monitoring if hasattr(metrics, "total_monitoring") else len(matches_df)
            summary_df = pd.DataFrame(
                {
                    "metric": [
                        "total_monitoring",
                        "total_source",
                        "total_matched",
                        "total_unmatched",
                        "match_rate",
                    ],
                    "value": [
                        total_mon,
                        getattr(metrics, "total_support", getattr(metrics, "total_rate", 0)),
                        len(matched_only),
                        len(unmatched_only),
                        f"{len(matched_only) / max(total_mon, 1) * 100:.1f}%",
                    ],
                }
            )
            summary_df.to_excel(
                writer, index=False, sheet_name=f"{sheet_prefix}summary"
            )

            # شیت توزیع امتیاز
            if not matched_only.empty and "confidence" in matched_only.columns:
                conf_bk = (
                    matched_only["confidence"]
                    .value_counts()
                    .sort_index()
                    .reset_index()
                )
                conf_bk.columns = ["confidence", "count"]
                conf_bk.to_excel(
                    writer,
                    index=False,
                    sheet_name=f"{sheet_prefix}confidence_bk",
                )

            # شیت آمار delta
            if not matched_only.empty and "delta_minutes" in matched_only.columns:
                delta_stats = (
                    matched_only["delta_minutes"].describe().reset_index()
                )
                delta_stats.columns = ["stat", "value"]
                delta_stats.to_excel(
                    writer,
                    index=False,
                    sheet_name=f"{sheet_prefix}delta_stats",
                )
    else:
        matches_df.to_csv(output_path, index=False)
