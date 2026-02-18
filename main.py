import pandas as pd

# Loaders
from loaders.monitoring_loader import MonitoringLoader
from loaders.rate_loader import RateLoader
from loaders.support_loader import SupportLoader
from loaders.mapping_loader import MappingLoader

# Normalization
from normalization.phone import normalize_phone
from normalization.datetime import parse_jalali_datetime
from normalization.agent import map_agent_name_to_ext

# Synchronization & Matching
from synchronization.offset_estimator import SupportOffsetEstimator
from synchronization.time_normalizer import TimeNormalizer
from matching.monitoring_support import MonitoringSupportMatcher

MONITORING_PATH = "1 الی 23 بهمن مونیتورینگ.xlsx"
OUTPUT_PATH = "matched_report.xlsx"

def compute_quality_metrics(monitoring_df, support_df, matches_df):
    total_monitoring = len(monitoring_df)
    total_support = len(support_df)
    total_matches = len(matches_df)
    match_rate = total_matches / total_monitoring if total_monitoring else 0

    confidence_breakdown = (
        matches_df["confidence"]
        .fillna("unknown")
        .value_counts(normalize=True)
        .rename("ratio")
        .reset_index()
        .rename(columns={"index": "confidence"})
    )

    delta_series = matches_df["delta_minutes"].dropna()
    if delta_series.empty:
        delta_stats = pd.DataFrame(columns=["metric", "value"])
    else:
        delta_stats = pd.DataFrame({
            "metric": [
                "mean", "median", "std", "min", "max",
                "p10", "p25", "p50", "p75", "p90", "p95"
            ],
            "value": [
                delta_series.mean(),
                delta_series.median(),
                delta_series.std(),
                delta_series.min(),
                delta_series.max(),
                delta_series.quantile(0.10),
                delta_series.quantile(0.25),
                delta_series.quantile(0.50),
                delta_series.quantile(0.75),
                delta_series.quantile(0.90),
                delta_series.quantile(0.95)
            ]
        })

    summary = pd.DataFrame({
        "metric": ["total_monitoring", "total_support", "total_matches", "match_rate"],
        "value": [total_monitoring, total_support, total_matches, match_rate]
    })

    return summary, confidence_breakdown, delta_stats


def export_to_excel(matches_df, summary_df, confidence_df, delta_stats_df, output_path):
    with pd.ExcelWriter(output_path) as writer:
        matches_df.to_excel(writer, index=False, sheet_name="matches")
        summary_df.to_excel(writer, index=False, sheet_name="summary")
        confidence_df.to_excel(writer, index=False, sheet_name="confidence_breakdown")
        delta_stats_df.to_excel(writer, index=False, sheet_name="delta_stats")


def main():
    # ----------------------------------------------------
    # 1. LOAD DATA
    # ----------------------------------------------------
    print("1. Loading Data...")
    monitoring_df = MonitoringLoader().load(MONITORING_PATH)
    rate_df = RateLoader().load(MONITORING_PATH)
    support_df = SupportLoader().load(MONITORING_PATH)
    mapping_df = MappingLoader().load(MONITORING_PATH)  # اگر در Loader هاردکد شده، باید خودش به exp بخورد

    # ----------------------------------------------------
    # 2. DATA NORMALIZATION
    # ----------------------------------------------------
    print("2. Normalizing Data (Phone and Time)...")

    monitoring_df["customer_10"] = monitoring_df["customer_raw"].apply(normalize_phone)
    rate_df["customer_10"] = rate_df["customer_raw"].apply(normalize_phone)
    support_df["customer_10"] = support_df["customer_raw"].apply(normalize_phone)

    support_df = map_agent_name_to_ext(support_df, mapping_df)

    monitoring_df.dropna(subset=["customer_10"], inplace=True)
    rate_df.dropna(subset=["customer_10"], inplace=True)
    support_df.dropna(subset=["customer_10"], inplace=True)

    monitoring_df["event_time"] = parse_jalali_datetime(monitoring_df["event_time"])
    support_df["event_time"] = parse_jalali_datetime(support_df["event_time"])

    if "connect_time_raw" in rate_df.columns:
        # تبدیل تاریخ 1404-11-23 به 1404/11/23 (فقط برای Rate)
        rate_df["connect_time_raw"] = (
            rate_df["connect_time_raw"]
            .astype(str)
            .str.replace("-", "/", regex=False)
        )

        rate_df["connect_time_raw"] = parse_jalali_datetime(rate_df["connect_time_raw"])

    monitoring_df["wait_seconds"] = pd.to_timedelta(
        monitoring_df["wait_time"]
    ).dt.total_seconds()

    rate_df["duration_seconds"] = pd.to_timedelta(
        rate_df["duration_time"]
    ).dt.total_seconds()

    # ----------------------------------------------------
    # 3. TIME SYNCHRONIZATION
    # ----------------------------------------------------
    print("3. Estimating Time Offset between Support and Monitoring...")
    offset_minutes = SupportOffsetEstimator().estimate(
        monitoring_df.rename(columns={"event_time": "monitoring_time"}),
        support_df.rename(columns={"event_time": "support_time"})
    )

    print(f"   -> Estimated Support Offset: {offset_minutes} minutes")

    monitoring_df = TimeNormalizer(0).apply(monitoring_df, "event_time")
    support_df = TimeNormalizer(offset_minutes).apply(support_df, "event_time")

    if "connect_time_raw" in rate_df.columns:
        rate_df = TimeNormalizer(0).apply(rate_df, "connect_time_raw")

    # ----------------------------------------------------
    # 4. MATCHING
    # ----------------------------------------------------
    print("4. Running Matching Algorithm (Monitoring <-> Support)...")
    mon_for_matching = monitoring_df[monitoring_df["status"] == "وصل شده"].copy()

    matches_df = MonitoringSupportMatcher().match(
        monitoring_df=mon_for_matching,
        support_df=support_df
    )

    final_matches = matches_df[matches_df["confidence"] >= 70]

    print("\n--- Results (Top 5 Matches) ---")
    print(final_matches[["customer_10", "monitoring_time", "support_time", "delta_minutes", "confidence"]].head())

    # ----------------------------------------------------
    # 5. QUALITY ANALYTICS & EXPORT
    # ----------------------------------------------------
    print("\n5. Generating Quality Report and Exporting...")
    summary_df, confidence_df, delta_stats_df = compute_quality_metrics(
        monitoring_df=mon_for_matching,
        support_df=support_df,
        matches_df=matches_df
    )

    export_to_excel(
        matches_df=matches_df,
        summary_df=summary_df,
        confidence_df=confidence_df,
        delta_stats_df=delta_stats_df,
        output_path=OUTPUT_PATH
    )

    print(f"Exported report to: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
