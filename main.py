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


def main():
    # ----------------------------------------------------
    # 1. LOAD DATA
    # ----------------------------------------------------
    print("1. Loading Data...")
    monitoring_df = MonitoringLoader().load(MONITORING_PATH)
    rate_df = RateLoader().load(MONITORING_PATH)
    support_df = SupportLoader().load(MONITORING_PATH)
    mapping_df = MappingLoader().load(MONITORING_PATH)

    # ----------------------------------------------------
    # 2. DATA NORMALIZATION (Shared Logic)
    # ----------------------------------------------------

    # 2.1 Normalize Phone Numbers (All DFs)
    print("2. Normalizing Data (Phone and Time)...")

    monitoring_df["customer_10"] = monitoring_df["customer_raw"].apply(normalize_phone)
    rate_df["customer_10"] = rate_df["customer_raw"].apply(normalize_phone)
    support_df["customer_10"] = support_df["customer_raw"].apply(normalize_phone)

    # 2.2 Map Agent Names to Extensions (Support DF)
    support_df = map_agent_name_to_ext(support_df, mapping_df)

    # Drop rows where phone number could not be extracted (e.g., ********)
    monitoring_df.dropna(subset=["customer_10"], inplace=True)
    rate_df.dropna(subset=["customer_10"], inplace=True)
    support_df.dropna(subset=["customer_10"], inplace=True)

    # 2.3 Parse Datetime and Calculate Durations (Monitoring & Rate)
    monitoring_df["event_time"] = parse_jalali_datetime(monitoring_df["event_time"])
    support_df["event_time"] = parse_jalali_datetime(support_df["event_time"])

    # Calculate Wait Seconds for Monitoring
    monitoring_df["wait_seconds"] = pd.to_timedelta(
        monitoring_df["wait_time"]
    ).dt.total_seconds()

    # Calculate Duration Seconds for Rate (assuming rate duration is also H:MM:SS)
    rate_df["duration_seconds"] = pd.to_timedelta(
        rate_df["duration_time"]
    ).dt.total_seconds()

    # ----------------------------------------------------
    # 3. TIME SYNCHRONIZATION (The core logic to enable/disable)
    # ----------------------------------------------------

    # 3.1 Estimate Offset (using Monitoring as base and Support as target)
    print("3. Estimating Time Offset between Support and Monitoring...")
    offset_minutes = SupportOffsetEstimator().estimate(
        monitoring_df.rename(columns={"event_time": "monitoring_time"}),
        support_df.rename(columns={"event_time": "support_time"})
    )

    print(f"   -> Estimated Support Offset: {offset_minutes} minutes")

    # 3.2 Apply Synchronization (normalize all times based on offset)
    normalizer = TimeNormalizer(offset_minutes=offset_minutes)

    # Monitoring is the base, so offset is 0
    monitoring_df = TimeNormalizer(0).apply(monitoring_df, "event_time")

    # Apply offset to support time
    support_df = normalizer.apply(support_df, "event_time")

    # NOTE: Rate offset will be handled later, for now we assume Rate is close to Monitoring
    rate_df = TimeNormalizer(0).apply(rate_df, "connect_time_raw")  # Rename column later

    # ----------------------------------------------------
    # 4. MATCHING (The main task)
    # ----------------------------------------------------
    print("4. Running Matching Algorithm (Monitoring <-> Support)...")

    # We only match answered calls in Monitoring (or calls that generated support activity)
    mon_for_matching = monitoring_df[monitoring_df["status"] == "وصل شده"].copy()

    matches_df = MonitoringSupportMatcher().match(
        monitoring_df=mon_for_matching,
        support_df=support_df
    )

    # Filter for high confidence matches (e.g., score >= 70)
    final_matches = matches_df[matches_df["confidence"] >= 70]

    print("\n--- Results (Top 5 Matches) ---")
    print(final_matches[["customer_10", "monitoring_time", "support_time", "delta_minutes", "confidence"]].head())

    # ----------------------------------------------------
    # 5. ANALYTICS (Test one simple metric)
    # ----------------------------------------------------
    print("\n--- Basic Analytics ---")

    # Use normalized data for clean calculation
    avg_wait_time = monitoring_df["wait_seconds"].mean() / 60
    print(f"Average Wait Time (Uncorrected): {avg_wait_time:.2f} minutes")

    # You can now proceed to merge final_matches back into monitoring_df
    # to create the Canonical Event Model for deeper analysis.


if __name__ == "__main__":
    main()
