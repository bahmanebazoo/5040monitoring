"""
5040 Monitoring — Main Pipeline
================================
main فقط orchestrate می‌کنه. هیچ بیزینس لاجیکی اینجا نیست.
"""

from config.config import (
    MONITORING_PATH,
    MATCHED_REPORT_PATH,
    ENRICHED_PATH,
    ANALYTICS_PATH,
    SUPPORT_MATCH_THRESHOLD,
    RATE_MATCH_THRESHOLD,
)

from pipeline import DataLoader, DataNormalizer, TimeSynchronizer, Matcher
from sheets import MatchedReportExporter, generate_enriched_excel, generate_analytics_excel
from analytics_dashboard import generate_analytics_dashboard
from pathlib import Path


def main():
    # STEP 1 — Load
    raw = DataLoader().load(MONITORING_PATH)

    # STEP 2 — Normalize
    normalized = DataNormalizer().normalize(raw)

    # STEP 3 — Time sync
    synced = TimeSynchronizer().synchronize(normalized)

    # STEP 4 — Match (Support + Rate, هر کدام فقط یک بار)
    match_result = Matcher(
        support_threshold=SUPPORT_MATCH_THRESHOLD,
        rate_threshold=RATE_MATCH_THRESHOLD,
    ).run(synced)

    # STEP 5 — Matched report Excel
    MatchedReportExporter().export(
        support_matches=match_result.support_matches,
        rate_matches=match_result.rate_matches,
        output_path=MATCHED_REPORT_PATH,
    )

    # STEP 6 — Enriched monitoring Excel
    print("\n" + "=" * 60)
    print("STEP 6 — Enriched Monitoring Excel")
    print("=" * 60)
    try:
        path = generate_enriched_excel(
            monitoring_df=synced.monitoring,
            support_matches_df=match_result.support_matches,
            rate_matches_df=match_result.rate_matches,
            output_path=ENRICHED_PATH,
        )
        print(f"   ✅ {path}")
    except Exception as e:
        print(f"   ❌ Failed: {e}")

    # STEP 7 — Analytics report Excel
    print("\n" + "=" * 60)
    print("STEP 7 — Analytics Report")
    print("=" * 60)
    try:
        path = generate_analytics_excel(
            monitoring_df=synced.monitoring,
            rate_df=synced.rate,
            support_matches_df=match_result.support_matches,
            rate_matches_df=match_result.rate_matches,
            output_path=ANALYTICS_PATH,
        )
        print(f"   ✅ {path}")
    except Exception as e:
        print(f"   ❌ Failed: {e}")

    # STEP 8 — Analytics Dashboard with Charts
    print("\n" + "=" * 60)
    print("STEP 8 — Analytics Dashboard with Charts")
    print("=" * 60)
    try:
        dashboard_path = generate_analytics_dashboard(
            enriched_path=ENRICHED_PATH,
            output_path=Path("analytics_dashboard.xlsx"),
        )
        print(f"   ✅ {dashboard_path}")
    except Exception as e:
        print(f"   ❌ Failed: {e}")

    # Done
    print("\n" + "=" * 60)
    print("✅ ALL DONE")
    print(f"   📊 {MATCHED_REPORT_PATH}  — 9 sheets")
    print(f"   📋 {ENRICHED_PATH}         — enriched raw data")
    print(f"   📈 {ANALYTICS_PATH}        — analytics dashboard")
    print("=" * 60)


if __name__ == "__main__":
    main()
