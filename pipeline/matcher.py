"""STEP 4 — اجرای مچینگ Support و Rate."""

import pandas as pd
from dataclasses import dataclass

from matching.monitoring_support import MonitoringSupportMatcher
from matching.monitoring_rate import MonitoringRateMatcher

from .time_synchronizer import SyncedData


@dataclass
class MatchResult:
    """خروجی مچینگ‌ها."""
    support_matches: pd.DataFrame
    rate_matches: pd.DataFrame
    connected_monitoring: pd.DataFrame  # فقط وصل‌شده‌ها (برای ارجاع بعدی)


class Matcher:
    """اجرای هر دو مچینگ Support و Rate — هر کدام فقط یک بار."""

    def __init__(self, support_threshold: int = 75, rate_threshold: int = 85):
        self.support_threshold = support_threshold
        self.rate_threshold = rate_threshold

    def run(self, synced: SyncedData) -> MatchResult:
        print("\n" + "=" * 60)
        print("STEP 4 — Matching")
        print("=" * 60)

        mon_connected = synced.monitoring[
            synced.monitoring["status"] == "وصل شده"
        ].copy()
        print(f"   Connected calls: {len(mon_connected):,}")

        # ── 4a  Support ──
        print(f"\n   ── 4a  Support (threshold={self.support_threshold}) ──")
        support_matches = MonitoringSupportMatcher(
            threshold=self.support_threshold
        ).match(
            monitoring_df=mon_connected,
            support_df=synced.support,
        )
        n_sup = int((support_matches["match_status"] == "matched").sum())
        print(f"   ✅ Support matched: {n_sup:,} / {len(mon_connected):,} "
              f"({n_sup / len(mon_connected):.1%})")

        # ── 4b  Rate ──
        print(f"\n   ── 4b  Rate (threshold={self.rate_threshold}) ──")
        rate_matches = MonitoringRateMatcher(
            threshold=self.rate_threshold
        ).match(
            monitoring_df=mon_connected,
            rate_df=synced.rate,
        )
        n_rate = int((rate_matches["match_status"] == "matched").sum())
        print(f"   ✅ Rate matched:    {n_rate:,} / {len(mon_connected):,} "
              f"({n_rate / len(mon_connected):.1%})")

        # ── توزیع confidence Rate ──
        print("\n   Rate confidence distribution:")
        dist = rate_matches["confidence"].value_counts().sort_index()
        for conf_val, cnt in dist.items():
            print(f"      {conf_val:>3}: {cnt:>5,}")

        return MatchResult(
            support_matches=support_matches,
            rate_matches=rate_matches,
            connected_monitoring=mon_connected,
        )
