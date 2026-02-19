"""STEP 3 — تخمین offset بین Support و Monitoring و نرمال‌سازی زمان."""

import pandas as pd
from dataclasses import dataclass

from synchronization.offset_estimator import SupportOffsetEstimator
from synchronization.time_normalizer import TimeNormalizer

from .data_normalizer import NormalizedData


@dataclass
class SyncedData:
    """داده‌ها پس از همگام‌سازی زمانی."""
    monitoring: pd.DataFrame
    rate: pd.DataFrame
    support: pd.DataFrame
    offset_minutes: float


class TimeSynchronizer:
    """تخمین offset و اعمال نرمال‌سازی زمانی."""

    def synchronize(self, data: NormalizedData) -> SyncedData:
        print("\n" + "=" * 60)
        print("STEP 3 — Time synchronization")
        print("=" * 60)

        monitoring = data.monitoring.copy()
        rate = data.rate.copy()
        support = data.support.copy()

        # ── تخمین offset ──
        offset_minutes = SupportOffsetEstimator().estimate(
            monitoring.rename(columns={"event_time": "monitoring_time"}),
            support.rename(columns={"event_time": "support_time"}),
        )
        print(f"   Support offset: {offset_minutes} min")

        # ── اعمال نرمال‌سازی ──
        monitoring = TimeNormalizer(0).apply(monitoring, "event_time")
        support = TimeNormalizer(offset_minutes).apply(support, "event_time")

        if "connect_time_raw" in rate.columns:
            rate = TimeNormalizer(0).apply(rate, "connect_time_raw")
            rate.rename(
                columns={"connect_time_raw_normalized": "event_time_normalized"},
                inplace=True,
            )

        # ── گزارش بازه‌ها ──
        print(f"   Mon  : {monitoring['event_time_normalized'].min()} → "
              f"{monitoring['event_time_normalized'].max()}")
        print(f"   Sup  : {support['event_time_normalized'].min()} → "
              f"{support['event_time_normalized'].max()}")
        print(f"   Rate : {rate['event_time_normalized'].min()} → "
              f"{rate['event_time_normalized'].max()}")

        return SyncedData(
            monitoring=monitoring,
            rate=rate,
            support=support,
            offset_minutes=offset_minutes,
        )
