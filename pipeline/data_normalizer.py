"""STEP 2 — نرمال‌سازی شماره تلفن، زمان، اپراتور."""

import pandas as pd
from dataclasses import dataclass

from normalization.phone import normalize_phone
from normalization.datetime import parse_jalali_datetime
from normalization.agent import map_agent_name_to_ext

from .data_loader import RawData


@dataclass
class NormalizedData:
    """ظرف نگه‌داری داده‌های نرمال‌شده."""
    monitoring: pd.DataFrame
    rate: pd.DataFrame
    support: pd.DataFrame


class DataNormalizer:
    """نرمال‌سازی شماره، زمان، اپراتور روی هر سه دیتافریم."""

    def normalize(self, raw: RawData) -> NormalizedData:
        print("\n" + "=" * 60)
        print("STEP 2 — Normalizing")
        print("=" * 60)

        monitoring = raw.monitoring.copy()
        rate = raw.rate.copy()
        support = raw.support.copy()
        mapping = raw.mapping

        # ── شماره تلفن → 10 رقم ──
        monitoring["customer_10"] = monitoring["customer_raw"].apply(normalize_phone)
        rate["customer_10"] = rate["customer_raw"].apply(normalize_phone)
        support["customer_10"] = support["customer_raw"].apply(normalize_phone)

        # ── اپراتور Support: نام → داخلی ──
        support = map_agent_name_to_ext(support, mapping)

        # ── حذف بدون شماره ──
        monitoring.dropna(subset=["customer_10"], inplace=True)
        rate.dropna(subset=["customer_10"], inplace=True)
        support.dropna(subset=["customer_10"], inplace=True)

        # ── زمان جلالی → datetime ──
        monitoring["event_time"] = parse_jalali_datetime(monitoring["event_time"])
        support["event_time"] = parse_jalali_datetime(support["event_time"])

        if "connect_time_raw" in rate.columns:
            rate["connect_time_raw"] = (
                rate["connect_time_raw"]
                .astype(str)
                .str.replace("-", "/", regex=False)
            )
            rate["connect_time_raw"] = parse_jalali_datetime(rate["connect_time_raw"])

        # ── مدت‌ها → ثانیه ──
        monitoring["wait_seconds"] = pd.to_timedelta(
            monitoring["wait_time"]
        ).dt.total_seconds()

        rate["duration_seconds"] = pd.to_timedelta(
            rate["duration_time"]
        ).dt.total_seconds()

        # ── اصلاح agent_ext (حذف .0) ──
        for df in [monitoring, rate, support]:
            if "agent_ext" in df.columns:
                df["agent_ext"] = (
                    df["agent_ext"]
                    .astype(str)
                    .str.strip()
                    .str.replace(r"\.0$", "", regex=True)
                    .replace({"nan": None, "None": None, "NONE": None})
                )

        print(f"   Monitoring : {len(monitoring):,}")
        print(f"   Rate       : {len(rate):,}")
        print(f"   Support    : {len(support):,}")

        return NormalizedData(
            monitoring=monitoring,
            rate=rate,
            support=support,
        )
