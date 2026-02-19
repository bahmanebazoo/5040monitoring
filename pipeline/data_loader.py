"""STEP 1 — لود تمام شیت‌ها از اکسل."""

import pandas as pd
from dataclasses import dataclass

from loaders.monitoring_loader import MonitoringLoader
from loaders.rate_loader import RateLoader
from loaders.support_loader import SupportLoader
from loaders.mapping_loader import MappingLoader


@dataclass
class RawData:
    """ظرف نگه‌داری داده‌های خام لود شده."""
    monitoring: pd.DataFrame
    rate: pd.DataFrame
    support: pd.DataFrame
    mapping: pd.DataFrame


class DataLoader:
    """لود تمام داده‌ها از یک فایل اکسل."""

    def load(self, path: str) -> RawData:
        print("=" * 60)
        print("STEP 1 — Loading data")
        print("=" * 60)

        monitoring = MonitoringLoader().load(path)
        rate = RateLoader().load(path)
        support = SupportLoader().load(path)
        mapping = MappingLoader().load(path)

        print(f"   Monitoring : {len(monitoring):,}")
        print(f"   Rate       : {len(rate):,}")
        print(f"   Support    : {len(support):,}")
        print(f"   Mapping    : {len(mapping):,}")

        return RawData(
            monitoring=monitoring,
            rate=rate,
            support=support,
            mapping=mapping,
        )
