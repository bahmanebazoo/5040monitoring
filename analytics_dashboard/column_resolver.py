"""
شناسایی هوشمند ستون‌ها — SRP
================================
enriched اکسل ممکنه ستون‌ها رو با نام‌های مختلف
(فارسی / انگلیسی / با prefix) داشته باشه.
"""

from __future__ import annotations
from typing import Optional, Sequence
import pandas as pd


class ColumnResolver:

    _MAP: dict[str, list[str]] = {
        "datetime": [
            "mon_datetime", "datetime", "تاریخ_زمان", "start_datetime",
            "mon_date", "call_datetime", "زمان_شروع",
            "event_time_normalized", "monitoring_time", "زمان",
        ],
        "agent": [
            "mon_agent", "agent", "اپراتور", "operator", "agent_name",
            "نام_اپراتور", "کارشناس", "agent_ext", "monitoring_agent_ext",
        ],
        "status": [
            "status", "وضعیت", "call_status", "mon_status",
            "monitoring_status",
        ],
        "support_match": [
            # ⭐ ستون‌های boolean — اول چک بشن
            "has_support_match",
            # ستون‌های وضعیت متنی
            "support_match_status", "match_status", "support_matched",
            "matched_support", "match_support", "support_status",
            "وضعیت_مچ_پشتیبانی", "وضعیت_تطبیق_پشتیبانی",
        ],
        "rate_match": [
            # ⭐ ستون boolean
            "has_rate_match",
            # ستون‌های متنی
            "rate_match_status", "rate_matched", "matched_rate",
            "match_rate_status", "وضعیت_مچ_نرخ",
        ],
        "confidence": [
            "confidence", "support_confidence", "match_confidence",
            "امتیاز_اطمینان", "امتیاز",
        ],
        "delta": [
            "delta_minutes", "support_delta_minutes", "time_delta",
            "فاصله_زمانی",
        ],
        "wait": [
            "wait_seconds", "wait_time", "زمان_انتظار", "queue_time",
            "monitoring_wait_seconds",
        ],
        "talk": [
            "talk_seconds", "talk_time", "مدت_مکالمه", "duration",
            "call_duration", "مدت_تماس", "duration_seconds",
            "rate_duration_seconds",
        ],
        "customer": [
            "customer_10", "customer", "شماره_مشتری", "phone",
            "caller_number", "شماره",
        ],
    }

    def __init__(self, columns: Sequence[str]):
        self._columns = list(columns)
        self._cache: dict[str, Optional[str]] = {}

    # ── resolve اصلی ──

    def resolve(self, key: str) -> Optional[str]:
        """برای کلید داخلی، اولین ستون موجود رو برمی‌گردونه."""
        if key in self._cache:
            return self._cache[key]

        candidates = self._MAP.get(key, [])
        for c in candidates:
            if c in self._columns:
                self._cache[key] = c
                return c

        # fallback: fuzzy
        for c in self._columns:
            cl = c.lower().replace(" ", "_")
            if key in cl:
                self._cache[key] = c
                return c

        self._cache[key] = None
        return None

    # ── resolve مخصوص مچینگ (رفع باگ صفر بودن) ──

    def resolve_match_column(
        self, key: str, df: pd.DataFrame
    ) -> tuple[Optional[str], int]:
        """
        ستون مچ رو پیدا کن + تعداد matched رو بشمار.
        اگه dedicated پیدا نشد → از match_status عمومی استفاده کن.
        """
        col = self.resolve(key)

        if col is not None and col in df.columns:
            count = self._count_matched(df[col])
            if count > 0:
                return col, count

        # fallback عمومی
        for fallback in ["match_status", "وضعیت_تطبیق"]:
            if fallback in df.columns:
                count = self._count_matched(df[fallback])
                if count > 0:
                    return fallback, count

        return col, 0

    @staticmethod
    def _count_matched(series: pd.Series) -> int:
        """
        ⭐ تعداد ردیف‌هایی که matched / TRUE / True / 1 هستن.

        پشتیبانی از:
          - بولی واقعی پایتون:    True / False
          - بولی پانداز:         pd.BooleanDtype
          - رشته‌ای:             "True", "TRUE", "true"
          - رشته‌ای:             "matched", "Matched", "MATCHED"
          - عددی:               1, 1.0
          - فارسی:              "بله"
        """
        if series.empty:
            return 0

        # ① اگه dtype واقعاً bool هست
        if pd.api.types.is_bool_dtype(series):
            return int(series.sum())

        # ② تبدیل به رشته و نرمال‌سازی
        s = series.fillna("").astype(str).str.lower().str.strip()

        return int(
            s.isin([
                "true", "matched", "1", "1.0",
                "yes", "بله",
            ]).sum()
        )

    # ── ابزارها ──

    def has(self, key: str) -> bool:
        return self.resolve(key) is not None

    def report(self) -> dict[str, Optional[str]]:
        return {k: self.resolve(k) for k in self._MAP}
