"""
آماده‌سازی داده — SRP
=====================
بارگذاری enriched + ساختن ستون‌های کمکی _hour, _agent, ...
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pandas as pd

from .column_resolver import ColumnResolver


@dataclass
class PreparedData:
    """داده آماده‌شده برای شیت‌ساز‌ها."""

    df: pd.DataFrame
    resolver: ColumnResolver
    total: int
    connected_count: int
    support_match_col: Optional[str]
    support_match_count: int
    rate_match_col: Optional[str]
    rate_match_count: int


class DataPreparator:

    # ── بارگذاری ──

    def load(self, enriched_path: Path) -> pd.DataFrame:
        enriched_path = Path(enriched_path)
        if not enriched_path.exists():
            raise FileNotFoundError(
                f"فایل enriched پیدا نشد: {enriched_path}"
            )

        for sheet_name in [
            "داده_خام_غنی شده",
            "داده_خام_غنی_شده",
            "enriched",
            None,
        ]:
            try:
                df = pd.read_excel(
                    enriched_path,
                    sheet_name=0 if sheet_name is None else sheet_name,
                )
                print(
                    f"   ✅ خوانده شد: "
                    f"{len(df)} ردیف × {len(df.columns)} ستون"
                )
                return df
            except Exception:
                continue

        raise ValueError("هیچ شیت مناسبی پیدا نشد")

    # ── آماده‌سازی ──

    def prepare(self, df: pd.DataFrame) -> PreparedData:
        df = df.copy()
        resolver = ColumnResolver(df.columns)

        self._print_column_report(resolver)
        self._add_datetime_cols(df, resolver)
        self._add_agent_col(df, resolver)
        self._add_status_col(df, resolver)
        self._add_numeric_col(df, resolver, "confidence", "_confidence")
        self._add_numeric_col(df, resolver, "delta", "_delta_minutes")
        self._add_numeric_col(df, resolver, "wait", "_wait_seconds")
        self._add_numeric_col(df, resolver, "talk", "_talk_seconds")

        # ⭐ ستون ساعت کاری — بعد از datetime و قبل از آمار
        self._add_work_hours_col(df)

        connected = self._count_connected(df)
        s_col, s_count = resolver.resolve_match_column("support_match", df)
        r_col, r_count = resolver.resolve_match_column("rate_match", df)

        # ⭐ آمار تایم کاری
        in_wh = int(df["_in_work_hours"].sum()) if "_in_work_hours" in df.columns else 0
        out_wh = len(df) - in_wh

        print(f"\n   📊 آمار اولیه:")
        print(f"      کل تماس‌ها:       {len(df)}")
        print(f"      وصل‌شده:          {connected}")
        print(f"      مچ پشتیبانی:      {s_count}  (ستون: {s_col})")
        print(f"      مچ نرخ‌دهی:       {r_count}  (ستون: {r_col})")
        print(f"      ⏰ داخل تایم کاری: {in_wh}")
        print(f"      🌙 خارج تایم کاری: {out_wh}")

        return PreparedData(
            df=df,
            resolver=resolver,
            total=len(df),
            connected_count=connected,
            support_match_col=s_col,
            support_match_count=s_count,
            rate_match_col=r_col,
            rate_match_count=r_count,
        )

    # ── private helpers ──

    @staticmethod
    def _print_column_report(resolver: ColumnResolver):
        report = resolver.report()
        print("\n   🔍 شناسایی ستون‌ها:")
        for k, v in report.items():
            status = f"✅ {v}" if v else "❌ یافت نشد"
            print(f"      {k:20s} → {status}")

    @staticmethod
    def _add_datetime_cols(df: pd.DataFrame, resolver: ColumnResolver):
        dt_col = resolver.resolve("datetime")
        if not dt_col:
            return
        df[dt_col] = pd.to_datetime(df[dt_col], errors="coerce")
        valid = df[dt_col].notna()
        if not valid.any():
            return
        df.loc[valid, "_hour"] = df.loc[valid, dt_col].dt.hour.astype(int)
        df.loc[valid, "_date"] = df.loc[valid, dt_col].dt.date
        df.loc[valid, "_day_of_week"] = df.loc[valid, dt_col].dt.day_name()
        df.loc[valid, "_weekday_num"] = df.loc[valid, dt_col].dt.weekday
        df.loc[valid, "_month"] = df.loc[valid, dt_col].dt.month

    @staticmethod
    def _add_agent_col(df: pd.DataFrame, resolver: ColumnResolver):
        col = resolver.resolve("agent")
        if col:
            df["_agent"] = df[col].fillna("نامشخص").astype(str)

    @staticmethod
    def _add_status_col(df: pd.DataFrame, resolver: ColumnResolver):
        col = resolver.resolve("status")
        if col:
            df["_status"] = df[col].fillna("نامشخص").astype(str)

    @staticmethod
    def _add_numeric_col(
        df: pd.DataFrame,
        resolver: ColumnResolver,
        key: str,
        target: str,
    ):
        col = resolver.resolve(key)
        if col:
            df[target] = pd.to_numeric(df[col], errors="coerce")

    @staticmethod
    def _count_connected(df: pd.DataFrame) -> int:
        if "_status" not in df.columns:
            return 0
        return int(
            df["_status"]
            .str.contains(
                "وصل|connected|answered|ANSWERED", case=False, na=False
            )
            .sum()
        )

    # ──────────────────────────────────────────────
    # ⭐ ستون ساعت کاری
    # ──────────────────────────────────────────────
    #   شنبه تا چهارشنبه:  8:00 – 18:00
    #   پنجشنبه:           8:00 – 14:00
    #   جمعه:              تعطیل (خارج تایم)
    #
    #   _weekday_num (پانداز):
    #     Monday=0 … Saturday=5, Sunday=6
    #
    #   تقویم ایرانی → میلادی:
    #     شنبه    = Saturday  = 5
    #     یکشنبه  = Sunday    = 6
    #     دوشنبه  = Monday    = 0
    #     سه‌شنبه = Tuesday   = 1
    #     چهارشنبه= Wednesday = 2
    #     پنجشنبه = Thursday  = 3
    #     جمعه    = Friday    = 4
    # ──────────────────────────────────────────────

    @staticmethod
    def _add_work_hours_col(df: pd.DataFrame):
        """ستون بولی _in_work_hours را اضافه می‌کند."""

        if "_weekday_num" not in df.columns or "_hour" not in df.columns:
            df["_in_work_hours"] = False
            return

        weekday = df["_weekday_num"]  # 0=Mon … 6=Sun
        hour = df["_hour"]

        # شنبه–چهارشنبه (Sat=5, Sun=6, Mon=0, Tue=1, Wed=2) → 8–18
        sat_to_wed = weekday.isin([5, 6, 0, 1, 2])
        in_sat_wed = sat_to_wed & (hour >= 8) & (hour < 18)

        # پنجشنبه (Thu=3) → 8–14
        thu = weekday == 3
        in_thu = thu & (hour >= 8) & (hour < 14)

        # جمعه (Fri=4) → خارج تایم
        # هیچ شرطی نداره

        df["_in_work_hours"] = in_sat_wed | in_thu

        # ⭐ لیبل فارسی هم اضافه کن (برای نمایش در شیت)
        df["_work_hours_label"] = df["_in_work_hours"].map(
            {True: "داخل تایم", False: "خارج تایم"}
        )
