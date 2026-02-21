"""
DashboardGenerator — ارکستراتور (DIP)
=======================================
به SheetCreator انتزاعی وابسته‌ست، نه به پیاده‌سازی خاص.
"""

from __future__ import annotations

import traceback
from datetime import datetime
from pathlib import Path
from typing import Optional

from openpyxl import Workbook

from .data_preparator import DataPreparator, PreparedData
from .sheets import build_default_creators
from .sheets.base import SheetCreator
from .style_manager import StyleManager
from .sheets.work_hours_sheet import WorkHoursSheetCreator


class DashboardGenerator:

    def __init__(
        self,
        preparator: DataPreparator | None = None,
        style: StyleManager | None = None,
        sheet_creators: list[SheetCreator] | None = None,
    ):
        self._preparator = preparator or DataPreparator()
        self._style = style or StyleManager()

        if sheet_creators is None:
            self._creators = build_default_creators(self._style)
        else:
            self._creators = sheet_creators

    def generate(
        self,
        enriched_path: str | Path,
        output_path: str | Path | None = None,
    ) -> Path:
        enriched_path = Path(enriched_path)

        if output_path is None:
            output_path = (
                enriched_path.parent
                / f"analytics_dashboard_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
        else:
            output_path = Path(output_path)

        output_path.parent.mkdir(parents=True, exist_ok=True)

        print("=" * 60)
        print("🚀 شروع تولید داشبورد تحلیلی (SOLID)")
        print("=" * 60)

        print("\n📥 خواندن داده غنی‌شده...")
        raw_df = self._preparator.load(enriched_path)

        print("\n🔧 آماده‌سازی داده...")
        data = self._preparator.prepare(raw_df)

        wb = Workbook()
        wb.remove(wb.active)

        print(f"\n📊 تولید {len(self._creators)} شیت:")
        for creator in self._creators:
            name = creator.sheet_name
            print(f"   📊 {name} ...", end=" ")
            try:
                creator.create(wb, data)
                print("✅")
            except Exception as exc:
                print(f"❌ {exc}")
                traceback.print_exc()

        print(f"\n💾 ذخیره: {output_path}")
        wb.save(str(output_path))

        print("\n" + "=" * 60)
        print(f"✅ داشبورد آماده — {len(self._creators)} شیت")
        print(f"   📁 {output_path}")
        print("=" * 60)

        return output_path
