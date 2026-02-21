"""
شیت تحلیل اپراتور
"""

from __future__ import annotations

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

from ..data_preparator import PreparedData
from .base import SheetCreator


class AgentSheetCreator(SheetCreator):

    @property
    def sheet_name(self) -> str:
        return "تحلیل_اپراتور"

    def create(self, wb: Workbook, data: PreparedData) -> None:
        ws = wb.create_sheet(self.sheet_name)
        ws.sheet_view.rightToLeft = True

        df = data.df

        if "_agent" not in df.columns:
            ws.cell(row=1, column=1, value="داده اپراتور موجود نیست")
            return

        stats = df.groupby("_agent").agg(
            تعداد=("_agent", "size")
        ).reset_index()
        stats.rename(columns={"_agent": "اپراتور"}, inplace=True)

        # ❌ حذف شد: وصل_شده، نرخ_پاسخ_%  (وصل شدن ربطی به اپراتور نداره)
        # ❌ حذف شد: میانگین_انتظار          (زمان انتظار ربطی به پشتیبان نداره)

        if "_talk_seconds" in df.columns:
            ta = df.groupby("_agent")["_talk_seconds"].mean()
            stats["میانگین_مکالمه"] = stats["اپراتور"].map(ta).round(1)

        s_mask = self._match_mask(df, data.support_match_col)
        if s_mask.any():
            sm = df[s_mask].groupby("_agent").size()
            stats["مچ_پشتیبانی"] = (
                stats["اپراتور"].map(sm).fillna(0).astype(int)
            )

        stats.sort_values("تعداد", ascending=False, inplace=True)
        self._write_table(ws, stats)
        self.style.auto_fit(ws)

        ncols = len(stats.columns)
        n = min(len(stats), 30)

        bar = BarChart()
        bar.type = "col"
        bar.title = "تعداد تماس به تفکیک اپراتور"
        bar.style = 10
        bar.width, bar.height = 30, 15
        bar.add_data(
            Reference(ws, min_col=2, min_row=1, max_row=n + 1),
            titles_from_data=True,
        )
        bar.set_categories(
            Reference(ws, min_col=1, min_row=2, max_row=n + 1)
        )
        ws.add_chart(bar, f"{get_column_letter(ncols + 2)}1")

        # ❌ حذف شد: نمودار راداری نرخ_پاسخ_%
