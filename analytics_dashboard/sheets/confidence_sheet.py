"""
شیت تحلیل امتیاز اعتماد (Confidence)
"""

from __future__ import annotations

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList

from ..data_preparator import PreparedData
from .base import SheetCreator


class ConfidenceSheetCreator(SheetCreator):

    @property
    def sheet_name(self) -> str:
        return "تحلیل_اعتماد"

    def create(self, wb: Workbook, data: PreparedData) -> None:
        ws = wb.create_sheet(self.sheet_name)
        ws.sheet_view.rightToLeft = True

        df = data.df

        if "_confidence" not in df.columns:
            ws.cell(row=1, column=1, value="داده Confidence موجود نیست")
            return

        cs = df["_confidence"].dropna()
        if cs.empty:
            ws.cell(row=1, column=1, value="Confidence خالی")
            return

        vc = cs.value_counts().sort_index().reset_index()
        vc.columns = ["امتیاز", "تعداد"]
        vc["درصد"] = (vc["تعداد"] / vc["تعداد"].sum() * 100).round(1)

        self._write_table(ws, vc)
        self.style.auto_fit(ws)
        n = len(vc)

        # Bar
        bar = BarChart()
        bar.title = "توزیع امتیاز اعتماد"
        bar.style = 10
        bar.width, bar.height = 25, 14
        bar.add_data(
            Reference(ws, min_col=2, min_row=1, max_row=n + 1),
            titles_from_data=True,
        )
        bar.set_categories(
            Reference(ws, min_col=1, min_row=2, max_row=n + 1)
        )
        ws.add_chart(bar, "E1")

        # Pie
        pie = PieChart()
        pie.title = "سهم هر سطح اعتماد"
        pie.style = 10
        pie.width, pie.height = 18, 12
        pie.add_data(
            Reference(ws, min_col=2, min_row=1, max_row=n + 1),
            titles_from_data=True,
        )
        pie.set_categories(
            Reference(ws, min_col=1, min_row=2, max_row=n + 1)
        )
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showPercent = True
        ws.add_chart(pie, "E28")

        # آمار توصیفی
        if pd.api.types.is_numeric_dtype(cs):
            sr = n + 4
            desc = {
                "میانگین": cs.mean(),
                "میانه": cs.median(),
                "انحراف معیار": cs.std(),
                "حداقل": cs.min(),
                "حداکثر": cs.max(),
            }
            ws.cell(row=sr, column=1, value="آمار")
            ws.cell(row=sr, column=2, value="مقدار")
            self.style.apply_header(ws, row=sr, max_col=2)
            for i, (k, v) in enumerate(desc.items(), sr + 1):
                ws.cell(row=i, column=1, value=k)
                ws.cell(row=i, column=2, value=round(v, 2))
