"""
شیت تحلیل وضعیت تماس‌ها
"""

from __future__ import annotations

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList

from ..data_preparator import PreparedData
from .base import SheetCreator


class StatusSheetCreator(SheetCreator):

    @property
    def sheet_name(self) -> str:
        return "تحلیل_وضعیت"

    def create(self, wb: Workbook, data: PreparedData) -> None:
        ws = wb.create_sheet(self.sheet_name)
        ws.sheet_view.rightToLeft = True

        df = data.df

        if "_status" not in df.columns:
            ws.cell(row=1, column=1, value="داده وضعیت موجود نیست")
            return

        vc = df["_status"].value_counts().reset_index()
        vc.columns = ["وضعیت", "تعداد"]
        vc["درصد"] = (vc["تعداد"] / vc["تعداد"].sum() * 100).round(1)

        self._write_table(ws, vc)
        self.style.auto_fit(ws)
        n = len(vc)

        # Pie
        pie = PieChart()
        pie.title = "وضعیت تماس‌ها"
        pie.style = 10
        pie.width, pie.height = 20, 14
        pie.add_data(
            Reference(ws, min_col=2, min_row=1, max_row=n + 1),
            titles_from_data=True,
        )
        pie.set_categories(
            Reference(ws, min_col=1, min_row=2, max_row=n + 1)
        )
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showPercent = True
        pie.dataLabels.showVal = True
        ws.add_chart(pie, "E1")

        # Stacked bar ساعتی
        if "_hour" in df.columns:
            self._add_hourly_status(ws, df, n)

    def _add_hourly_status(
        self, ws, df: pd.DataFrame, status_count: int
    ):
        cr = status_count + 4
        cross = pd.crosstab(df["_hour"], df["_status"])

        ws.cell(row=cr, column=1, value="ساعت")
        for ci, cn in enumerate(cross.columns, 2):
            ws.cell(row=cr, column=ci, value=cn)
        self.style.apply_header(
            ws, row=cr, max_col=len(cross.columns) + 1
        )

        for ri, (hr, rd) in enumerate(cross.iterrows(), cr + 1):
            ws.cell(row=ri, column=1, value=int(hr))
            for ci, v in enumerate(rd, 2):
                ws.cell(row=ri, column=ci, value=int(v))

        stk = BarChart()
        stk.type = "col"
        stk.grouping = "stacked"
        stk.title = "وضعیت به تفکیک ساعت"
        stk.style = 10
        stk.width, stk.height = 30, 15
        stk.add_data(
            Reference(
                ws,
                min_col=2,
                max_col=len(cross.columns) + 1,
                min_row=cr,
                max_row=cr + len(cross),
            ),
            titles_from_data=True,
        )
        stk.set_categories(
            Reference(
                ws, min_col=1, min_row=cr + 1, max_row=cr + len(cross)
            )
        )
        ws.add_chart(stk, "E28")
