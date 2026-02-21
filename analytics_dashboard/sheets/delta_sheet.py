"""
شیت تحلیل دلتا (فاصله زمانی مچینگ)
"""

from __future__ import annotations

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

from ..data_preparator import PreparedData
from .base import SheetCreator


class DeltaSheetCreator(SheetCreator):

    @property
    def sheet_name(self) -> str:
        return "تحلیل_دلتا"

    def create(self, wb: Workbook, data: PreparedData) -> None:
        ws = wb.create_sheet(self.sheet_name)
        df = data.df

        if "_delta_minutes" not in df.columns:
            ws.cell(row=1, column=1, value="داده Delta موجود نیست")
            return

        d = df["_delta_minutes"].dropna()
        if d.empty:
            ws.cell(row=1, column=1, value="Delta خالی")
            return

        # آمار
        ws.cell(row=1, column=1, value="آمار دلتا (دقیقه)")
        ws.cell(row=1, column=2, value="مقدار")
        self.style.apply_header(ws, max_col=2)
        rc = 2
        for k, v in {
            "میانگین": d.mean(),
            "میانه": d.median(),
            "std": d.std(),
            "min": d.min(),
            "max": d.max(),
            "p10": d.quantile(0.1),
            "p25": d.quantile(0.25),
            "p75": d.quantile(0.75),
            "p90": d.quantile(0.9),
            "p95": d.quantile(0.95),
        }.items():
            ws.cell(row=rc, column=1, value=k)
            ws.cell(row=rc, column=2, value=round(v, 2))
            rc += 1

        # توزیع
        rc += 1
        bins = [0, 1, 2, 5, 10, 15, 30, 60, float("inf")]
        lbls = [
            "0-1m", "1-2m", "2-5m", "5-10m",
            "10-15m", "15-30m", "30-60m", "60m+",
        ]
        dist = pd.cut(d.abs(), bins=bins, labels=lbls, right=False)
        dist = dist.value_counts().sort_index().reset_index()
        dist.columns = ["بازه", "تعداد"]

        ws.cell(row=rc, column=1, value="بازه")
        ws.cell(row=rc, column=2, value="تعداد")
        self.style.apply_header(ws, row=rc, max_col=2)
        ds = rc
        rc += 1
        for _, rd in dist.iterrows():
            ws.cell(row=rc, column=1, value=rd["بازه"])
            ws.cell(row=rc, column=2, value=int(rd["تعداد"]))
            rc += 1

        bar = BarChart()
        bar.title = "توزیع فاصله زمانی مچینگ"
        bar.style = 10
        bar.width, bar.height = 25, 14
        bar.add_data(
            Reference(ws, min_col=2, min_row=ds, max_row=ds + len(dist)),
            titles_from_data=True,
        )
        bar.set_categories(
            Reference(
                ws, min_col=1, min_row=ds + 1, max_row=ds + len(dist)
            )
        )
        ws.add_chart(bar, "D1")

        self.style.auto_fit(ws)
