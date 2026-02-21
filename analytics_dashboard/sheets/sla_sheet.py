"""
شیت SLA
"""

from __future__ import annotations

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference

from ..data_preparator import PreparedData
from .base import SheetCreator


class SLASheetCreator(SheetCreator):

    @property
    def sheet_name(self) -> str:
        return "SLA"

    def create(self, wb: Workbook, data: PreparedData) -> None:
        ws = wb.create_sheet(self.sheet_name)
        df = data.df

        if (
            "_wait_seconds" not in df.columns
            or "_status" not in df.columns
        ):
            ws.cell(row=1, column=1, value="داده SLA موجود نیست")
            return

        conn_df = df[self._connected_mask(df)].copy()
        total_conn = len(conn_df)

        thresholds = [10, 15, 20, 30, 45, 60, 90, 120]

        # ── هدر جدول اصلی ──
        ws.cell(row=1, column=1, value="آستانه (ثانیه)")
        ws.cell(row=1, column=2, value="تعداد_پاسخ")
        ws.cell(row=1, column=3, value="SLA_%")
        ws.cell(row=1, column=4, value="≥2min")       # ⭐ ستون جدید
        self.style.apply_header(ws, max_col=4)

        for ri, th in enumerate(thresholds, 2):
            cnt = int((conn_df["_wait_seconds"] <= th).sum())
            pct = cnt / max(total_conn, 1) * 100

            # ⭐ تعداد تماس‌هایی که انتظارشون بالای ۱۲۰ ثانیه بوده
            #    (بین تماس‌هایی که تا این آستانه پاسخ نگرفتن)
            over_2min = int((conn_df["_wait_seconds"] >= 120).sum())

            ws.cell(row=ri, column=1, value=th)
            ws.cell(row=ri, column=2, value=cnt)
            ws.cell(row=ri, column=3, value=f"{pct:.1f}%")
            ws.cell(row=ri, column=4, value=over_2min)

        self.style.style_data(ws)
        n = len(thresholds)

        # ── خلاصه آمار ≥2min ──
        summary_row = n + 3
        over_2min_total = int(
            (conn_df["_wait_seconds"] >= 120).sum()
        )
        over_2min_pct = (
            over_2min_total / max(total_conn, 1) * 100
        )
        ws.cell(
            row=summary_row, column=1,
            value="📊 تماس‌های ≥2min"
        )
        ws.cell(row=summary_row, column=2, value=over_2min_total)
        ws.cell(
            row=summary_row, column=3,
            value=f"{over_2min_pct:.1f}%"
        )
        ws.cell(
            row=summary_row, column=4,
            value=f"از {total_conn} وصل‌شده"
        )
        self.style.apply_header(ws, row=summary_row, max_col=4)

        # ── Line Chart ──
        line = LineChart()
        line.title = "منحنی SLA"
        line.style = 10
        line.y_axis.title = "تعداد"
        line.x_axis.title = "ثانیه"
        line.width, line.height = 25, 14
        line.add_data(
            Reference(ws, min_col=2, min_row=1, max_row=n + 1),
            titles_from_data=True,
        )
        line.set_categories(
            Reference(ws, min_col=1, min_row=2, max_row=n + 1)
        )
        if line.series:
            line.series[0].graphicalProperties.line.width = 25000
        ws.add_chart(line, "F1")

        # ── SLA ساعتی ──
        if "_hour" in conn_df.columns:
            self._add_hourly_sla(ws, conn_df, summary_row + 2)

        self.style.auto_fit(ws)

    # ── جدول ساعتی ──

    def _add_hourly_sla(self, ws, conn_df, sr: int):
        ws.cell(row=sr, column=1, value="ساعت")
        ws.cell(row=sr, column=2, value="کل_وصل")
        ws.cell(row=sr, column=3, value="≤20s")
        ws.cell(row=sr, column=4, value="SLA_%")
        ws.cell(row=sr, column=5, value="≥2min")        # ⭐ ستون جدید
        ws.cell(row=sr, column=6, value="≥2min_%")       # ⭐ درصدش هم
        self.style.apply_header(ws, row=sr, max_col=6)

        for h in range(24):
            hdf = conn_df[conn_df["_hour"] == h]
            hc = len(hdf)
            s20 = (
                int((hdf["_wait_seconds"] <= 20).sum()) if hc else 0
            )
            # ⭐ تعداد تماس‌های ≥120 ثانیه در این ساعت
            over2 = (
                int((hdf["_wait_seconds"] >= 120).sum()) if hc else 0
            )
            over2_pct = over2 / max(hc, 1) * 100

            r = sr + 1 + h
            ws.cell(row=r, column=1, value=h)
            ws.cell(row=r, column=2, value=hc)
            ws.cell(row=r, column=3, value=s20)
            ws.cell(
                row=r, column=4,
                value=f"{s20 / max(hc, 1) * 100:.1f}%",
            )
            ws.cell(row=r, column=5, value=over2)        # ⭐
            ws.cell(row=r, column=6, value=f"{over2_pct:.1f}%")  # ⭐

        # ── Bar Chart ساعتی (شامل ≥2min) ──
        bar = BarChart()
        bar.type = "col"
        bar.title = "SLA ساعتی (≤20s و ≥2min)"
        bar.style = 10
        bar.width, bar.height = 30, 15
        bar.add_data(
            Reference(
                ws,
                min_col=2,
                max_col=5,        # ⭐ شامل ستون ≥2min
                min_row=sr,
                max_row=sr + 24,
            ),
            titles_from_data=True,
        )
        bar.set_categories(
            Reference(
                ws, min_col=1, min_row=sr + 1, max_row=sr + 24
            )
        )
        ws.add_chart(bar, "H18")
