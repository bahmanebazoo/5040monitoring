"""
شیت خلاصه KPI — نمایش سطری + نمودارهای زیر هم
هر آبجکت (جدول یا نمودار) از cursor شروع شده و cursor را به‌روز می‌کند.
"""

from __future__ import annotations

import numpy as np
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint

from ..data_preparator import PreparedData
from .base import SheetCreator


class KPISheetCreator(SheetCreator):

    GAP = 2               # فاصله استاندارد بین آبجکت‌ها
    CHART_HEIGHT = 15      # ارتفاع تقریبی نمودار (به ردیف)

    @property
    def sheet_name(self) -> str:
        return "خلاصه_کلی"

    # ══════════════════════════════════════
    #  create
    # ══════════════════════════════════════

    def create(self, wb: Workbook, data: PreparedData) -> None:
        ws = wb.create_sheet(self.sheet_name)
        ws.sheet_view.rightToLeft = True

        df = data.df
        total = data.total
        conn = data.connected_count
        s_count = data.support_match_count
        r_count = data.rate_match_count

        avg_wait = (
            df["_wait_seconds"].mean()
            if "_wait_seconds" in df.columns
            else float("nan")
        )
        avg_talk = (
            df["_talk_seconds"].mean()
            if "_talk_seconds" in df.columns
            else float("nan")
        )

        sla_count = 0
        if "_wait_seconds" in df.columns:
            sla_count = int(
                ((df["_wait_seconds"] <= 20) & self._connected_mask(df)).sum()
            )
        sla_pct = sla_count / max(total, 1) * 100
        agents = df["_agent"].nunique() if "_agent" in df.columns else 0

        # ─── cursor = ردیف شروع آبجکت بعدی ───
        cursor = 1

        # ══════════════════════════════════════
        #  فاز ۱: همه جداول
        # ══════════════════════════════════════

        # ① جدول KPI
        cursor = self._write_kpi_table(
            ws, cursor,
            total=total, conn=conn, s_count=s_count, r_count=r_count,
            avg_wait=avg_wait, avg_talk=avg_talk,
            sla_pct=sla_pct, sla_count=sla_count, agents=agents, data=data,
        )
        cursor += self.GAP

        # ② جدول داده Pie
        pie_refs, cursor = self._write_pie_data(ws, cursor, conn, total)
        cursor += self.GAP

        # ③ جدول داده Bar
        bar_refs, cursor = self._write_bar_data(ws, cursor, s_count, r_count, total)
        cursor += self.GAP

        # ══════════════════════════════════════
        #  فاز ۲: همه نمودارها زیر هم
        # ══════════════════════════════════════

        # ④ نمودار Pie
        ws.add_chart(self._build_pie(ws, pie_refs), f"A{cursor}")
        cursor += self.CHART_HEIGHT + self.GAP

        # ⑤ نمودار Bar
        ws.add_chart(self._build_bar(ws, bar_refs), f"A{37}")

        # عرض ستون‌ها
        self.style.auto_fit(ws)

    # ══════════════════════════════════════
    #  ① جدول KPI سطری
    # ══════════════════════════════════════

    def _write_kpi_table(
        self, ws, cursor: int, *,
        total, conn, s_count, r_count,
        avg_wait, avg_talk, sla_pct, sla_count, agents, data,
    ) -> int:
        headers = [
            "کل تماس‌ها", "وصل‌شده", "بی‌پاسخ",
            "مچ پشتیبانی ✅", "مچ پنل ✅",
            "میانگین انتظار (ثانیه)", "میانگین مکالمه (ثانیه)",
            "SLA ≤20s", "تعداد اپراتور",
        ]
        values = [
            total, conn, total - conn, s_count, r_count,
            f"{avg_wait:.1f}" if not np.isnan(avg_wait) else "N/A",
            f"{avg_talk:.1f}" if not np.isnan(avg_talk) else "N/A",
            f"{sla_pct:.1f}%", 10, # تعدا اپراتورها
        ]
        descriptions = [
            "همه تماس‌های مانیتورینگ",
            f"{conn / max(total, 1) * 100:.1f}%",
            f"{(total - conn) / max(total, 1) * 100:.1f}%",
            f"{s_count / max(total, 1) * 100:.1f}% :تماسهایی که داده پشتیبانی دارند ",
            f"{r_count / max(total, 1) * 100:.1f}% :تماسهایی که داده پنل دارند ",
            "ASA", "AHT",
            f"{sla_count}/{total}", "",
        ]

        ncols = len(headers)

        # عنوان بخش (merge & center)
        next_r = self.style.write_section_title(
            ws, row=cursor,
            title="📊 شاخص‌های کلیدی عملکرد",
            col_count=ncols,
        )

        # هدر
        for ci, h in enumerate(headers, 1):
            ws.cell(row=next_r, column=ci, value=h)
        self.style.apply_header(ws, row=next_r, max_col=ncols)

        # مقدار
        val_row = next_r + 1
        for ci, v in enumerate(values, 1):
            cell = ws.cell(row=val_row, column=ci, value=v)
            cell.fill = self.style.KPI_FILL

        # توضیحات
        desc_row = val_row + 1
        for ci, d in enumerate(descriptions, 1):
            ws.cell(row=desc_row, column=ci, value=d)

        self.style.style_data(
            ws, start_row=val_row, end_row=desc_row, max_col=ncols,
        )

        return desc_row  # آخرین ردیف اشغال‌شده

    # ══════════════════════════════════════
    #  ② جدول داده Pie
    # ══════════════════════════════════════

    def _write_pie_data(self, ws, cursor: int, conn: int, total: int):
        ncols = 2

        next_r = self.style.write_section_title(
            ws, row=cursor,
            title="📈 توزیع وضعیت تماس‌ها",
            col_count=ncols,
        )

        ws.cell(row=next_r, column=1, value="وضعیت")
        ws.cell(row=next_r, column=2, value="تعداد")
        self.style.apply_header(ws, row=next_r, max_col=ncols)

        ws.cell(row=next_r + 1, column=1, value="وصل شده")
        ws.cell(row=next_r + 1, column=2, value=conn)
        ws.cell(row=next_r + 2, column=1, value="بی‌پاسخ")
        ws.cell(row=next_r + 2, column=2, value=total - conn)

        self.style.style_data(
            ws, start_row=next_r + 1, end_row=next_r + 2, max_col=ncols,
        )

        refs = {
            "header_row": next_r,
            "data_start": next_r + 1,
            "data_end": next_r + 2,
        }
        return refs, next_r + 2  # آخرین ردیف اشغال‌شده

    # ══════════════════════════════════════
    #  ③ جدول داده Bar
    # ══════════════════════════════════════

    def _write_bar_data(
        self, ws, cursor: int,
        s_count: int, r_count: int, total: int,
    ):
        ncols = 3

        next_r = self.style.write_section_title(
            ws, row=cursor,
            title="📊 وضعیت مچینگ",
            col_count=ncols,
        )

        ws.cell(row=next_r, column=1, value="نوع")
        ws.cell(row=next_r, column=2, value="مچ‌شده")
        ws.cell(row=next_r, column=3, value="مچ‌نشده")
        self.style.apply_header(ws, row=next_r, max_col=ncols)

        ws.cell(row=next_r + 1, column=1, value="پشتیبانی")
        ws.cell(row=next_r + 1, column=2, value=s_count)
        ws.cell(row=next_r + 1, column=3, value=total - s_count)
        ws.cell(row=next_r + 2, column=1, value="پنل")
        ws.cell(row=next_r + 2, column=2, value=r_count)
        ws.cell(row=next_r + 2, column=3, value=total - r_count)

        self.style.style_data(
            ws, start_row=next_r + 1, end_row=next_r + 2, max_col=ncols,
        )

        refs = {
            "header_row": next_r,
            "data_start": next_r + 1,
            "data_end": next_r + 2,
        }
        return refs, next_r + 2

    # ══════════════════════════════════════
    #  ④ ساخت نمودار Pie
    # ══════════════════════════════════════

    def _build_pie(self, ws, refs: dict) -> PieChart:
        pie = PieChart()
        pie.title = "توزیع وضعیت تماس‌ها"
        pie.style = 10
        pie.width, pie.height = 16, 10

        labels = Reference(
            ws, min_col=1,
            min_row=refs["data_start"],
            max_row=refs["data_end"],
        )
        d_ref = Reference(
            ws, min_col=2,
            min_row=refs["header_row"],
            max_row=refs["data_end"],
        )
        pie.add_data(d_ref, titles_from_data=True)
        pie.set_categories(labels)

        if pie.series:
            pt0 = DataPoint(idx=0)
            pt0.graphicalProperties.solidFill = "2CA02C"
            pt1 = DataPoint(idx=1)
            pt1.graphicalProperties.solidFill = "D62728"
            pie.series[0].data_points = [pt0, pt1]

        pie.dataLabels = DataLabelList()
        pie.dataLabels.showPercent = True
        pie.dataLabels.showVal = True

        return pie

    # ══════════════════════════════════════
    #  ⑤ ساخت نمودار Bar
    # ══════════════════════════════════════

    def _build_bar(self, ws, refs: dict) -> BarChart:
        bar = BarChart()
        bar.type = "col"
        bar.grouping = "stacked"
        bar.title = "وضعیت مچینگ"
        bar.style = 10
        bar.width, bar.height = 16, 10

        cats = Reference(
            ws, min_col=1,
            min_row=refs["data_start"],
            max_row=refs["data_end"],
        )
        d_ref = Reference(
            ws, min_col=2, max_col=3,
            min_row=refs["header_row"],
            max_row=refs["data_end"],
        )
        bar.add_data(d_ref, titles_from_data=True)
        bar.set_categories(cats)

        bar.dataLabels = DataLabelList()
        bar.dataLabels.showVal = True

        return bar
