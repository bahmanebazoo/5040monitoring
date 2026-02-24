"""
شیت تحلیل زمانی (ساعتی + روزانه)
"""

from __future__ import annotations

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import AreaChart, BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter

from ..data_preparator import PreparedData
from .base import SheetCreator


def _to_jalali(gregorian_date) -> str:
    """تبدیل تاریخ میلادی به شمسی (جلالی)."""
    try:
        import jdatetime

        if isinstance(gregorian_date, str):
            from datetime import datetime

            gregorian_date = datetime.strptime(
                str(gregorian_date)[:10], "%Y-%m-%d"
            ).date()
        if hasattr(gregorian_date, "date"):
            gregorian_date = gregorian_date.date()
        jd = jdatetime.date.fromgregorian(date=gregorian_date)
        return jd.strftime("%Y/%m/%d")
    except ImportError:
        try:
            from khayyam import JalaliDate

            if isinstance(gregorian_date, str):
                from datetime import datetime

                gregorian_date = datetime.strptime(
                    str(gregorian_date)[:10], "%Y-%m-%d"
                ).date()
            if hasattr(gregorian_date, "date"):
                gregorian_date = gregorian_date.date()
            jd = JalaliDate.from_date(gregorian_date)
            return str(jd)
        except ImportError:
            return str(gregorian_date)


def _configure_axis(chart, y_title: str, x_title: str = "", y_num_fmt: str = "#,##0"):
    """تنظیمات مشترک محورهای Y و X برای همه نمودارها."""

    # ── محور Y ──
    chart.y_axis.title = y_title
    chart.y_axis.delete = False            # محور Y حتماً نمایش داده شود
    chart.y_axis.numFmt = y_num_fmt        # فرمت عددی مقادیر محور
    chart.y_axis.tickLblPos = "low"        # لیبل‌ها در پایین‌ترین جای محور
    chart.y_axis.majorGridlines = None     # خطوط شبکه اصلی (None = پیش‌فرض)

    # ── محور X ──
    if x_title:
        chart.x_axis.title = x_title
    chart.x_axis.delete = False            # محور X حتماً نمایش داده شود
    chart.x_axis.tickLblPos = "low"


class HourlySheetCreator(SheetCreator):

    @property
    def sheet_name(self) -> str:
        return "تحلیل_زمانی"

    def create(self, wb: Workbook, data: PreparedData) -> None:
        ws = wb.create_sheet(self.sheet_name)
        ws.sheet_view.rightToLeft = True

        df = data.df

        if "_hour" not in df.columns:
            ws.cell(row=1, column=1, value="داده زمانی موجود نیست")
            return

        hourly = (
            df.groupby("_hour")
            .agg(تعداد_تماس=("_hour", "size"))
            .reset_index()
            .rename(columns={"_hour": "ساعت"})
        )

        if "_wait_seconds" in df.columns:
            w = (
                df.groupby("_hour")["_wait_seconds"]
                .mean()
                .round(1)
                .reset_index()
            )
            w.columns = ["ساعت", "میانگین_انتظار"]
            hourly = hourly.merge(w, on="ساعت", how="left")

        if "_talk_seconds" in df.columns:
            t = (
                df.groupby("_hour")["_talk_seconds"]
                .mean()
                .round(1)
                .reset_index()
            )
            t.columns = ["ساعت", "میانگین_مکالمه"]
            hourly = hourly.merge(t, on="ساعت", how="left")

        end_row = self._write_table(ws, hourly)
        self.style.auto_fit(ws)
        nrows = len(hourly)
        ncols = len(hourly.columns)

        # ═══════════════════════════════════════════
        # ── نمودار ۱: Line — تعداد تماس بر حسب ساعت ──
        # ═══════════════════════════════════════════
        line = LineChart()
        line.title = "روند تعداد تماس بر حسب ساعت"
        line.style = 10
        line.width, line.height = 28, 15

        # ✅ تنظیم محورها با مقادیر مشخص
        _configure_axis(line, y_title="تعداد", x_title="ساعت", y_num_fmt="#,##0")

        cats = Reference(ws, min_col=1, min_row=2, max_row=nrows + 1)
        d_ref = Reference(ws, min_col=2, min_row=1, max_row=nrows + 1)
        line.add_data(d_ref, titles_from_data=True)
        line.set_categories(cats)

        ws.add_chart(line, f"{get_column_letter(ncols + 2)}1")

        # ═══════════════════════════════════════════
        # ── نمودار ۲: Area — میانگین زمان انتظار ──
        # ═══════════════════════════════════════════
        if "میانگین_انتظار" in hourly.columns:
            wci = list(hourly.columns).index("میانگین_انتظار") + 1
            area = AreaChart()
            area.title = "میانگین زمان انتظار بر حسب ساعت"
            area.style = 10
            area.width, area.height = 28, 15

            # ✅ تنظیم محورها — فرمت اعشاری برای ثانیه
            _configure_axis(area, y_title="ثانیه", x_title="ساعت", y_num_fmt="#,##0.0")

            area.add_data(
                Reference(ws, min_col=wci, min_row=1, max_row=nrows + 1),
                titles_from_data=True,
            )
            area.set_categories(
                Reference(ws, min_col=1, min_row=2, max_row=nrows + 1)
            )

            ws.add_chart(area, f"{get_column_letter(ncols + 2)}31")

        # ═══════════════════════════════════════════
        # ── نمودار ۳: Bar — تماس به تفکیک روز ──
        # ═══════════════════════════════════════════
        if "_date" in df.columns:
            self._add_daily_chart(ws, df, end_row, ncols)

    def _add_daily_chart(
        self, ws, df: pd.DataFrame, start_row: int, ncols: int
    ):
        dr = start_row + 2
        daily = df.groupby("_date").size().reset_index(name="تعداد")
        daily.columns = ["تاریخ", "تعداد"]

        # ── تبدیل تاریخ میلادی به شمسی ──
        daily["تاریخ"] = daily["تاریخ"].apply(_to_jalali)

        # ── هدر جدول دوم ──
        ws.cell(row=dr, column=1, value="تاریخ")
        ws.cell(row=dr, column=2, value="تعداد")
        self.style.apply_header(ws, row=dr, max_col=2)

        # ── داده‌های جدول دوم ──
        for i, (_, rd) in enumerate(daily.iterrows(), dr + 1):
            ws.cell(row=i, column=1, value=rd["تاریخ"])
            ws.cell(row=i, column=2, value=int(rd["تعداد"]))

        # ── اعمال border + font + alignment روی ردیف‌های داده جدول دوم ──
        data_end_row = dr + len(daily)
        self.style.style_data(
            ws, start_row=dr + 1, end_row=data_end_row, max_col=2
        )

        # ── نمودار ستونی ──
        bar = BarChart()
        bar.title = "تماس‌ها به تفکیک روز"
        bar.style = 10
        bar.width, bar.height = 28, 15

        # ✅ تنظیم محورها با مقادیر مشخص
        _configure_axis(bar, y_title="تعداد تماس", x_title="تاریخ", y_num_fmt="#,##0")

        bar.add_data(
            Reference(ws, min_col=2, min_row=dr, max_row=dr + len(daily)),
            titles_from_data=True,
        )
        bar.set_categories(
            Reference(ws, min_col=1, min_row=dr + 1, max_row=dr + len(daily)),
        )

        # ✅ نمایش مقدار روی هر ستون
        if bar.series:
            bar.series[0].dLbls = DataLabelList()
            bar.series[0].dLbls.showVal = True
            bar.series[0].dLbls.numFmt = "#,##0"

        ws.add_chart(bar, f"{get_column_letter(ncols + 2)}60")
