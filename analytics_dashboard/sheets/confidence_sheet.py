"""
شیت تحلیل امتیاز اعتماد (Confidence)
"""

from __future__ import annotations

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

from ..data_preparator import PreparedData
from .base import SheetCreator


class ConfidenceSheetCreator(SheetCreator):

    @property
    def sheet_name(self) -> str:
        return "تحلیل_اعتماد"

    # ── جدول راهنمای امتیازدهی ──────────────────────────────────
    _GUIDE_HEADERS = [
        "وضعیت",
        "مشتری",
        "اپراتور",
        "زمان",
        "امتیاز",
        "Sup ≥ 75",
        "Rate ≥ 80",
    ]

    _GUIDE_ROWS = [
        ("مشتری + اپراتور + < 3 دقیقه",   40, 35, 25, 100, "✅", "✅"),
        ("مشتری + اپراتور + < 5 دقیقه",   40, 35, 20,  95, "✅", "✅"),
        ("مشتری + اپراتور + < 10 دقیقه",  40, 35, 15,  90, "✅", "✅"),
        ("مشتری + اپراتور + < 15 دقیقه",  40, 35, 10,  85, "✅", "✅"),
        ("مشتری + اپراتور + < 30 دقیقه",  40, 35,  5,  80, "✅", "✅"),
        ("مشتری + اپراتور + < 45 دقیقه",  40, 35,  2,  77, "✅", "❌"),
        ("مشتری + اپراتور + ≥ 45 دقیقه",  40, 35,  0,  75, "✅", "❌"),
        ("مشتری + بدون اپراتور + < 3 دقیقه",  40,  0, 25, 65, "❌", "❌"),
        ("مشتری + بدون اپراتور + < 5 دقیقه",  40,  0, 20, 60, "❌", "❌"),
        ("مشتری + بدون اپراتور + < 10 دقیقه", 40,  0, 15, 55, "❌", "❌"),
        ("مشتری + بدون اپراتور + < 15 دقیقه", 40,  0, 10, 50, "❌", "❌"),
        ("مشتری + بدون اپراتور + < 30 دقیقه", 40,  0,  5, 45, "❌", "❌"),
        ("مشتری + بدون اپراتور + < 45 دقیقه", 40,  0,  2, 42, "❌", "❌"),
        ("مشتری + بدون اپراتور + ≥ 45 دقیقه", 40,  0,  0, 40, "❌", "❌"),
        ("بدون مشتری + اپراتور + < 3 دقیقه",   0, 35, 25, 60, "❌", "❌"),
        ("بدون مشتری + اپراتور + < 5 دقیقه",   0, 35, 20, 55, "❌", "❌"),
        ("بدون مشتری + اپراتور + < 10 دقیقه",  0, 35, 15, 50, "❌", "❌"),
        ("بدون مشتری + اپراتور + < 15 دقیقه",  0, 35, 10, 45, "❌", "❌"),
        ("بدون مشتری + اپراتور + < 30 دقیقه",  0, 35,  5, 40, "❌", "❌"),
        ("بدون مشتری + اپراتور + < 45 دقیقه",  0, 35,  2, 37, "❌", "❌"),
        ("بدون مشتری + اپراتور + ≥ 45 دقیقه",  0, 35,  0, 35, "❌", "❌"),
        ("بدون مشتری + بدون اپراتور",           0,  0,  0,  0, "❌", "❌"),
    ]

    # ================================================================
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

        # ── 1) جدول توزیع امتیاز ─────────────────────────────────
        vc = cs.value_counts().sort_index().reset_index()
        vc.columns = ["امتیاز", "تعداد"]
        vc["درصد"] = (vc["تعداد"] / vc["تعداد"].sum() * 100).round(1)

        self._write_table(ws, vc)
        n = len(vc)

        # ✅ border جدول توزیع
        self.style.style_data(
            ws,
            start_row=2,
            end_row=n + 1,
            max_col=3,
        )

        # ── 2) نمودار Bar ────────────────────────────────────────
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

        # ✅ محور Y مشخص
        bar.y_axis.delete = False
        bar.y_axis.numFmt = "#,##0"
        bar.y_axis.tickLblPos = "low"
        bar.y_axis.title = "تعداد"

        # ✅ نمایش مقدار روی ستون‌ها
        for s in bar.series:
            s.dLbls = DataLabelList()
            s.dLbls.showVal = True
            s.dLbls.numFmt = "#,##0"

        ws.add_chart(bar, "E1")

        # ── 3) نمودار Pie ────────────────────────────────────────
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
        pie.dataLabels.showCatName = True
        ws.add_chart(pie, "E28")

        # ── 4) آمار توصیفی ───────────────────────────────────────
        if pd.api.types.is_numeric_dtype(cs):
            sr = n + 4
            desc = {
                "میانگین": cs.mean(),
                "میانه": cs.median(),
                "انحراف معیار": cs.std(),
                "حداقل": cs.min(),
                "حداکثر": cs.max(),
                "تعداد کل": cs.count(),
            }

            # هدر آمار توصیفی
            ws.cell(row=sr, column=1, value="آمار")
            ws.cell(row=sr, column=2, value="مقدار")
            self.style.apply_header(ws, row=sr, max_col=2)

            for i, (k, v) in enumerate(desc.items(), sr + 1):
                ws.cell(row=i, column=1, value=k)
                ws.cell(row=i, column=2, value=round(v, 2))

            stat_end = sr + len(desc)

            # ✅ border جدول آمار توصیفی
            self.style.style_data(
                ws,
                start_row=sr + 1,
                end_row=stat_end,
                max_col=2,
            )

        # ── 5) جدول راهنمای نمره‌دهی از سلول U5 ─────────────────
        self._write_scoring_guide(ws)

        self.style.auto_fit(ws)

    # ================================================================
    #  جدول راهنمای امتیازدهی — شروع از U5
    # ================================================================
    def _write_scoring_guide(self, ws) -> None:
        """
        جدول راهنمای محاسبه امتیاز اعتماد (Confidence) را از سلول U5
        با استایل کامل و border رسم می‌کند.
        """

        start_col = 40          # ستون U
        start_row = 5
        num_cols = len(self._GUIDE_HEADERS)     # 7 ستون

        thin = Side(style="thin")
        border_all = Border(top=thin, bottom=thin, left=thin, right=thin)

        # ── عنوان اصلی (مرج شده) ────────────────────────────────
        title_row = start_row
        title_cell = ws.cell(
            row=title_row,
            column=start_col,
            value="📋 راهنمای محاسبه امتیاز اعتماد (Confidence Score)",
        )
        ws.merge_cells(
            start_row=title_row,
            start_column=start_col,
            end_row=title_row,
            end_column=start_col + num_cols - 1,
        )
        title_cell.font = Font(name="B Nazanin", size=14, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill("solid", fgColor="4472C4")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        title_cell.border = border_all
        # border تمام سلول‌های مرج‌شده عنوان
        for ci in range(start_col, start_col + num_cols):
            ws.cell(row=title_row, column=ci).border = border_all

        # ── توضیح زیرعنوان ───────────────────────────────────────
        desc_row = title_row + 1
        desc_cell = ws.cell(
            row=desc_row,
            column=start_col,
            value=(
                "امتیاز اعتماد از 0 تا 100 محاسبه می‌شود:  "
                "تطبیق مشتری = 40 امتیاز  |  "
                "تطبیق اپراتور = 35 امتیاز  |  "
                "نزدیکی زمانی = حداکثر 25 امتیاز"
            ),
        )
        ws.merge_cells(
            start_row=desc_row,
            start_column=start_col,
            end_row=desc_row,
            end_column=start_col + num_cols - 1,
        )
        desc_cell.font = Font(name="B Nazanin", size=11, italic=True, color="333333")
        desc_cell.fill = PatternFill("solid", fgColor="D9E2F3")
        desc_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        desc_cell.border = border_all
        for ci in range(start_col, start_col + num_cols):
            ws.cell(row=desc_row, column=ci).border = border_all

        # ── هدر جدول ─────────────────────────────────────────────
        header_row = desc_row + 1       # = start_row + 2
        header_font = Font(name="B Nazanin", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill("solid", fgColor="2F5496")
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for ci, h in enumerate(self._GUIDE_HEADERS, start_col):
            cell = ws.cell(row=header_row, column=ci, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = border_all

        # ── ردیف‌های داده ─────────────────────────────────────────
        data_font = Font(name="B Nazanin", size=11)
        data_align = Alignment(horizontal="center", vertical="center")

        # رنگ‌بندی متناوب
        fill_white = PatternFill("solid", fgColor="FFFFFF")
        fill_light = PatternFill("solid", fgColor="F2F2F2")

        # رنگ ویژه برای ✅ و ❌
        green_font = Font(name="B Nazanin", size=12, bold=True, color="217346")
        red_font = Font(name="B Nazanin", size=12, bold=True, color="C00000")

        # رنگ ویژه برای سطح امتیاز
        def _score_fill(score: int) -> PatternFill:
            """رنگ پس‌زمینه بر اساس سطح امتیاز."""
            if score >= 90:
                return PatternFill("solid", fgColor="C6EFCE")   # سبز روشن
            elif score >= 75:
                return PatternFill("solid", fgColor="FFEB9C")   # زرد روشن
            elif score >= 50:
                return PatternFill("solid", fgColor="FFF2CC")   # نارنجی خیلی روشن
            else:
                return PatternFill("solid", fgColor="FFC7CE")   # قرمز روشن

        for ri, row_data in enumerate(self._GUIDE_ROWS):
            excel_row = header_row + 1 + ri
            row_fill = fill_light if ri % 2 == 0 else fill_white

            for ci, val in enumerate(row_data):
                col_idx = start_col + ci
                cell = ws.cell(row=excel_row, column=col_idx, value=val)
                cell.alignment = data_align
                cell.border = border_all

                # ستون «وضعیت» (اولین ستون) — چپ‌چین
                if ci == 0:
                    cell.font = Font(name="B Nazanin", size=11, bold=False)
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.fill = row_fill
                # ستون «امتیاز» (ci == 4) — رنگ‌بندی سطحی
                elif ci == 4:
                    cell.font = Font(name="B Nazanin", size=11, bold=True)
                    cell.fill = _score_fill(val)
                # ستون‌های ✅ / ❌ (ci == 5, 6)
                elif ci in (5, 6):
                    cell.font = green_font if val == "✅" else red_font
                    cell.fill = row_fill
                else:
                    cell.font = data_font
                    cell.fill = row_fill

        # ── پاورقی ───────────────────────────────────────────────
        footer_row = header_row + 1 + len(self._GUIDE_ROWS)
        notes = [
            "💡 Sup ≥ 75 : امتیاز کافی برای تأیید تطبیق پشتیبانی",
            "💡 Rate ≥ 80 : امتیاز کافی برای تأیید تطبیق نرخ‌دهی (پنل)",
            "💡 فاصله زمانی (دلتا) = فاصله بین زمان تماس و ثبت در سیستم مقصد",
        ]
        for ni, note_text in enumerate(notes):
            nr = footer_row + ni
            note_cell = ws.cell(row=nr, column=start_col, value=note_text)
            ws.merge_cells(
                start_row=nr,
                start_column=start_col,
                end_row=nr,
                end_column=start_col + num_cols - 1,
            )
            note_cell.font = Font(name="B Nazanin", size=10, italic=True, color="4472C4")
            note_cell.alignment = Alignment(horizontal="right", vertical="center")

        # ── تنظیم عرض ستون‌های جدول راهنما ────────────────────────
        col_widths = [38, 10, 10, 8, 10, 12, 12]
        for ci, w in enumerate(col_widths):
            from openpyxl.utils import get_column_letter
            col_letter = get_column_letter(start_col + ci)
            ws.column_dimensions[col_letter].width = w
