"""
شیت داده خام
"""

from __future__ import annotations

import re
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from ..data_preparator import PreparedData
from .base import SheetCreator


class RawDataSheetCreator(SheetCreator):

    # ── دیکشنری ترجمه ستون‌ها به فارسی ──
    _COLUMN_TRANSLATIONS: dict[str, str] = {
        # زمان و تاریخ
        "datetime": "تاریخ_زمان",
        "mon_datetime": "تاریخ_زمان مونیتورینگ",
        "start_datetime": "تاریخ_شروع",
        "call_datetime": "تاریخ_تماس",
        "event_time_normalized": "زمان_نرمال",
        "monitoring_time": "زمان_مانیتورینگ",
        "date": "تاریخ",
        "time": "زمان",
        "hour": "ساعت",

        # اپراتور / کارشناس
        "agent": "اپراتور",
        "mon_agent": "اپراتور مونیتورینگ",
        "agent_name": "نام_اپراتور",
        "agent_ext": "داخلی_اپراتور",
        "monitoring_agent_ext": "داخلی_مانیتورینگ",
        "operator": "اپراتور",

        # وضعیت
        "status": "وضعیت",
        "call_status": "وضعیت_تماس",
        "mon_status": "وضعیت_مانیتورینگ",
        "monitoring_status": "وضعیت_مانیتورینگ",

        # مچینگ پشتیبانی
        "support_match": "مچ_پشتیبانی",
        "support_match_status": "وضعیت_مچ_پشتیبانی",
        "has_support_match": "مچ_پشتیبانی_دارد",
        "support_matched": "پشتیبانی_مچ_شده",
        "matched_support": "پشتیبانی_تطبیق",
        "match_support": "تطبیق_پشتیبانی",
        "support_status": "وضعیت_پشتیبانی",
        "match_status": "وضعیت_تطبیق",

        # مچینگ پنل
        "rate_match": "مچ_پنل",
        "rate_match_status": "وضعیت_مچ_پنل",
        "has_rate_match": "مچ_پنل_دارد",
        "rate_matched": "پنل_مچ_شده",
        "matched_rate": "پنل_تطبیق",

        # اعتماد / امتیاز
        "confidence": "امتیاز_اطمینان",
        "support_confidence": "اطمینان_پشتیبانی",
        "match_confidence": "اطمینان_تطبیق",

        # دلتا / فاصله زمانی
        "delta_minutes": "فاصله_زمانی_دقیقه",
        "support_delta_minutes": "دلتا_پشتیبانی",
        "rate_delta_minutes": "دلتا_پنل",
        "time_delta": "فاصله_زمانی",

        # انتظار
        "wait_seconds": "زمان_انتظار_ثانیه",
        "wait_time": "زمان_انتظار",
        "queue_time": "زمان_صف",
        "monitoring_wait_seconds": "انتظار_مانیتورینگ",

        # مکالمه
        "talk_seconds": "مدت_مکالمه_ثانیه",
        "talk_time": "مدت_مکالمه",
        "duration": "مدت",
        "call_duration": "مدت_تماس",
        "duration_seconds": "مدت_ثانیه",
        "rate_duration_seconds": "مدت_پنل",

        # مشتری
        "customer": "مشتری",
        "customer_10": "شماره_مشتری",
        "phone": "شماره_تلفن",
        "caller_number": "شماره_تماس‌گیرنده",

        # سایر
        "id": "شناسه",
        "row_id": "شماره_ردیف",
        "call_id": "شناسه_تماس",
        "unique_id": "شناسه_یکتا",
        "queue": "صف",
        "queue_name": "نام_صف",
        "direction": "جهت",
        "call_type": "نوع_تماس",
        "recording": "ضبط",
        "recording_path": "مسیر_ضبط",
        "note": "یادداشت",
        "notes": "یادداشت‌ها",
        "comment": "توضیحات",
        "comments": "توضیحات",
        "rating": "امتیاز",
        "score": "نمره",
        "result": "نتیجه",
        "outcome": "خروجی",
        "category": "دسته‌بندی",
        "type": "نوع",
        "source": "منبع",
        "destination": "مقصد",
        "extension": "داخلی",
        "ext": "داخلی",
        "department": "دپارتمان",
        "team": "تیم",
        "shift": "شیفت",
        "day": "روز",
        "month": "ماه",
        "year": "سال",
        "week": "هفته",
        "weekday": "روز_هفته",
        "day_of_week": "روز_هفته",
        "is_answered": "پاسخ_داده_شده",
        "is_abandoned": "رها_شده",
        "is_missed": "از_دست_رفته",
        "total": "مجموع",
        "count": "تعداد",
        "sum": "جمع",
        "average": "میانگین",
        "avg": "میانگین",
        "min": "حداقل",
        "max": "حداکثر",
        "percent": "درصد",
        "percentage": "درصد",
        "rate": "پنل",
        "ratio": "نسبت",
    }

    @property
    def sheet_name(self) -> str:
        return "داده_خام"

    def create(self, wb: Workbook, data: PreparedData) -> None:
        ws = wb.create_sheet(self.sheet_name)
        ws.sheet_view.rightToLeft = True

        df = data.df

        # ستون‌های internal رو حذف کن
        cols = [c for c in df.columns if not c.startswith("_")]
        export = df[cols].head(10_000)

        # ✅ ترجمه نام ستون‌ها به فارسی
        for ci, cn in enumerate(cols, 1):
            persian_name = self._translate_column(cn)
            ws.cell(row=1, column=ci, value=persian_name)
        self.style.apply_header(ws, max_col=len(cols))

        for ri, row in enumerate(export.itertuples(index=False), 2):
            for ci, v in enumerate(row, 1):
                cell = ws.cell(row=ri, column=ci)
                try:
                    if pd.isna(v):
                        cell.value = ""
                    elif isinstance(v, (pd.Timestamp, datetime)):
                        cell.value = v
                        cell.number_format = "YYYY-MM-DD HH:MM:SS"
                    else:
                        cell.value = v
                except (TypeError, ValueError):
                    cell.value = str(v)

        # ✅ Border برای ردیف‌های داده
        if len(export) > 0:
            self.style.style_data(
                ws,
                start_row=2,
                end_row=len(export) + 1,
                max_col=len(cols),
            )

        self.style.auto_fit(ws)
        ws.auto_filter.ref = (
            f"A1:{get_column_letter(len(cols))}{len(export) + 1}"
        )

    # ── متد ترجمه ستون ──

    def _translate_column(self, col_name: str) -> str:
        """
        نام ستون را به فارسی ترجمه می‌کند.

        ترتیب اولویت:
        1. اگر خودش فارسی است → همان را برگردان
        2. اگر در دیکشنری ترجمه هست → ترجمه را برگردان
        3. اگر با حروف کوچک در دیکشنری هست → ترجمه را برگردان
        4. در غیر این صورت → نام اصلی با _ به جای فاصله
        """
        # اگر خودش فارسی است (حداقل یک کاراکتر فارسی دارد)
        if any('\u0600' <= c <= '\u06FF' for c in col_name):
            return col_name

        # جستجوی مستقیم
        if col_name in self._COLUMN_TRANSLATIONS:
            return self._COLUMN_TRANSLATIONS[col_name]

        # جستجو با حروف کوچک
        col_lower = col_name.lower().strip()
        if col_lower in self._COLUMN_TRANSLATIONS:
            return self._COLUMN_TRANSLATIONS[col_lower]

        # جستجوی جزئی — اگر بخشی از نام در دیکشنری باشد
        col_normalized = col_lower.replace(" ", "_").replace("-", "_")
        for eng, fa in self._COLUMN_TRANSLATIONS.items():
            if eng in col_normalized or col_normalized in eng:
                return fa

        # اگر پیدا نشد، نام اصلی را با فرمت بهتر برگردان
        result = col_name.replace(" ", "_").replace("-", "_")

        # تبدیل CamelCase به snake_case
        result = re.sub(r'([a-z])([A-Z])', r'\1_\2', result).lower()

        return result
