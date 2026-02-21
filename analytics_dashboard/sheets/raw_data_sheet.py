"""
شیت داده خام
"""

from __future__ import annotations

from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import DataBarRule
from openpyxl.utils import get_column_letter

from ..data_preparator import PreparedData
from .base import SheetCreator


class RawDataSheetCreator(SheetCreator):

    @property
    def sheet_name(self) -> str:
        return "داده_خام"

    def create(self, wb: Workbook, data: PreparedData) -> None:
        ws = wb.create_sheet(self.sheet_name)
        df = data.df

        # ستون‌های internal رو حذف کن
        cols = [c for c in df.columns if not c.startswith("_")]
        export = df[cols].head(10_000)

        for ci, cn in enumerate(cols, 1):
            ws.cell(row=1, column=ci, value=cn)
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

        self.style.auto_fit(ws)
        ws.auto_filter.ref = (
            f"A1:{get_column_letter(len(cols))}{len(export) + 1}"
        )

        # Data bars برای ستون‌های عددی
        for ci, cn in enumerate(cols, 1):
            if cn in df.columns and pd.api.types.is_numeric_dtype(df[cn]):
                cl = get_column_letter(ci)
                ws.conditional_formatting.add(
                    f"{cl}2:{cl}{len(export) + 1}",
                    DataBarRule(
                        start_type="min",
                        end_type="max",
                        color="638EC6",
                    ),
                )
