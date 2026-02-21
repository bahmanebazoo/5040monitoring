"""
شیت تحلیل مچینگ
"""

from __future__ import annotations

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import DoughnutChart, Reference

from ..data_preparator import PreparedData
from .base import SheetCreator


class MatchingSheetCreator(SheetCreator):

    @property
    def sheet_name(self) -> str:
        return "تحلیل_مچینگ"

    def create(self, wb: Workbook, data: PreparedData) -> None:
        ws = wb.create_sheet(self.sheet_name)
        df = data.df
        rc = 1

        # ── جدول وضعیت هر نوع مچ ──
        ws.cell(row=rc, column=1, value="نوع")
        ws.cell(row=rc, column=2, value="وضعیت")
        ws.cell(row=rc, column=3, value="تعداد")
        ws.cell(row=rc, column=4, value="درصد")
        self.style.apply_header(ws, row=rc, max_col=4)
        rc += 1

        for col, label in [
            (data.support_match_col, "پشتیبانی"),
            (data.rate_match_col, "نرخ‌دهی"),
        ]:
            if col and col in df.columns:
                vc = (
                    df[col]
                    .fillna("نامشخص")
                    .astype(str)
                    .value_counts()
                )
                for st, cnt in vc.items():
                    ws.cell(row=rc, column=1, value=label)
                    ws.cell(row=rc, column=2, value=st)
                    ws.cell(row=rc, column=3, value=cnt)
                    ws.cell(
                        row=rc,
                        column=4,
                        value=f"{cnt / data.total * 100:.1f}%",
                    )
                    rc += 1

        self.style.style_data(ws, start_row=2)

        # ── ترکیبی ──
        s_mask = self._match_mask(df, data.support_match_col)
        r_mask = self._match_mask(df, data.rate_match_col)

        rc += 1
        ws.cell(row=rc, column=1, value="وضعیت ترکیبی")
        ws.cell(row=rc, column=2, value="تعداد")
        self.style.apply_header(ws, row=rc, max_col=2)
        combo_header = rc
        rc += 1

        if s_mask.any() or r_mask.any():
            combos = [
                ("هر دو مچ", int((s_mask & r_mask).sum())),
                ("فقط پشتیبانی", int((s_mask & ~r_mask).sum())),
                ("فقط نرخ‌دهی", int((~s_mask & r_mask).sum())),
                ("هیچکدام", int((~s_mask & ~r_mask).sum())),
            ]
        else:
            combos = [("داده کافی نیست", 0)]

        combo_start = rc
        for lb, vl in combos:
            ws.cell(row=rc, column=1, value=lb)
            ws.cell(row=rc, column=2, value=vl)
            rc += 1

        if len(combos) > 1:
            dnt = DoughnutChart()
            dnt.title = "توزیع ترکیبی مچینگ"
            dnt.style = 10
            dnt.width, dnt.height = 20, 14
            dnt.add_data(
                Reference(
                    ws,
                    min_col=2,
                    min_row=combo_header,
                    max_row=combo_start + len(combos) - 1,
                ),
                titles_from_data=True,
            )
            dnt.set_categories(
                Reference(
                    ws,
                    min_col=1,
                    min_row=combo_start,
                    max_row=combo_start + len(combos) - 1,
                )
            )
            ws.add_chart(dnt, "F1")

        self.style.auto_fit(ws)
