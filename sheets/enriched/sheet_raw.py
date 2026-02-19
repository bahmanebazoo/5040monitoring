"""
sheets/enriched/sheet_raw.py
────────────────────────────
شیت‌های: داده_خام_غنی‌شده، وصل_شده، رها_شده
"""

import pandas as pd
from .utils import connected_mask, abandoned_mask


def write_raw_sheets(writer: pd.ExcelWriter, df: pd.DataFrame) -> None:
    """سه شیت اصلی داده خام."""

    # شیت ۱: تمام داده‌ها
    df.to_excel(writer, index=False, sheet_name="داده_خام_غنی‌شده")

    # شیت ۲: وصل‌شده‌ها
    df[connected_mask(df)].to_excel(
        writer, index=False, sheet_name="وصل_شده"
    )

    # شیت ۳: رها‌شده‌ها
    df[abandoned_mask(df)].to_excel(
        writer, index=False, sheet_name="رها_شده"
    )
