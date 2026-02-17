import pandas as pd
import jdatetime


def parse_jalali_datetime(series: pd.Series) -> pd.Series:
    def convert(value):
        if pd.isna(value):
            return None

        date_part, time_part = str(value).split(" ")
        y, m, d = map(int, date_part.split("/"))
        hh, mm, ss = map(int, time_part.split(":"))

        return jdatetime.datetime(y, m, d, hh, mm, ss).togregorian()

    return series.apply(convert)
