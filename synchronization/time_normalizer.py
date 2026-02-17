import pandas as pd


class TimeNormalizer:
    def __init__(self, offset_minutes: int = 0):
        self.offset_minutes = offset_minutes

    def apply(self, df: pd.DataFrame, column: str) -> pd.DataFrame:
        df[column + "_normalized"] = (
            df[column] + pd.Timedelta(minutes=self.offset_minutes)
        )
        return df
