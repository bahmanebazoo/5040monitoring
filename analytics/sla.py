import pandas as pd


def calculate_asa(monitoring_df: pd.DataFrame) -> float:
    connected = monitoring_df[monitoring_df["status"] == "وصل شده"]

    return connected["wait_seconds"].mean()
