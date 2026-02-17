import pandas as pd


class SupportOffsetEstimator:
    """
    Estimates time offset between Monitoring and Support
    """

    def estimate(
        self,
        monitoring_df: pd.DataFrame,
        support_df: pd.DataFrame
    ) -> int:
        merged = monitoring_df.merge(
            support_df,
            on="customer_10",
            suffixes=("_mon", "_sup")
        )

        merged["delta_minutes"] = (
            merged["support_time"] - merged["monitoring_time"]
        ).dt.total_seconds() / 60

        # Robust estimator
        return int(merged["delta_minutes"].median())
