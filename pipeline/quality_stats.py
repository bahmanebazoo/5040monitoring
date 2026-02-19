"""محاسبه آمار کیفیت مچینگ (summary, confidence, delta)."""

import pandas as pd
from dataclasses import dataclass
from typing import Tuple


@dataclass
class Stats:
    """آمار یک مچینگ."""
    summary: pd.DataFrame
    confidence: pd.DataFrame
    delta: pd.DataFrame


class QualityStats:
    """محاسبه آمار کیفیت برای یک DataFrame مچینگ."""

    @staticmethod
    def compute(matches_df: pd.DataFrame, label: str) -> Stats:
        matched = matches_df[matches_df["match_status"] == "matched"]
        total = len(matches_df)
        n_match = len(matched)

        summary = pd.DataFrame({
            "metric": [
                f"{label}_total_rows",
                f"{label}_matched",
                f"{label}_match_rate",
            ],
            "value": [
                total,
                n_match,
                round(n_match / total, 4) if total else 0,
            ],
        })

        # توزیع confidence
        confidence = (
            matches_df["confidence"]
            .value_counts()
            .sort_index()
            .rename("count")
            .reset_index()
        )
        confidence.columns = ["confidence", "count"]

        # آمار delta
        ds = matched["delta_minutes"].dropna()
        if ds.empty:
            delta = pd.DataFrame(columns=["metric", "value"])
        else:
            delta = pd.DataFrame({
                "metric": [
                    "mean", "median", "std", "min", "max",
                    "p10", "p25", "p75", "p90", "p95",
                ],
                "value": [
                    ds.mean(), ds.median(), ds.std(), ds.min(), ds.max(),
                    ds.quantile(.10), ds.quantile(.25),
                    ds.quantile(.75), ds.quantile(.90), ds.quantile(.95),
                ],
            })

        return Stats(summary=summary, confidence=confidence, delta=delta)
