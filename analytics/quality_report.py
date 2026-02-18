import pandas as pd
from dataclasses import dataclass

@dataclass
class QualityMetrics:
    total_monitoring: int
    total_support: int
    total_matches: int
    match_rate: float
    confidence_breakdown: pd.Series
    delta_stats: pd.DataFrame

def compute_quality_metrics(
    monitoring_df: pd.DataFrame,
    support_df: pd.DataFrame,
    matches_df: pd.DataFrame
) -> QualityMetrics:
    total_monitoring = len(monitoring_df)
    total_support = len(support_df)
    total_matches = len(matches_df)

    match_rate = total_matches / total_monitoring if total_monitoring else 0

    confidence_breakdown = (
        matches_df['confidence']
        .fillna('unknown')
        .value_counts(normalize=True)
        .rename('ratio')
    )

    delta_series = matches_df['delta_minutes'].dropna()
    if delta_series.empty:
        delta_stats = pd.DataFrame()
    else:
        delta_stats = pd.DataFrame({
            'metric': [
                'mean', 'median', 'std', 'min', 'max',
                'p10', 'p25', 'p50', 'p75', 'p90', 'p95'
            ],
            'value': [
                delta_series.mean(),
                delta_series.median(),
                delta_series.std(),
                delta_series.min(),
                delta_series.max(),
                delta_series.quantile(0.10),
                delta_series.quantile(0.25),
                delta_series.quantile(0.50),
                delta_series.quantile(0.75),
                delta_series.quantile(0.90),
                delta_series.quantile(0.95)
            ]
        })

    return QualityMetrics(
        total_monitoring=total_monitoring,
        total_support=total_support,
        total_matches=total_matches,
        match_rate=match_rate,
        confidence_breakdown=confidence_breakdown,
        delta_stats=delta_stats
    )
