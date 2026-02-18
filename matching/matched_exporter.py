import pandas as pd
from pathlib import Path

def export_matches(
    matches_df: pd.DataFrame,
    metrics,
    output_path: Path,
    as_excel: bool = True
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if as_excel:
        with pd.ExcelWriter(output_path) as writer:
            matches_df.to_excel(writer, index=False, sheet_name="matches")

            summary_df = pd.DataFrame({
                'metric': [
                    'total_monitoring', 'total_support',
                    'total_matches', 'match_rate'
                ],
                'value': [
                    metrics.total_monitoring,
                    metrics.total_support,
                    metrics.total_matches,
                    metrics.match_rate
                ]
            })
            summary_df.to_excel(writer, index=False, sheet_name="summary")

            metrics.confidence_breakdown.reset_index().rename(
                columns={'index': 'confidence'}
            ).to_excel(writer, index=False, sheet_name="confidence_breakdown")

            if not metrics.delta_stats.empty:
                metrics.delta_stats.to_excel(
                    writer, index=False, sheet_name="delta_stats"
                )
    else:
        matches_df.to_csv(output_path, index=False)

        summary_df = pd.DataFrame({
            'metric': [
                'total_monitoring', 'total_support',
                'total_matches', 'match_rate'
            ],
            'value': [
                metrics.total_monitoring,
                metrics.total_support,
                metrics.total_matches,
                metrics.match_rate
            ]
        })
        summary_df.to_csv(output_path.with_suffix('.summary.csv'), index=False)

        metrics.confidence_breakdown.reset_index().rename(
            columns={'index': 'confidence'}
        ).to_csv(output_path.with_suffix('.confidence.csv'), index=False)

        if not metrics.delta_stats.empty:
            metrics.delta_stats.to_csv(
                output_path.with_suffix('.delta.csv'), index=False
            )
