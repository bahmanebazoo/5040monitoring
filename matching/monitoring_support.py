import pandas as pd
from matching.confidence import confidence_score


class MonitoringSupportMatcher:
    def match(
        self,
        monitoring_df: pd.DataFrame,
        support_df: pd.DataFrame
    ) -> pd.DataFrame:

        results = []

        for _, mon in monitoring_df.iterrows():
            candidates = support_df[
                (support_df["customer_10"] == mon["customer_10"]) &
                (
                    support_df["event_time_normalized"].between(
                        mon["event_time_normalized"] - pd.Timedelta(minutes=45),
                        mon["event_time_normalized"] + pd.Timedelta(minutes=5)
                    )
                )
            ]

            for _, sup in candidates.iterrows():
                delta = (
                    sup["event_time_normalized"] -
                    mon["event_time_normalized"]
                ).total_seconds() / 60

                score = confidence_score(
                    customer_match=True,
                    agent_match=mon["agent_ext"] == sup["agent_ext"],
                    delta_minutes=delta
                )

                results.append({
                    "customer_10": mon["customer_10"],
                    "monitoring_time": mon["event_time_normalized"],
                    "support_time": sup["event_time_normalized"],
                    "agent_ext": sup["agent_ext"],
                    "delta_minutes": delta,
                    "confidence": score
                })

        return pd.DataFrame(results)
