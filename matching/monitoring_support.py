import pandas as pd
from matching.confidence import confidence_score


class MonitoringSupportMatcher:
    """
    Greedy one-to-one matching بین Monitoring و Support.
    آستانه پیش‌فرض: 75
    """

    def __init__(self, threshold: int = 75):
        self.threshold = threshold

    def match(
        self,
        monitoring_df: pd.DataFrame,
        support_df: pd.DataFrame,
        min_confidence: int = None,
    ) -> pd.DataFrame:

        # اگر min_confidence داده شده از آن استفاده کن، وگرنه از self.threshold
        threshold = min_confidence if min_confidence is not None else self.threshold

        # مرتب‌سازی مونیتورینگ: قدیمی‌ترین اول
        monitoring_sorted = monitoring_df.sort_values(
            "event_time_normalized"
        ).reset_index(drop=True)

        # مجموعه ایندکس‌های Support مصرف‌شده
        used_support_indices: set = set()

        matched_results: list = []
        unmatched_monitoring: list = []

        for _, mon in monitoring_sorted.iterrows():

            # فیلتر کاندیداها
            candidates = support_df[
                (support_df["customer_10"] == mon["customer_10"])
                & (
                    support_df["event_time_normalized"].between(
                        mon["event_time_normalized"] - pd.Timedelta(minutes=45),
                        mon["event_time_normalized"] + pd.Timedelta(minutes=5),
                    )
                )
                & (~support_df.index.isin(used_support_indices))
            ]

            if candidates.empty:
                unmatched_monitoring.append(
                    {
                        "customer_10": mon["customer_10"],
                        "monitoring_time": mon["event_time_normalized"],
                        "monitoring_agent_ext": mon.get("agent_ext"),
                        "monitoring_status": mon.get("status"),
                        "monitoring_wait_seconds": mon.get("wait_seconds"),
                        "support_time": pd.NaT,
                        "support_agent_ext": None,
                        "delta_minutes": None,
                        "confidence": 0,
                        "match_status": "no_candidate",
                    }
                )
                continue

            # محاسبه امتیاز
            scored = []

            for sup_idx, sup in candidates.iterrows():
                delta = (
                    sup["event_time_normalized"]
                    - mon["event_time_normalized"]
                ).total_seconds() / 60.0

                score = confidence_score(
                    customer_match=True,
                    agent_match=(
                        str(mon.get("agent_ext", "")).strip()
                        == str(sup.get("agent_ext", "")).strip()
                    ),
                    delta_minutes=delta,
                )
                scored.append((sup_idx, score, delta))

            # بالاترین امتیاز → کمترین |delta|
            scored.sort(key=lambda x: (-x[1], abs(x[2])))

            best_idx, best_score, best_delta = scored[0]

            if best_score >= threshold:
                best_sup = support_df.loc[best_idx]
                used_support_indices.add(best_idx)

                matched_results.append(
                    {
                        "customer_10": mon["customer_10"],
                        "monitoring_time": mon["event_time_normalized"],
                        "monitoring_agent_ext": mon.get("agent_ext"),
                        "monitoring_status": mon.get("status"),
                        "monitoring_wait_seconds": mon.get("wait_seconds"),
                        "support_time": best_sup["event_time_normalized"],
                        "support_agent_ext": best_sup["agent_ext"],
                        "delta_minutes": best_delta,
                        "confidence": best_score,
                        "match_status": "matched",
                    }
                )
            else:
                unmatched_monitoring.append(
                    {
                        "customer_10": mon["customer_10"],
                        "monitoring_time": mon["event_time_normalized"],
                        "monitoring_agent_ext": mon.get("agent_ext"),
                        "monitoring_status": mon.get("status"),
                        "monitoring_wait_seconds": mon.get("wait_seconds"),
                        "support_time": pd.NaT,
                        "support_agent_ext": None,
                        "delta_minutes": None,
                        "confidence": best_score,
                        "match_status": f"below_threshold (best={best_score})",
                    }
                )

        # ترکیب
        all_rows = matched_results + unmatched_monitoring
        result_df = pd.DataFrame(all_rows)

        if not result_df.empty:
            result_df.sort_values("monitoring_time", inplace=True)
            result_df.reset_index(drop=True, inplace=True)

        return result_df
