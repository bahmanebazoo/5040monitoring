import pandas as pd
from matching.confidence import confidence_score


class MonitoringRateMatcher:
    """
    Greedy one-to-one matching بین Monitoring و Rate.

    زمان اتصال مونیتورینگ = event_time_normalized + wait_seconds
    سپس مقایسه با event_time_normalized در Rate صورت می‌گیرد.

    آستانه پیش‌فرض: 85
    """

    def __init__(self, threshold: int = 85):
        self.threshold = threshold

    def match(
        self,
        monitoring_df: pd.DataFrame,
        rate_df: pd.DataFrame,
        min_confidence: int = None,
    ) -> pd.DataFrame:

        threshold = min_confidence if min_confidence is not None else self.threshold

        # ──────────────────────────────────────────────
        # ① محاسبه زمان اتصال در Monitoring
        #    connect_time = event_time + wait_seconds
        # ──────────────────────────────────────────────
        monitoring_work = monitoring_df.copy()

        # اطمینان از اینکه wait_seconds عددی است
        monitoring_work["wait_seconds"] = pd.to_numeric(
            monitoring_work["wait_seconds"], errors="coerce"
        ).fillna(0)

        monitoring_work["connect_time"] = (
            monitoring_work["event_time_normalized"]
            + pd.to_timedelta(monitoring_work["wait_seconds"], unit="s")
        )

        # مرتب‌سازی: قدیمی‌ترین اول
        monitoring_sorted = monitoring_work.sort_values(
            "connect_time"
        ).reset_index(drop=True)

        # مجموعه ایندکس‌های Rate مصرف‌شده
        used_rate_indices: set = set()

        matched_results: list = []
        unmatched_monitoring: list = []

        for _, mon in monitoring_sorted.iterrows():

            mon_connect_time = mon["connect_time"]

            # ──────────────────────────────────────────────
            # ② فیلتر کاندیداها:
            #    شماره مشتری یکسان
            #    + بازه زمانی حول connect_time (نه event_time)
            #    + مصرف‌نشده
            # ──────────────────────────────────────────────
            candidates = rate_df[
                (rate_df["customer_10"] == mon["customer_10"])
                & (
                    rate_df["event_time_normalized"].between(
                        mon_connect_time - pd.Timedelta(minutes=5),
                        mon_connect_time + pd.Timedelta(minutes=45),
                    )
                )
                & (~rate_df.index.isin(used_rate_indices))
            ]

            if candidates.empty:
                unmatched_monitoring.append(
                    {
                        "customer_10": mon["customer_10"],
                        "monitoring_event_time": mon["event_time_normalized"],
                        "monitoring_connect_time": mon_connect_time,
                        "monitoring_agent_ext": mon.get("agent_ext"),
                        "monitoring_status": mon.get("status"),
                        "monitoring_wait_seconds": mon.get("wait_seconds"),
                        "rate_time": pd.NaT,
                        "rate_agent_ext": None,
                        "rate_duration_seconds": None,
                        "delta_minutes": None,
                        "confidence": 0,
                        "match_status": "no_candidate",
                    }
                )
                continue

            # ──────────────────────────────────────────────
            # ③ محاسبه امتیاز بر اساس اختلاف
            #    rate_time - connect_time
            # ──────────────────────────────────────────────
            scored = []

            for rate_idx, rate_row in candidates.iterrows():
                delta = (
                    rate_row["event_time_normalized"] - mon_connect_time
                ).total_seconds() / 60.0

                score = confidence_score(
                    customer_match=True,
                    agent_match=(
                        str(mon.get("agent_ext", "")).strip()
                        == str(rate_row.get("agent_ext", "")).strip()
                    ),
                    delta_minutes=delta,
                )
                scored.append((rate_idx, score, delta))

            # بالاترین امتیاز → کمترین |delta|
            scored.sort(key=lambda x: (-x[1], abs(x[2])))

            best_idx, best_score, best_delta = scored[0]

            if best_score >= threshold:
                best_rate = rate_df.loc[best_idx]
                used_rate_indices.add(best_idx)

                matched_results.append(
                    {
                        "customer_10": mon["customer_10"],
                        "monitoring_event_time": mon["event_time_normalized"],
                        "monitoring_connect_time": mon_connect_time,
                        "monitoring_agent_ext": mon.get("agent_ext"),
                        "monitoring_status": mon.get("status"),
                        "monitoring_wait_seconds": mon.get("wait_seconds"),
                        "rate_time": best_rate["event_time_normalized"],
                        "rate_agent_ext": best_rate["agent_ext"],
                        "rate_duration_seconds": best_rate.get("duration_seconds"),
                        "rate_score": best_rate.get("score"),
                        "delta_minutes": best_delta,
                        "confidence": best_score,
                        "match_status": "matched",
                        "rate_jalali_date": best_rate.get("rate_jalali_date"),
                    }
                )
            else:
                unmatched_monitoring.append(
                    {
                        "customer_10": mon["customer_10"],
                        "monitoring_event_time": mon["event_time_normalized"],
                        "monitoring_connect_time": mon_connect_time,
                        "monitoring_agent_ext": mon.get("agent_ext"),
                        "monitoring_status": mon.get("status"),
                        "monitoring_wait_seconds": mon.get("wait_seconds"),
                        "rate_time": pd.NaT,
                        "rate_agent_ext": None,
                        "rate_duration_seconds": None,
                        "rate_score": None,
                        "delta_minutes": None,
                        "confidence": best_score,
                        "match_status": f"below_threshold (best={best_score})",
                        "rate_jalali_date": None,
                    }
                )

        # ترکیب نتایج
        all_rows = matched_results + unmatched_monitoring
        result_df = pd.DataFrame(all_rows)

        if not result_df.empty:
            result_df.sort_values("monitoring_connect_time", inplace=True)
            result_df.reset_index(drop=True, inplace=True)

        return result_df
