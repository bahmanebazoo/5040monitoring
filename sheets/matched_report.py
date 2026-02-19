"""Export گزارش مچینگ Support + Rate به اکسل."""

import pandas as pd
from pipeline.quality_stats import QualityStats


class MatchedReportExporter:
    """تولید matched_report.xlsx با ۹ شیت."""

    def export(
        self,
        support_matches: pd.DataFrame,
        rate_matches: pd.DataFrame,
        output_path: str,
    ) -> str:
        print("\n" + "=" * 60)
        print("STEP 5 — Matched report (Support + Rate)")
        print("=" * 60)

        sup = QualityStats.compute(support_matches, "support")
        rate = QualityStats.compute(rate_matches, "rate")

        combined_summary = pd.concat(
            [sup.summary, rate.summary], ignore_index=True
        )

        with pd.ExcelWriter(output_path, engine="openpyxl") as w:
            support_matches.to_excel(w, index=False, sheet_name="support_matches")
            sup.summary.to_excel(w, index=False, sheet_name="support_summary")
            sup.confidence.to_excel(w, index=False, sheet_name="support_confidence")
            sup.delta.to_excel(w, index=False, sheet_name="support_delta")

            rate_matches.to_excel(w, index=False, sheet_name="rate_matches")
            rate.summary.to_excel(w, index=False, sheet_name="rate_summary")
            rate.confidence.to_excel(w, index=False, sheet_name="rate_confidence")
            rate.delta.to_excel(w, index=False, sheet_name="rate_delta")

            combined_summary.to_excel(w, index=False, sheet_name="combined_summary")

        print(f"   ✅ Saved → {output_path}  (9 sheets)")
        return output_path
