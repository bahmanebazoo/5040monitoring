import pandas as pd


class MonitoringLoader:
    def load(self, path: str) -> pd.DataFrame:
        df = pd.read_excel(path, sheet_name="Monitoring")

        df = df.rename(columns={
            "مشتری": "customer_raw",
            "وضعیت": "status",
            "زمان انتظار": "wait_time",
            "داخلی": "agent_ext",
            "تاریخ": "event_time"
        })

        return df
