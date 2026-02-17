import pandas as pd


class RateLoader:
    def load(self, path: str) -> pd.DataFrame:
        # فرض بر این است که نام شیت "Rate" است
        df = pd.read_excel(path, sheet_name="Rate")

        df = df.rename(columns={
            "مشتری": "customer_raw",
            "داخلی اپراتور": "agent_ext",
            "زمان شروع اتصال": "connect_time_raw",
            "طول تماس": "duration_time"
        })

        # اگر ستون "داخلی اپراتور" از نوع عددی است، آن را به رشته تبدیل می‌کنیم تا با Monitoring مچ شود
        if 'agent_ext' in df.columns:
            df['agent_ext'] = df['agent_ext'].astype(str).str.strip()

        return df
