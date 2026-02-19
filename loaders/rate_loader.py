import pandas as pd


class RateLoader:
    def load(self, path: str) -> pd.DataFrame:
        # فرض بر این است که نام شیت "Rate" است
        df = pd.read_excel(path, sheet_name="Rate")

        if "تاریخ تماس" in df.columns:
            df["rate_jalali_date"] = df["تاریخ تماس"].copy()

        df = df.rename(columns={
            "موبایل": "customer_raw",
            "داخلی کاربر": "agent_ext",
            "تاریخ تماس": "connect_time_raw",
            "مدت زمان مکالمه": "duration_time",
            "امتیاز": "score"
        })

        # # اگر ستون "داخلی اپراتور" از نوع عددی است، آن را به رشته تبدیل می‌کنیم تا با Monitoring مچ شود
        # if 'agent_ext' in df.columns:
        #     df['agent_ext'] = df['agent_ext'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

        return df
