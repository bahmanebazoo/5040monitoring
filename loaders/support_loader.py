import pandas as pd


class SupportLoader:
    def load(self, path: str) -> pd.DataFrame:
        # فرض بر این است که نام شیت "Support" است
        df = pd.read_excel(path, sheet_name="Support")

        if "تاریخ ایجاد فعالیت" in df.columns:
            df["support_jalali_date"] = df["تاریخ ایجاد فعالیت"].copy()

        df = df.rename(columns={
            "موبایل": "customer_raw",
            "پشتیبان": "agent_name",
            "تاریخ ایجاد فعالیت": "event_time"
        })

        # حذف هرگونه سطر با نام پشتیبان خالی که قابل مپ شدن نیست
        df.dropna(subset=['agent_name'], inplace=True)

        return df
