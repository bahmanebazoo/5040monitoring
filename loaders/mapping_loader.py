import pandas as pd


class MappingLoader:
    def load(self, path: str) -> pd.DataFrame:
        # فرض بر این است که نام شیت "Mapping" است
        df = pd.read_excel(path, sheet_name="exp")

        df = df.rename(columns={
            "نام": "agent_name",
            "داخلی": "agent_ext"
        })

        # تمیزکاری داخلی و نام
        if 'agent_ext' in df.columns:
            df['agent_ext'] = df['agent_ext'].astype(str).str.strip()
        if 'agent_name' in df.columns:
            df['agent_name'] = df['agent_name'].astype(str).str.strip()

        return df
