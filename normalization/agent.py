import pandas as pd


def map_agent_name_to_ext(
    support_df: pd.DataFrame,
    mapping_df: pd.DataFrame
) -> pd.DataFrame:

    return support_df.merge(
        mapping_df,
        how="left",
        left_on="agent_name",
        right_on="agent_name"
    )
