# core/duplicates.py

import pandas as pd


def detect_duplicates(df: pd.DataFrame, key_columns: list[str]) -> pd.DataFrame:
    """
    Returns all rows involved in duplicates for the given key columns.
    """
    return df[df.duplicated(subset=key_columns, keep=False)]
