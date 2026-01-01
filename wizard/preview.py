import pandas as pd
from core.normalizer import normalize_name

def generate_preview(file_path, sheet, name_col, chn_col, rows=5):
    df = pd.read_excel(file_path, sheet_name=sheet)
    preview = df[[name_col, chn_col]].head(rows).copy()
    preview["NORMALIZED_NAME"] = preview[name_col].apply(normalize_name)
    return preview
