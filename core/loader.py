import pandas as pd

def load_excel(path, sheet):
    return pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
