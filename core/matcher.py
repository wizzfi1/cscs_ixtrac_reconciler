from core.normalizer import normalize_name, first_two_names

def build_cscs_index(df):
    index = {}
    for _, row in df.iterrows():
        key = (normalize_name(row["NAME"]), row["CHN"])
        index.setdefault(key, []).append(row)
    return index

def build_cscs_index_2name(df):
    index = {}
    for _, row in df.iterrows():
        key = (first_two_names(row["NAME"]), row["CHN"])
        index.setdefault(key, []).append(row)
    return index
