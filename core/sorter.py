from config.rules import STATUS_PRIORITY

def sort_dataframe(df):
    df["__rank"] = df["MATCH_STATUS"].map(STATUS_PRIORITY)
    df = df.sort_values("__rank").drop(columns="__rank")
    return df
