import pandas as pd


def load_excel(path):
    df = pd.read_excel(path, skiprows=5, dtype=str)
    df.columns = df.columns.str.strip()
    df = df.fillna("")
    return df