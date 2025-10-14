# helper functions (extend as needed)
import pandas as pd

def normalize_columns(df):
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    return df
