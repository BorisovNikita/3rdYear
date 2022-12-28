import pandas as pd

def mapp(s: pd.Series, f: callable) -> pd.Series:
    return s.map(f)
