
import pandas as pd

def get_tag_sum_count_from_file(file: str) -> dict:
    df = pd.read_csv(file)
    df2 = df.groupby("tags")["n_steps"].agg(sum = 'sum', count = 'count')
    return df2.to_dict('index')
