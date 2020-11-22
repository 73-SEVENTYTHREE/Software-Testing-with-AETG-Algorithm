import pandas as pd


def read_excel1(filename, col_number):
    df = pd.read_excel(filename, usecols=[col_number], names=None)
    df = df.dropna(axis=0, how='any')
    df_li = df.values.tolist()
    result = []
    for s_li in df_li:
        result.append(s_li[0])
    return result

