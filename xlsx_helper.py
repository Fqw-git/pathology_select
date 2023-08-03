import pandas
import pandas as pd

scanned = pd.read_excel(r"scanned.xlsx")
scanned = scanned.values
scanned_id = [str(o[0]).lower() for o in scanned]

total_range = range(1803130, 2219852)

range_cabinet = [
    range(1803130, 1916309),
    range(1916309, 2104349),
    range(2104349, 2208877),
    range(2208877, 2219852)
]


def select_no_scanned(i):
    s = str(i).lower()
    return s not in scanned_id


def select_format(i):
    s = str(i).lower()
    head = s[0]
    return head == 's' or head == '1'


def select_range(i):
    s = str(i)
    v = int(s[1:])
    return v in total_range


def select(i):
    return select_format(i) and select_no_scanned(i) and select_range(i)


def pathology2value(s):
    return int(str(s)[1:])


def sort_key_fn(s: pd.Series):
    return s.apply(lambda s: int(str(s)[1:]))


def cabinet_localize(v, start=0):
    for i in range(start, len(range_cabinet)):
        if v in range_cabinet[i]:
            return i


def add_cabinet(df: pandas.DataFrame):
    for i, idx in enumerate(df.index):
        v = df.iloc[i]["病理号"]
        v = pathology2value(v)
        this_cabinet = cabinet_localize(v)
        df.loc[idx, "储藏柜"] = this_cabinet
    return df


def add_column(df, disease):
    df["疾病"] = [disease] * len(df.index)
    df["是否有切片"] = [float('nan')] * len(df.index)
    df["储藏柜"] = [float('nan')] * len(df.index)
    df["切片数量"] = [float('nan')] * len(df.index)
    return df


def process_disease(disease):
    df = pd.read_excel(r"all.xlsx", sheet_name=disease)
    selected = df.loc[df["病理号"].apply(select)]
    selected = selected.sort_values(by=["病理号"], key=sort_key_fn)
    ret = add_column(selected, disease)
    ret = add_cabinet(ret)
    ret = ret[["疾病", "储藏柜", "病理号", "是否有切片", "切片数量", "病理诊断"]]
    return ret


def process_diseases(disease_list):
    df_list = [process_disease(d) for d in disease_list]
    ret = pd.concat(df_list)
    ret = ret.sort_values(by="病理号", key=sort_key_fn)
    return ret


if __name__ == '__main__':
    ret = process_diseases(["胶质母"])
    ret.to_excel("result/111.xlsx", index=False)
