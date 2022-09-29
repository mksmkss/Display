import pandas as pd


def toArray(string):
    return string.split("|")


def get_description_list(excel_path):
    _description_list = pd.read_excel(excel_path, index_col=0, usecols=[10]).index

    description_list = []
    for i in _description_list:
        # nanはfloat64型なのでnanの判定はこうする．i=="nan"ではだめ．
        if type(i) != float:
            for j in toArray(i):
                description_list.append(j)
    print(description_list)

    return description_list


if __name__ == "__main__":
    get_description_list("/Users/masataka/Desktop/写真展フォーム　テンプレート.xlsx")
