import pandas as pd
import math


def toArray(string):
    return string.split("|")


def get_plates_list(excel_path):
    df = pd.read_excel(excel_path)
    name_list = df["お名前"]
    title_list = df["[写真の詳細] タイトル"]
    penname_list = df["ペンネーム"]

    k = 0

    # [title, penname]の集合体
    plates_list = []

    while k < len(name_list):
        # 1人何個作品出したか
        works_num = len(toArray(title_list[k]))
        for i in range(works_num):
            # pennameを記入してくれない問題児くんがいた時用
            if len(toArray(penname_list[k])) != works_num:
                plates_list.append([toArray(title_list[k])[i], penname_list[k]])
            else:
                plates_list.append(
                    [toArray(title_list[k])[i], toArray(penname_list[k])[i]]
                )

        k += 1
    return plates_list
