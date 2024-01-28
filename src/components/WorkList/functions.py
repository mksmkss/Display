import pandas as pd
import platform
import json
from os import system
from PIL import Image
from reportlab.lib.pagesizes import A4


def toArray(string):
    return string.split("|")


def to_px(mm):
    A4_mm = (210, 297)  # A4用紙のサイズをmmで指定(横、縦)
    l = A4[0] / A4_mm[0]  # mmをpxに変換
    return mm * l


def to_mm(px):
    A4_mm = (210, 297)  # A4用紙のサイズをmmで指定(横、縦)
    l = A4[0] / A4_mm[0]  # mmをpxに変換
    return px / l


def get_plates_list(excel_path):
    df = pd.read_excel(excel_path)
    name_list = df["お名前"]
    title_list = df["[写真の詳細] タイトル"]
    penname_list = df["ペンネーム"]

    k = 0

    # [title, penname]の集合体
    plates_list = []
    penname_to_name = {}

    while k < len(name_list):
        # 1人何個作品出したか
        works_num = len(toArray(title_list[k]))
        # 本名
        name = name_list[k]
        for i in range(works_num):
            plates_list.append([toArray(title_list[k])[i], penname_list[k]])
            # dictに追加
            penname_to_name[name] = penname_list[k]
        # print(penname_to_name)
        k += 1
        # get_ids_dictで使うため，JSONに出力
    # with open("assets/json/penname_to_name.json", mode="w", encoding="utf-8") as fp:
    #     json.dump(penname_to_name, fp)
    return plates_list


if __name__ == "__main__":
    # get_ids_dict(
    #     "/Users/masataka/Coding/Pythons/Licosha/Display/assets/excel/リコシャ　2023早稲田祭展　写真収集フォーム .xlsx",
    # )
    # get_plates_list(
    #     "/Users/masataka/Coding/Pythons/Licosha/Display/assets/excel/リコシャ　2023早稲田祭展　写真収集フォーム .xlsx",
    # )
    #     generate_qr(qr_link="aa",sns="instagram", qr_name="test.png")
    get_permission_dict(
        "/Users/masataka/Coding/Pythons/Licosha/Display/assets/excel/リコシャ　2023早稲田祭展　写真収集フォーム .xlsx"
    )
