import pandas as pd
import platform
import json
from os import system
from PIL import Image
from reportlab.lib.pagesizes import A4

if __name__ == "__main__":
    from qrcode_generate import QRGenerator
else:
    from .qrcode_generate import QRGenerator


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


def get_description_list(excel_path):
    _description_list = pd.read_excel(excel_path, index_col=0, usecols=[9]).index
    # もし，説明文と題名等を分ける場合，空白ますは不要となるので，その場合は，二つのelseを削除すれば良い
    description_list = []
    for i in _description_list:
        # nanはfloat64型なのでnanの判定はこうする．i=="nan"ではだめ．
        if type(i) != float:
            for j in toArray(i):
                print(j, type(j) != float)
                if j != "":
                    description_list.append(j)
                else:
                    # 他の説明文はあるが，この時だけない場合
                    description_list.append("")
        else:
            # 一つも説明文がない場合
            description_list.append("")
    print(description_list)

    return description_list


def get_plates_list(excel_path):
    name_list = pd.read_excel(excel_path, index_col=0, usecols=[0]).index
    title_list = pd.read_excel(excel_path, index_col=0, usecols=[8]).index
    penname_list = pd.read_excel(excel_path, index_col=0, usecols=[10]).index

    k = 0

    # [title, penname]の集合体
    plates_list = []
    penname_to_name = {}

    print(title_list)

    while k < len(name_list):
        # 1人何個作品出したか
        works_num = len(toArray(title_list[k]))
        name = toArray(name_list[k])[0]
        for i in range(works_num):
            # pennameを記入してくれない問題児くんがいた時用
            if len(toArray(penname_list[k])) != works_num:
                plates_list.append([toArray(title_list[k])[i], penname_list[k]])
                penname_to_name[name] = penname_list[k]
            else:
                plates_list.append(
                    [toArray(title_list[k])[i], toArray(penname_list[k])[i]]
                )
                penname_to_name[name] = toArray(penname_list[k])[i]

        k += 1
    with open("assets/penname_to_name.json", mode="w", encoding="utf-8") as fp:
        json.dump(penname_to_name, fp)
    return plates_list


def generate_qr(qr_link, sns, qr_name, output_path, qr_ver=8):
    # QR versionについては，https://www.qrcode.com/about/version.htmlを参照
    QRGen = QRGenerator()
    system = platform.system()

    if system == "Darwin":
        # Mac OS
        if sns == "twitter":
            img = Image.open("assets/img/icons8-twitterx-150.png")
        elif sns == "instagram":
            img = Image.open("assets/img/icons8-instagram-150.png")
    elif system == "Windows":
        # Windows
        if sns == "twitter":
            img = Image.open("assets\img\icons8-twitterx-150.png")
        elif sns == "instagram":
            img = Image.open("assets\img\icons8-instagram-150.png")
    link = QRGen(qr_link, logo=img, qr="mono white", version=qr_ver)

    if system == "Darwin":
        link.save("{}/QRcode/{}".format(output_path, qr_name))
        print("{}/QRcode/{}".format(output_path, qr_name))
    else:
        link.save("{}\\QRcode\\{}".format(output_path, qr_name))

    return link


def get_id_list(excel_path, sns):
    _id_list = pd.read_excel(
        excel_path, index_col=0, usecols=[2 if sns == "instagram" else 3]
    ).index
    # もし，QRのみ出力する場合，空白ますは不要となるので，その場合は，二つのelseを削除すれば良い
    id_list = []
    print("{}のIDリストを取得します".format(sns))
    for i in _id_list:
        # nanはfloat64型なのでnanの判定はこうする．i=="nan"ではだめ．
        if type(i) != float:
            id_list.append([i, sns])
        else:
            # 一つも説明文がない場合
            id_list.append(["", sns])
    print(id_list)
    return id_list


def get_ids_dict(excel_path):
    # 人ごとにまとめるver
    _name_list = pd.read_excel(excel_path, index_col=0, usecols=[0]).index
    _instagram_list = pd.read_excel(excel_path, index_col=0, usecols=[2]).index
    _twitter_list = pd.read_excel(excel_path, index_col=0, usecols=[3]).index
    id_dict = {}
    length = len(_name_list)
    with open("assets/penname_to_name.json", mode="r", encoding="utf-8") as fp:
        penname_to_name_dict = json.load(fp)
        for i in range(length):
            penname = penname_to_name_dict[_name_list[i]]
            _id_list = []
            if type(_instagram_list[i]) != float:
                _id_list.append([_instagram_list[i], "instagram"])
            if type(_twitter_list[i]) != float:
                _id_list.append([_twitter_list[i], "twitter"])
            id_dict[penname] = _id_list
    print(id_dict)
    with open("assets/penname_to_sns.json", mode="w", encoding="utf-8") as fp:
        json.dump(id_dict, fp)
    return id_dict


if __name__ == "__main__":
    get_ids_dict(
        "/Users/masataka/Coding/Pythons/Licosha/Display/assets/リコシャ　2022早稲田祭展　写真収集フォーム.xlsx",
    )
#     generate_qr(qr_link="aa",sns="instagram", qr_name="test.png")
