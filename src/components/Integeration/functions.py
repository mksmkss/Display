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
    df = pd.read_excel(excel_path)
    _description_list = df["[写真の詳細] 説明"]
    # もし，説明文と題名等を分ける場合，空白ますは不要となるので，その場合は，二つのelseを削除すれば良い
    description_list = []
    for i in _description_list:
        """
        nanはfloat64型なのでnanの判定はこうする．i=="nan"ではだめ．また，np.float64(i)だとcould not convert string to float: '|'
        というエラーが出る．
        """
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
    with open("assets/json/penname_to_name.json", mode="w", encoding="utf-8") as fp:
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


def get_ids_dict(excel_path):
    # 人ごとにsnsをまとめるver
    df = pd.read_excel(excel_path)
    _name_list = df["お名前"]
    _instagram_list = df["Instagramのアカウント"]
    _twitter_list = df["Xのアカウント"]
    ids_dict = {}
    length = len(_name_list)
    with open("assets/json/penname_to_name.json", mode="r", encoding="utf-8") as fp:
        penname_to_name_dict = json.load(fp)
        for i in range(length):
            penname = penname_to_name_dict[_name_list[i]]
            _id_list = []
            if type(_instagram_list[i]) != float:
                _id_list.append([_instagram_list[i], "instagram"])
            if type(_twitter_list[i]) != float:
                _id_list.append([_twitter_list[i], "twitter"])
            ids_dict[penname] = _id_list
    # 他のところ（QR）で使うため，一度JSONに変換する．
    with open("assets/json/penname_to_sns.json", mode="w", encoding="utf-8") as fp:
        json.dump(ids_dict, fp)
    return ids_dict


def get_permission_dict(excel_path):
    df = pd.read_excel(excel_path)
    penname_list = df["ペンネーム"]
    _permission_list = df["リコシャのHPやSNSに載せて良いか"]
    permission_dict = {}
    for i, penname in enumerate(penname_list):
        permission_dict[penname] = _permission_list[i]
    return permission_dict


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
