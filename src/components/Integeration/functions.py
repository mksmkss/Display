import pandas as pd
import platform
from os import system
from PIL import Image

# if __name__ == "__main__":
#     from qrcode_generate import QRGenerator
# else:
#     from .qrcode_generate import QRGenerator

from qrcode_generate import QRGenerator


def toArray(string):
    return string.split("|")


def get_description_list(excel_path):
    _description_list = pd.read_excel(excel_path, index_col=0, usecols=[10]).index
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
    title_list = pd.read_excel(excel_path, index_col=0, usecols=[9]).index
    penname_list = pd.read_excel(excel_path, index_col=0, usecols=[11]).index

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


def generate_qr(qr_link, sns, qr_name, output_path, qr_ver=8):
    # QR versionについては，https://www.qrcode.com/about/version.htmlを参照
    QRGen = QRGenerator()
    system = platform.system()

    if system == "Darwin":
        # Mac OS
        if sns == "twitter":
            img = Image.open("assets/img/icons8-ツイッター-150.png")
        elif sns == "instagram":
            img = Image.open("assets/img/icons8-instagram-150.png")
    elif system == "Windows":
        # Windows
        if sns == "twitter":
            img = Image.open("assets\img\icons8-ツイッター-150.png")
        elif sns == "instagram":
            img = Image.open("assets\img\icons8-instagram-150.png")
    link = QRGen(qr_link, logo=img, qr="colored blue", version=qr_ver)

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


def get_ids_list(excel_path):
    # 人ごとにまとめるver
    _instagram_list = pd.read_excel(excel_path, index_col=0, usecols=[2]).index
    _twitter_list = pd.read_excel(excel_path, index_col=0, usecols=[3]).index
    # もし，QRのみ出力する場合，空白ますは不要となるので，その場合は，二つのelseを削除すれば良い
    id_list = []
    print("IDリストを取得します")
    for i in zip(_instagram_list, _twitter_list):
        _id_list = []
        # nanはfloat64型なのでnanの判定はこうする．i=="nan"ではだめ．
        if type(i[0]) != float:
            _id_list.append([i[0], "instagram"])
        if type(i[1]) != float:
            _id_list.append([i[1], "twitter"])
        if type(i[0]) == float and type(i[1]) == float:
            # 一つもQRがない場合
            _id_list.append("null")
        id_list.append(_id_list)
    print(id_list)
    return id_list


if __name__ == "__main__":
    get_ids_list(
        "/Users/masataka/Coding/Pythons/Licosha/Display/assets/リコシャ　2022早稲田祭展　写真収集フォーム.xlsx",
    )
#     generate_qr(qr_link="aa",sns="instagram", qr_name="test.png")
