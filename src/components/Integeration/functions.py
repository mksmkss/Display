import pandas as pd
import numpy as np
import platform
import json
from os import system
from PIL import Image
import uuid
import pylibdmtx.pylibdmtx as dmtx
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
                # print(j, type(j) != float)
                if j != "":
                    description_list.append(j)
                else:
                    # 他の説明文はあるが，この時だけない場合
                    description_list.append("")
        else:
            # 一つも説明文がない場合
            description_list.append("")
    # print(description_list)

    return description_list


def get_plates_list(excel_path):
    df = pd.read_excel(excel_path)

    # uuidの列が存在しない場合は追加する
    if "uuid" not in df.columns:
        df["uuid"] = ""

    name_list = df["お名前"]
    title_list = df["[写真の詳細] タイトル"]
    penname_list = df["ペンネーム"]

    k = 0
    plates_list = []
    penname_to_name = {}

    while k < len(name_list):
        works_num = len(toArray(title_list[k]))
        name = name_list[k]

        # 既存のUUIDを取得
        existing_uuids = df.loc[k, "uuid"] if pd.notna(df.loc[k, "uuid"]) else ""
        existing_uuid_list = existing_uuids.split("|") if existing_uuids else []

        # 必要な分だけ新しいUUIDを生成
        new_uuid_count = works_num - len(existing_uuid_list)
        new_uuids = [str(uuid.uuid4()) for _ in range(new_uuid_count)]

        # 既存のUUIDと新しいUUIDを結合
        uuid_list = existing_uuid_list + new_uuids

        for i in range(works_num):
            plates_list.append([toArray(title_list[k])[i], penname_list[k]])
            penname_to_name[name] = penname_list[k]

        # 更新されたUUIDリストをDataFrameに設定
        df.loc[k, "uuid"] = "|".join(
            uuid_list[:works_num]
        )  # works_numまでのUUIDのみを使用

        k += 1

    with open("assets/json/penname_to_name.json", mode="w", encoding="utf-8") as fp:
        json.dump(penname_to_name, fp)

    # 変更を保存
    df.to_excel(excel_path, index=False)
    # print(plates_list)
    return plates_list


def get_uuid_list(excel_path):
    df = pd.read_excel(excel_path)
    uuid_series = df["uuid"]

    def process_uuid(uuid):
        if pd.isna(uuid) or uuid == "":
            return [""]
        else:
            return uuid.split("|")

    # 各UUIDエントリを処理し、結果をリストのリストとして取得
    uuid_nested_list = uuid_series.apply(process_uuid).tolist()

    # ネストされたリストを平坦化
    uuid_list = [uuid for sublist in uuid_nested_list for uuid in sublist]

    return uuid_list


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


def generate_data_matrix(data, output_path, size=24):
    """
    文字列からデータマトリックスを生成し、画像として保存する

    :param data: エンコードする文字列
    :param size: データマトリックスのサイズ (デフォルト: 24)
    :param output_path: 出力画像のパス
    """

    system = platform.system()
    # データマトリックスのエンコード
    encoded = dmtx.encode(data.encode("utf-8"), size=f"{size}x{size}")

    # エンコードされたデータをPIL Imageに変換
    pil_img = Image.frombytes("RGB", (encoded.width, encoded.height), encoded.pixels)

    # グレースケールに変換
    pil_img = pil_img.convert("L")

    # NumPy配列に変換
    img = np.array(pil_img)

    # 画像の拡大（見やすくするため）
    scale = 10
    img = np.repeat(np.repeat(img, scale, axis=0), scale, axis=1)

    # 白黒を#2c2c2eに変換
    img = np.where(img == 0, 44, 255).astype(np.uint8)  # uint8型に明示的に変換

    # PILイメージに再変換
    pil_img = Image.fromarray(img)

    # カラーモードをRGBに変換し、グレーを#2c2c2eに
    pil_img = pil_img.convert("RGB")
    pixels = pil_img.load()
    width, height = pil_img.size
    for x in range(width):
        for y in range(height):
            if pixels[x, y] == (44, 44, 44):  # 0x2cは10進数で44
                pixels[x, y] = (44, 44, 46)  # #2c2c2e

    # 画像の保存
    if system == "Darwin":
        pil_img.save(output_path)
    else:
        pil_img.save(output_path.replace("/", "\\"))


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

            # Instagramアカウントのチェック
            if not pd.isna(_instagram_list[i]):
                _id_list.append([_instagram_list[i], "instagram"])

            # Twitterアカウントのチェック
            if not pd.isna(_twitter_list[i]):
                _id_list.append([_twitter_list[i], "twitter"])

            ids_dict[penname] = _id_list

    # 他のところ（QR）で使うため，一度JSONに変換する．
    with open("assets/json/penname_to_sns.json", mode="w", encoding="utf-8") as fp:
        json.dump(ids_dict, fp, ensure_ascii=False, indent=2)

    return ids_dict


def get_permission_dict(excel_path):
    df = pd.read_excel(excel_path)
    penname_list = df["ペンネーム"]
    _permission_list = df["来場者が撮影可能か"]
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
    get_uuid_list(
        "/Users/masataka/Coding/Pythons/Licosha/Display/assets/excel/リコシャ　2023早稲田祭展　写真収集フォーム .xlsx"
    )
    #     generate_qr(qr_link="aa",sns="instagram", qr_name="test.png")
    # get_permission_dict(
    #     "/Users/masataka/Coding/Pythons/Licosha/Display/assets/excel/リコシャ　2023早稲田祭展　写真収集フォーム .xlsx"
    # )
    # generate_data_matrix("test", output_path="test.png")
