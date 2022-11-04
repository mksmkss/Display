import math
import pandas as pd
import platform
from os import system
from PIL import Image

if __name__ == "__main__":
    from qrcode_generate import QRGenerator
else:
    from .qrcode_generate import QRGenerator


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
    id_list = []
    print("{}のIDリストを取得します".format(sns))
    for i in _id_list:
        # nanはfloat64型なのでnanの判定はこうする．i=="nan"ではだめ．
        if type(i) != float:
            id_list.append([i, sns])
    print(id_list)
    return id_list


# if __name__ == "__main__":
#     get_id_list("写真展フォーム　テンプレート.xlsx","instagram")
#     generate_qr(qr_link="aa",sns="instagram", qr_name="test.png")
