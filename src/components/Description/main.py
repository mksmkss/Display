# from msilib.schema import Component
import sys
import math
import textwrap
import platform

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth

if __name__ == "__main__":
    from functions import *
else:
    from .functions import *


cards_num = (2, 5)  # カードをA4に何枚配置するか（横の枚数, 縦の枚数）
A4_mm = (210, 297)  # A4用紙のサイズをmmで指定(横、縦)
card_mm = (105, 59)  # 1枚のラベルのサイズをmmで指定(横、縦)
margin_mm = (0, 0)  # 余白をmmで指定(左右、上下)

to_px = A4[0] / A4_mm[0]  # mmをpxに変換
card = tuple((x * to_px for x in card_mm))  # カードのサイズpx
margin = tuple((x * to_px for x in margin_mm))  # 余白のサイズpx

system = platform.system()


def generate_description_pdf(excel_path, output_path, main_path):
    # わざわざsys.argv使っているのは、pyinstallerでexe化した時のエラーを回避するため
    if system == "Darwin":
        font_path = f"{main_path}/assets/ttf/YUMIN.TTF"
    else:
        font_path = f"{main_path}\\assets\\ttf\\MeiryoUI-03.ttf"
    pdfmetrics.registerFont(TTFont("usefont", font_path))

    _data_list = get_description_list(excel_path)
    page_len = math.ceil(len(_data_list) / cards_num[0] * cards_num[1])
    # print(page_len)

    isEnd = False
    for i in range(page_len):
        # iがページ数，jが各ページにおけるカード番号
        j = 0

        # A4のpdfに必要なものを描画する
        while j < cards_num[0] * cards_num[1]:
            # A4のpdfを作成する

            if system == "Darwin":
                file_name = f"{output_path}/Description PDF/description_{i}.pdf"

            else:
                file_name = f"{output_path}\\Description PDF\\description_{i}.pdf"

            page = canvas.Canvas(file_name, pagesize=A4)

            font_size = 16
            page.setFont("usefont", font_size)

            # 線の太さを指定
            page.setLineWidth(1)
            # 描画の初期地点
            pos = [margin[0], margin[1]]

            for y in range(cards_num[1]):
                for x in range(cards_num[0]):
                    # 図形（長方形、直線）とテキストの描画
                    page.rect(pos[0], pos[1], card[0], card[1])

                    # 最終ページはデータが途中で消えバグるので，回避
                    if j + 10 * i >= len(_data_list):
                        page.save()
                        isEnd = True
                        print("First loop is done!")
                        break

                    description = _data_list[j + 10 * i]

                    # wrap(テキストデータ,文字数)
                    description_list = textwrap.wrap(description, 16)

                    # それぞれのdescriptionの位置を取得，複数行にまたがる場合も考慮
                    description_width_list = []
                    x_list = []
                    y_list = []
                    for k in enumerate(description_list):
                        # ページのフォントを指定
                        page.setFont("usefont", font_size)
                        description_width_list.append(
                            round(
                                stringWidth(
                                    description_list[k[0]], "usefont", font_size
                                )
                            )
                        )

                        x_list.append(
                            pos[0] + card[0] / 2 - description_width_list[k[0]] / 2
                        )
                        y_list.append(
                            pos[1]
                            + (card[1] / 2)
                            + 12.5 * (len(description_list) - 2)
                            - font_size * k[0]
                        )

                    page.setFont("usefont", font_size)
                    for k in enumerate(description_list):
                        page.drawString(
                            x_list[k[0]], y_list[k[0]], description_list[k[0]]
                        )

                    j += 1
                    pos[0] += card[0]
                if isEnd:
                    print("Second loop is done!")
                    break

                pos[0] = margin[0]
                pos[1] += card[1]

            if isEnd:
                print("Third loop is done!")
                break

        if isEnd is False:
            # PDFファイルを保存
            page.save()
        else:
            print("Fourth loop is done!")
            break


if __name__ == "__main__":
    generate_description_pdf(
        "/Users/masataka/Desktop/リコシャ　2022早稲田祭展　写真収集フォーム.xlsx",
        "/Users/masataka/Desktop/Data",
        "/Users/masataka/Coding/Pythons/Licosha/Display",
    )
