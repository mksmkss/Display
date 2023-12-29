# from msilib.schema import Component
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


def generate_tag_pdf(excel_path, output_path, main_path):
    # わざわざsys.argv使っているのは、pyinstallerでexe化した時のエラーを回避するため
    if system == "Darwin":
        font_path = f"{main_path}/assets/ttf/MeiryoUI-03.ttf"
    else:
        font_path = f"{main_path}\\assets\\ttf\\MeiryoUI-03.ttf"
    pdfmetrics.registerFont(TTFont("Meiryo UI", font_path))

    _data_list = get_plates_list(excel_path)
    page_len = math.ceil(len(_data_list) / cards_num[0] * cards_num[1])
    print(page_len)

    isEnd = False
    for i in range(page_len):
        # iがページ数，jが各ページにおけるカード番号
        j = 0

        # A4のpdfに必要なものを描画する
        while j < cards_num[0] * cards_num[1]:
            # A4のpdfを作成する

            if system == "Darwin":
                file_name = f"{output_path}/Tag PDF/tag_{i}.pdf"
            else:
                file_name = f"{output_path}\\Tag PDF\\tag_{i}.pdf"
            page = canvas.Canvas(file_name, pagesize=A4)

            font_size = 20
            page.setFont("Meiryo UI", font_size)

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

                    title = _data_list[j + 10 * i][0]
                    penname = _data_list[j + 10 * i][1]

                    # wrap(テキストデータ,文字数)
                    title_list = textwrap.wrap(title, 10)

                    # それぞれのtitlenの位置を取得，複数行にまたがる場合も考慮
                    title_width_list = []
                    x_list = []
                    y_list = []
                    for k in range(len(title_list)):
                        # ページのフォントを指定
                        page.setFont("Meiryo UI", 25)
                        title_width_list.append(
                            round(page.stringWidth(title_list[k], "Meiryo UI", 25))
                        )

                        x_list.append(pos[0] + card[0] / 2 - title_width_list[k] / 2)
                        y_list.append(pos[1] + (card[1] / 5) * 3 - 25 * k)

                    # pennameの位置を取得
                    penname_width = round(page.stringWidth(penname, "Meiryo UI", 20))
                    x_list.append(pos[0] + card[0] / 2 - penname_width / 2)
                    y_list.append(pos[1] + card[1] / 5)

                    # タイトル
                    page.setFont("Meiryo UI", 25)
                    for k in range(len(title_list)):
                        page.drawString(x_list[k], y_list[k], title_list[k])

                    # ペンネーム
                    page.setFont("Meiryo UI", 20)
                    page.drawString(x_list[k + 1], y_list[k + 1], penname)

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
    generate_tag_pdf(
        "/Users/masataka/Desktop/写真展フォーム　テンプレート.xlsx",
        "/Users/masataka/Desktop/Data",
        "/Users/masataka/Coding/Pythons/Licosha/Display",
    )
