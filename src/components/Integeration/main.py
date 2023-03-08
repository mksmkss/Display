import math
import textwrap
import platform

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

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


def generate_caption_pdf(excel_path, output_path, main_path):
    # わざわざsys.argv使っているのは、pyinstallerでexe化した時のエラーを回避するため
    if system == "Darwin":
        font_path = f"{main_path}/assets/MeiryoUI-03.ttf"
    else:
        font_path = f"{main_path}\\assets\\MeiryoUI-03.ttf"
    # pdfmetrics.registerFont(TTFont("usefont", UnicodeCIDFont("HeiseiMin-W3")))
    pdfmetrics.registerFont(TTFont("usefont", font_path))

    _plates_list = get_plates_list(excel_path)
    _description_list = get_description_list(excel_path)
    page_len = math.ceil(len(_plates_list) / cards_num[0] * cards_num[1])
    print(page_len)

    isEnd = False
    for i in range(page_len):

        # iがページ数，jが各ページにおけるカード番号
        j = 0

        # A4のpdfに必要なものを描画する
        while j < cards_num[0] * cards_num[1]:

            # A4のpdfを作成する

            if system == "Darwin":
                file_name = f"{output_path}/Caption PDF/caption_{i}.pdf"
            else:
                file_name = f"{output_path}\\Caption PDF\\caption_{i}.pdf"
            page = canvas.Canvas(file_name, pagesize=A4)

            title_size = 16
            penname_size = 13
            description_size = 12

            # 線の太さを指定
            page.setLineWidth(1)
            # 描画の初期地点
            pos = [margin[0], margin[1]]

            for y in range(cards_num[1]):
                for x in range(cards_num[0]):
                    # 図形（長方形、直線）とテキストの描画
                    page.rect(pos[0], pos[1], card[0], card[1])

                    # # 下の黒いところの描画
                    # page.setFillColorRGB(44, 44, 46)
                    # page.rect(pos[0], pos[1], card[0], card[1] / 4.2, fill=True)

                    # 最終ページはデータが途中で消えバグるので，回避
                    if j + 10 * i >= len(_plates_list):
                        page.save()
                        isEnd = True
                        print("First loop is done!")
                        break

                    title = _plates_list[j + 10 * i][0]
                    penname = _plates_list[j + 10 * i][1]
                    description = _description_list[j + 10 * i]

                    # wrap(テキストデータ,文字数)
                    # title_list = textwrap.wrap(title, 10)
                    description_list = textwrap.wrap(description, 16)

                    # それぞれのtitleの位置を取得，複数行にまたがる場合も考慮
                    title_width_list = []
                    title_x = []
                    title_y = []

                    description_width_list = []
                    description_x = []
                    description_y = []

                    # titleは複数行を考慮しない
                    page.setFont("usefont", title_size)
                    title_width_list.append(
                        round(page.stringWidth(title, "usefont", title_size))
                    )
                    # 座標の基準は文字の左下
                    title_x.append(pos[0] + card[0] * 0.08)
                    title_y.append(pos[1] + card[1] * 0.8)

                    for k in enumerate(description_list):
                        # description
                        page.setFont("usefont", description_size)
                        description_x.append(pos[0] + card[0] * 0.15)
                        description_y.append(
                            pos[1]
                            + (card[1] / 2)
                            + (description_size / 2) * (len(description_list) - 2)
                            - description_size * k[0]
                        )

                    # pennameの位置を取得
                    penname_width = round(
                        page.stringWidth(penname, "usefont", penname_size)
                    )
                    # titke_x,yを流用
                    title_x.append(pos[0] + card[0] - card[0] * 0.08)
                    title_y.append(pos[1] + card[1] * 0.08)

                    # titleの描画
                    # page.setFillColorRGB(44, 44, 46)
                    page.setFont("usefont", title_size)
                    page.drawString(title_x[0], title_y[0], title)

                    # pennameの描画
                    page.setFillColorRGB(237, 237, 235)
                    page.setFont("usefont", penname_size)
                    page.drawRightString(title_x[1], title_y[1], penname)

                    # descriptionの描画
                    page.setFillColorRGB(44, 44, 46)
                    page.setFont("usefont", description_size)
                    for k in enumerate(description_list):
                        page.drawString(
                            description_x[k[0]],
                            description_y[k[0]],
                            description_list[k[0]],
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
    generate_caption_pdf(
        "/Users/masataka/Coding/Pythons/Licosha/Display/assets/リコシャ　2022早稲田祭展　写真収集フォーム.xlsx",
        "/Users/masataka/Desktop/Plate",
        "/Users/masataka/Coding/Pythons/Licosha/Display",
    )
