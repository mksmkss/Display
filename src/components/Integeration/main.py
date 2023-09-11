import math
import json
import textwrap
import platform

from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Paragraph
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics, cidfonts
from reportlab.lib.styles import getSampleStyleSheet

from reportlab.pdfbase.pdfmetrics import stringWidth

if __name__ == "__main__":
    from functions import *
else:
    from .functions import *

cards_num = (2, 5)  # カードをA4に何枚配置するか（横の枚数, 縦の枚数）
A4_mm = (210, 297)  # A4用紙のサイズをmmで指定(横、縦)
card_mm = (105, 59)  # 1枚のラベルのサイズをmmで指定(横、縦)
margin_mm = (0, 0)  # 余白をmmで指定(左右、上下)

card = tuple((to_px(x) for x in card_mm))  # カードのサイズpx
margin = tuple((to_px(x) for x in margin_mm))  # 余白のサイズpx

# 各パラメータの調整．ちなみにQRコードの大きさもrect_heightで自動調整される
title_size = 16
penname_size = 13
description_size = 12

rect_height = 14

system = platform.system()


def textParagraph(c, text, x, y):
    style = getSampleStyleSheet()
    width, height = letter
    p = Paragraph(text, style=style["Normal"])
    p.wrapOn(c, width, height)
    p.drawOn(c, x, y, mm)


def generate_caption_pdf(excel_path, output_path, main_path):
    # わざわざsys.argv使っているのは、pyinstallerでexe化した時のエラーを回避するため
    if system == "Darwin":
        font_path = f"{main_path}/assets/MeiryoUI-03.ttf"
    else:
        font_path = f"{main_path}\\assets\\MeiryoUI-03.ttf"
    pdfmetrics.registerFont(cidfonts.UnicodeCIDFont("HeiseiMin-W3"))
    pdfmetrics.registerFont(TTFont("usefont", font_path))

    _plates_list = get_plates_list(excel_path)
    _description_list = get_description_list(excel_path)
    _ids_dict = get_ids_dict(excel_path)
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


            # 線の太さを指定
            page.setLineWidth(1)
            # 描画の初期地点
            pos = [margin[0], margin[1]]

            for y in range(cards_num[1]):
                for x in range(cards_num[0]):
                    # 図形（長方形、直線）とテキストの描画
                    page.rect(pos[0], pos[1], card[0], card[1], fill=False)

                    # 下の黒いところの描画
                    page.setFillColor(HexColor("#2c2c2e"))
                    page.rect(pos[0], pos[1], card[0], to_px(rect_height), fill=1)

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
                    description_list = textwrap.wrap(description, 17)

                    # それぞれのtitleの位置を取得，複数行にまたがる場合も考慮
                    title_width_list = []
                    title_x = []
                    title_y = []

                    # description_width_list = []
                    description_x = []
                    description_y = []

                    # titleは複数行を考慮しない
                    page.setFont("HeiseiMin-W3", title_size)
                    title_width_list.append(
                        round(page.stringWidth(title, "HeiseiMin-W3", title_size))
                    )
                    # 座標の基準は文字の左下
                    title_x.append(pos[0] + card[0] * 0.08)
                    title_y.append(pos[1] + card[1] * 0.8)

                    for k in enumerate(description_list):
                        # description
                        page.setFont("HeiseiMin-W3", description_size)
                        description_x.append(pos[0] + card[0] * 0.15)
                        description_y.append(
                            pos[1]
                            + (card[1] / 2)
                            + (description_size / 2) * (len(description_list) - 2)
                            - description_size * k[0]
                        )

                    # pennameの位置を取得
                    penname_width = round(
                        page.stringWidth(penname, "HeiseiMin-W3", penname_size)
                    )

                    # pennameの位置,titke_x,yを流用
                    title_x.append(pos[0] + card[0] - card[0] * 0.08 - penname_width)
                    title_y.append(
                        pos[1] + to_px(rect_height) / 2 - penname_size / 2 + 2
                    )

                    # titleの描画
                    page.setFillColor(HexColor("#2c2c2e"))
                    page.setFont("HeiseiMin-W3", title_size)
                    page.drawString(title_x[0], title_y[0], title)

                    # pennameの描画
                    p_penname = f"""<font name="HeiseiMin-W3" color="rgb(237, 237, 235)" size={penname_size}>{penname}</font>"""
                    textParagraph(page, p_penname, title_x[1], title_y[1])

                    # descriptionの描画
                    page.setFillColor(HexColor("#2c2c2e"))
                    page.setFont("HeiseiMin-W3", description_size)
                    for k in enumerate(description_list):
                        page.drawString(
                            description_x[k[0]],
                            description_y[k[0]],
                            description_list[k[0]],
                        )

                    penname_to_sns_dict = _ids_dict
                    sns_list = penname_to_sns_dict[penname]
                    print(f"{sns_list}, {penname}")
                    for l in enumerate(sns_list):
                            id = l[1][0]
                            sns = l[1][1]
                            if sns == "instagram":
                                generate_qr(
                                    f"https://www.instagram.com/{id}?utm_source=qr",
                                    sns,
                                    f"{sns}_{id}.png",
                                    output_path,
                                )
                            else:
                                generate_qr(
                                    f"https://x.com/{id}",
                                    sns,
                                    f"{sns}_{id}.png",
                                    output_path,
                                )
                            if system == "Darwin":
                                image = Image.open(
                                    f"{output_path}/QRcode/{sns}_{id}.png"
                                )
                            else:
                                image = Image.open(
                                    f"{output_path}\\QRcode\\{sns}_{id}.png"
                                )

                            page.drawInlineImage(
                                image,
                                pos[0] + l[0] * to_px(rect_height + 1) + to_px(9),
                                pos[1] + to_px(1),
                                width=to_px(rect_height - 2),
                                height=to_px(rect_height - 2),
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
        "/Users/masataka/Desktop/2023 夏季展.xlsx",
        "/Users/masataka/Desktop/Plate2",
        "/Users/masataka/Coding/Pythons/Licosha/Display",
    )
