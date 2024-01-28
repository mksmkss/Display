"""
Integrationのこのファイルが一番大事
"""


import math
import json
import budoux
import textwrap
import platform

from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Paragraph, Table, TableStyle
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


A4_mm = (210, 297)  # A4用紙のサイズをmmで指定(横、縦)
title_pos = [0, 30]  # タイトルの位置をmmで指定(左、上),ただし，leftは後で計算するため，ここでは0
margin_mm = (10, 0)  # 余白をmmで指定(左右、上下)
title_size = 18  # タイトルのフォントサイズ
default_size = 10  # 本文のフォントサイズ


margin = tuple((to_px(x) for x in margin_mm))  # 余白のサイズpx
margin_title_intro = 20  # タイトルとはじめにの間の余白

column_num = 20  # 一行に並べる0の数
table_size = default_size * 2 + 2  # マークシートのサイズ


system = platform.system()


def textParagraph(c, text, x, y):
    style = getSampleStyleSheet()
    width, height = letter
    p = Paragraph(text, style=style["Normal"])
    p.wrapOn(c, width, height)
    p.drawOn(c, x, y, mm)


def generate_questionnaire_pdf(excel_path, output_path, main_path, exhibition_title):
    """
    左下の座標が(0,0)の座標系で考える．この時，13/20にマークシートを配置する．
    """
    # わざわざsys.argv使っているのは、pyinstallerでexe化した時のエラーを回避するため
    if system == "Darwin":
        font_path = f"{main_path}/assets/ttf/MeiryoUI-03.ttf"
    else:
        font_path = f"{main_path}\\assets\\ttf\\MeiryoUI-03.ttf"
    pdfmetrics.registerFont(cidfonts.UnicodeCIDFont("HeiseiMin-W3"))
    pdfmetrics.registerFont(TTFont("usefont", font_path))

    if system == "Darwin":
        file_name = f"{output_path}/Q're PDF/q're_{exhibition_title}.pdf"  # 保存するファイル名
    else:
        file_name = f"{output_path}\\Q're PDF\\q're_{exhibition_title}.pdf"  # 保存するファイル名

    # A4のキャンバスを作成
    page = canvas.Canvas(file_name, pagesize=A4)

    exhibition_title = f"{exhibition_title}に関するアンケート"
    # stringWidthは文字列の幅を返す関数.単位はpx
    exhi_title_width = stringWidth(exhibition_title, "usefont", title_size)
    title_pos[0] = (A4[0] - exhi_title_width) / 2

    # titleの描画
    page.setFont("usefont", title_size)
    page.drawString(
        title_pos[0],
        to_px(A4_mm[1] - title_pos[1]),
        exhibition_title,
    )

    # はじめにを描画
    intro_words = [
        "本日は早稲田大学リコシャ写真部の写真展にお越しいただきありがとうございます。写真展をご覧くださった方に、アンケートの",
        "ご記入をお願いしております。いただいた意見は、今後の写真展の改善に役立てさせていただきます。よろしくお願いいたします。",
        "なお，本アンケートは機械にて読み取りを行いますので，必ず黒色のボールペンで塗りつぶしてください。",
    ]

    for i, j in enumerate(intro_words):
        page.setFont("usefont", default_size)
        page.drawString(
            to_px(margin[0] - 10),
            to_px(A4_mm[1] - title_pos[1] - margin_title_intro - 8 * i),
            intro_words[i],
        )

    # 現在のカードの位置を記録する変数
    pos_y = A4_mm[1] - title_pos[1] - margin_title_intro - 8 * len(intro_words)

    print(pos_y)
    _plates_list = get_plates_list(excel_path)
    plate_len = len(_plates_list)

    # 説明
    expla_words = ["1.", "気に行った作品の作品番号を塗りつぶしてください。"]
    for i, j in enumerate(expla_words):
        page.setFont("usefont", default_size)
        page.drawString(
            to_px(margin[0] - 10),
            to_px(pos_y - margin_title_intro - 8 * i),
            expla_words[i],
        )

    # 現在のカードの位置を記録する変数
    pos_y -= 8 * len(expla_words) + margin_title_intro

    # 0をcolumn_num個ずつならべるようなリストを作成，行数はlen(_plates_list)//column_num+1
    table_data = [
        [0 for i in range(column_num)] for j in range(plate_len // column_num)
    ]
    # # 一行目に1,2,3,...,column_numを入れる
    # table_data.insert(0, [i for i in range(1, column_num + 1)])
    if plate_len % column_num != 0:
        table_data.append([0 for i in range(plate_len % column_num)])
    for i, j in enumerate(table_data):
        table = Table(
            table_data,
            colWidths=(table_size + 2),
            rowHeights=(table_size + 2),
        )
        table.setStyle(
            TableStyle(
                [
                    ("FONT", (0, 0), (-1, -1), "usefont", table_size),
                    ("INNERGRID", (0, 0), (-1, -1), 0.25, HexColor("#000000")),
                    ("BOX", (0, 0), (-1, -1), 0.25, HexColor("#000000")),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ]
            )
        )
        # wrapOnは，テーブルを描写するために必要なサイズを計算する関数.ややこしいのが，wrapOnの引数はpxで指定する必要があることと，wrapOnは，テーブルの左下の座標を指定する関数であること
        table.wrapOn(
            page,
            (A4[0] - (table_size + 2) * 20) / 2,
            to_px(pos_y)
            - (table_size + 2) * len(table_data)
            - to_px(margin_title_intro / 2),
        )
        # drawOnは，wrapOnで計算したサイズを元に，テーブルを描写する関数. y座標はーを加えると下に移動する．pos_yは，はじめにA4[1]からひいいていることに注意
        table.drawOn(
            page,
            (A4[0] - (table_size + 2) * 20) / 2,
            to_px(pos_y)
            - (table_size + 2) * len(table_data)
            - to_px(margin_title_intro / 2),
        )
        # # マークする部分の描写
        # page.setFont("usefont", default_size)
        # page.drawString(
        #     pos[0] + to_px(5),
        #     pos[1] - to_px(5),
        #     f"0 {plate_text[0]}",
        # )

        # # 以下の処理はカードの位置を設定する処理
        # pos[1] -= card[1]
        # # iがmod(cards_num[1])==cards_num[1]-1の時，つまり，カードの縦の枚数に達した時
        # if i % cards_num[1] == cards_num[1] - 1:
        #     pos[1] = default_pos[1]
        #     pos[0] += card[0]

    page.save()


if __name__ == "__main__":
    generate_questionnaire_pdf(
        "/Users/masataka/Coding/Pythons/Licosha/Display/assets/excel/リコシャ　2023早稲田祭展　写真収集フォーム .xlsx",
        "/Users/masataka/Desktop/plate",
        "/Users/masataka/Coding/Pythons/Licosha/Display",
        "2023早稲田祭展",
    )
