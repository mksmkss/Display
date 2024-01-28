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


A4_mm = (210, 297)  # A4用紙のサイズをmmで指定(横、縦)
title_pos = [0, 20]  # タイトルの位置をmmで指定(左、上),ただし，leftは後で計算するため，ここでは0
margin_mm = (10, 0)  # 余白をmmで指定(左右、上下)
title_size = 18  # タイトルのフォントサイズ
default_size = 12  # 本文のフォントサイズ


margin = tuple((to_px(x) for x in margin_mm))  # 余白のサイズpx
margin_title_worklist = 12  # タイトルとはじめにの間の余白


system = platform.system()


def textParagraph(c, text, x, y):
    style = getSampleStyleSheet()
    width, height = letter
    p = Paragraph(text, style=style["Normal"])
    p.wrapOn(c, width, height)
    p.drawOn(c, x, y, mm)


def generate_worklist_pdf(excel_path, output_path, main_path, _exhibition_title):
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

    # タイトルリストを取得
    _plates_list = get_plates_list(excel_path)
    plate_len = len(_plates_list)
    print(_plates_list)

    # 人数の取得. setで重複を削除している
    people_num = len(set([x[1] for x in _plates_list]))

    page_len = 1

    # _plates_listの長さに応じて，縦横の枚数を変更する
    cards_num = (2, (plate_len + people_num) // 2 + 1)
    if plate_len + people_num > 50:
        cards_num = (3, (plate_len + people_num) // 3 + 1)
    if plate_len + people_num > 100:
        cards_num = (4, (plate_len + people_num) // 4 + 1)
    if plate_len + people_num > 150:
        cards_num = (6, (plate_len + people_num) // 6 + 1)
        page_len = 2

    print(cards_num, plate_len, people_num)

    # 2ページに分ける場合，1ページに書かれる列は3列になる
    card = (
        (A4[0] - margin[0] * 2) / cards_num[0] * page_len,
        (A4[1] * 17 / 20) / cards_num[1],
    )  # カードのサイズをpxで指定(横、縦)

    for i in range(page_len):
        if system == "Darwin":
            file_name = f"{output_path}/PieceList PDF/piecelist_{_exhibition_title}_{i}.pdf"  # 保存するファイル名
        else:
            file_name = f"{output_path}\\PieceList PDF\\piecelist_{_exhibition_title}_{i}.pdf"  # 保存するファイル名

        # A4のキャンバスを作成
        page = canvas.Canvas(file_name, pagesize=A4)

        exhibition_title = f"{_exhibition_title} 作品リスト"
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

        # 現在のカードの位置を記録する変数
        pos_y = A4_mm[1] - title_pos[1] - margin_title_worklist

        default_pos = [margin[0], to_px(pos_y)]  # マークシートを書き始める左上の座標をpxで指定(x,y)
        pos = default_pos.copy()

        penname = ""
        pre_penname = ""

        # 記述済みのカードの枚数
        drew_num = 0
        # 現在の列数
        column_num = 0
        for j in range(plate_len):
            # 数字の幅
            num_width = stringWidth(f"({j+1})", "HeiseiMin-W3", 15)

            plate_text = _plates_list[j][0]
            penname = _plates_list[j][1]
            # 文字数の最大値は，カードの横幅を文字の幅で割ったもの
            max_char_num = math.floor(
                card[0] / stringWidth("あ", "usefont", default_size)
            )

            # pennameを比較して，違うなら描写する
            if penname != pre_penname:
                page.setFont("usefont", default_size)
                page.drawString(
                    pos[0] - to_px(2),
                    pos[1] - to_px(5),
                    f"{penname}",
                )
                pre_penname = penname
                pos[1] -= card[1] - to_px(2)
                drew_num += 1

                # jがmod(cards_num[1])==cards_num[1]-1の時，つまり，カードの縦の枚数に達した時
                if drew_num % cards_num[1] == cards_num[1] - 1:
                    pos[1] = default_pos[1]
                    pos[0] += card[0]
                    drew_num = 0
                    column_num += 1
                    continue

            # 文字数がmax_char_numを超えたら，省略する
            if len(plate_text) > max_char_num:
                plate_text = plate_text[:max_char_num] + "..."
            # マークする部分の描写
            page.setFont("usefont", default_size)
            page.drawString(
                pos[0] - num_width / 2 + to_px(5),
                pos[1] - to_px(5),
                f"{j+1} {plate_text}",
            )
            drew_num += 1

            # 以下の処理はカードの位置を設定する処理
            pos[1] -= card[1]
            # iがmod(cards_num[1])==cards_num[1]-1の時，つまり，カードの縦の枚数に達した時
            if drew_num % cards_num[1] == cards_num[1] - 1:
                pos[1] = default_pos[1]
                pos[0] += card[0]
                drew_num = 0
                column_num += 1
            # ページの右端に達した時
            if column_num == cards_num[0] or j == plate_len - 1:
                pos[0] = default_pos[0]
                pos[1] = default_pos[1]
                column_num = 0
                break

        page.save()


if __name__ == "__main__":
    generate_worklist_pdf(
        "/Users/masataka/Coding/Pythons/Licosha/Display/assets/excel/リコシャ　2023早稲田祭展　写真収集フォーム .xlsx",
        "/Users/masataka/Desktop/plate",
        "/Users/masataka/Coding/Pythons/Licosha/Display",
        "2023早稲田祭展",
    )
