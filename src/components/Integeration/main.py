"""
Integrationのこのファイルが一番大事
"""

import os
import math
import budoux
import platform
import textwrap

from PIL import Image
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

notaking_width = 12
data_matrix_width = 18

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
        font_path = f"{main_path}/assets/ttf/MeiryoUI-03.ttf"
    else:
        font_path = f"{main_path}\\assets\\ttf\\MeiryoUI-03.ttf"
    pdfmetrics.registerFont(cidfonts.UnicodeCIDFont("HeiseiMin-W3"))
    pdfmetrics.registerFont(TTFont("usefont", font_path))

    _plates_list = get_plates_list(excel_path)
    _description_list = get_description_list(excel_path)
    _ids_dict = get_ids_dict(excel_path)
    _permission_dict = get_permission_dict(excel_path)
    _uuid_list = get_uuid_list(excel_path)
    page_len = math.ceil(len(_plates_list) / cards_num[0] * cards_num[1])

    print(f"page_len: {page_len}")

    isEnd = False
    for i in range(page_len):
        # iがページ数，jが各ページにおけるカード番号
        j = 0

        # A4のpdfに必要なものを描画する, このwhile文はページごとに回る
        while j < cards_num[0] * cards_num[1]:
            # A4のpdfを作成する
            if system == "Darwin":
                file_name = f"{output_path}/Caption PDF/caption_{i}.pdf"
            else:
                try:
                    os.makedirs(f"{output_path}\\Caption PDF\\each PDF", exist_ok=True)
                except:
                    print("Error", f"{output_path}\\Caption PDF\\each PDF")
                    pass
                file_name = f"{output_path}\\Caption PDF\\each PDF\\caption_{i}.pdf"
            page = canvas.Canvas(file_name, pagesize=A4)

            # 線の太さを指定
            page.setLineWidth(1)
            # 描画の初期地点
            pos = [margin[0], margin[1]]
            # このfor文はカードごとに回る
            for y in range(cards_num[1]):
                for x in range(cards_num[0]):
                    """
                    pos[0]：紙の左端からの距離，pos[1]：紙の下端からの距離
                    card[0]：カードの横幅，card[1]：カードの縦幅 この大きさは上で定義してある
                    rect_height：下の黒いところの高さ，これは上で定義してある．QRコードの大きさもこれで自動調整される
                    """
                    # 図形（長方形、直線）とテキストの描画
                    page.rect(pos[0], pos[1], card[0], card[1], fill=False)

                    # カードの下の黒いところの描画
                    page.setFillColor(HexColor("#2c2c2e"))
                    page.rect(pos[0], pos[1], card[0], to_px(rect_height), fill=1)

                    # 最終ページはデータが途中で消えバグるので，回避．つまり，最終ページがカードが10枚未満になるとき用の変数がisEnd
                    if j + 10 * i >= len(_plates_list):
                        page.save()
                        isEnd = True
                        print("First loop is done!")
                        break

                    # j+10*i番目のカードの情報を取得
                    title = _plates_list[j + 10 * i][0]
                    penname = _plates_list[j + 10 * i][1]
                    description = _description_list[j + 10 * i]
                    each_uuid = _uuid_list[j + 10 * i]

                    print(
                        f"title: {title}, penname: {penname}, description: {description}, each_uuid: {each_uuid}"
                    )

                    # googleの日本語parserを使う.これにより自然な改行ができるようになった
                    parser = budoux.load_default_japanese_parser()
                    _description = parser.parse(description)
                    description_list = []
                    long = 0
                    line = ""
                    max_len = 18
                    for k in _description:
                        if (long + len(k)) <= max_len:
                            line += k
                            long += len(k)
                        else:
                            description_list.append(line)
                            line = k
                            long = len(k)
                    description_list.append(line)

                    # description_list[0]==''の時は，英語の可能性が高いので，日本語のparserを使わない.
                    if description_list[0] == "":
                        # wrap(テキストデータ,文字数)
                        description_list = textwrap.wrap(description, 40)

                    # それぞれのtitleの位置を取得
                    title_width_list = []
                    title_x = []
                    title_y = []

                    # それぞれのdescriptionの位置を取得,複数行になることを考慮
                    description_x = []
                    description_y = []

                    # titleは複数行を考慮しない
                    page.setFont("HeiseiMin-W3", title_size)
                    # stringWidthは文字列の幅を取得する関数.単位はpx
                    title_width_list.append(
                        round(page.stringWidth(title, "HeiseiMin-W3", title_size))
                    )
                    # 座標の基準は文字の左下
                    title_x.append(pos[0] + card[0] * 0.08)
                    title_y.append(pos[1] + card[1] * 0.8)

                    print(f"title_x: {title_x}, title_y: {title_y}")

                    # descriptionは複数行を考慮
                    for k in enumerate(description_list):
                        # description
                        page.setFont("HeiseiMin-W3", description_size)
                        description_x.append(pos[0] + card[0] * 0.15)
                        description_y.append(
                            pos[1]
                            + (card[1] / 2)
                            + (description_size / 2) * (len(description_list) - 2)
                            - description_size * k[0]
                            - 5
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
                    print(f"sns_list: {sns_list}")
                    # snsのQRコードの描画,xやinstagramのQRコードの描画
                    for l in enumerate(sns_list):
                        id = l[1][0]
                        sns = l[1][1]
                        if isinstance(id, float):
                            continue
                        # まず，QRコードを生成
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
                            image = Image.open(f"{output_path}/QRcode/{sns}_{id}.png")
                        else:
                            image = Image.open(f"{output_path}\\QRcode\\{sns}_{id}.png")

                        # 次に，QRコードを描画
                        page.drawInlineImage(
                            image,
                            pos[0] + l[0] * to_px(rect_height + 1) + to_px(9),
                            pos[1] + to_px(1),
                            width=to_px(rect_height - 2),
                            height=to_px(rect_height - 2),
                        )

                    print(f"each_uuid: {each_uuid}")
                    # escapeを使って，文字列をエスケープ
                    escaped_penname = penname.replace("/", "-")
                    escaped_title = title.replace("/", "-")

                    # DataMatrixの生成
                    generate_data_matrix(
                        each_uuid,
                        f"{output_path}/Data Matrix/{escaped_penname}_{escaped_title}.png",
                    )

                    # DataMatrixの描画
                    if system == "Darwin":
                        image = Image.open(
                            f"{output_path}/Data Matrix/{escaped_penname}_{escaped_title}.png"
                        )
                    else:
                        image = Image.open(
                            f"{output_path}\\Data Matrix\\{escaped_penname}_{escaped_title}.png"
                        )
                    page.drawInlineImage(
                        image,
                        pos[0] + card[0] - card[0] * 0.08 - to_px(data_matrix_width),
                        pos[1] + card[1] * 0.95 - to_px(data_matrix_width),
                        width=to_px(data_matrix_width),
                        height=to_px(data_matrix_width),
                    )

                    # No Takingの描画
                    if _permission_dict[penname] == "No":

                        if system == "Darwin":
                            image = Image.open(
                                f"{main_path}/assets/img/icons8-nocamera-100.png"
                            )
                        else:
                            image = Image.open(
                                f"{main_path}\\assets\\img\\icons8-nocamera-100.png"
                            )
                        page.drawInlineImage(
                            image,
                            pos[0]
                            + card[0]
                            - card[0] * 0.08
                            - to_px(notaking_width)
                            - 4
                            - to_px(data_matrix_width),
                            pos[1] + card[1] * 0.95 - to_px(notaking_width),
                            width=to_px(notaking_width),
                            height=to_px(notaking_width),
                        )

                    j += 1
                    # 次のカードの描画位置を指定，x軸方向にcard[0]分移動
                    pos[0] += card[0]

                if isEnd:
                    print("Second loop is done!")
                    break
                # 次のカードの描画位置を指定，y軸方向にcard[1]分移動, x軸方向は初期位置に戻す
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
        "/Users/masataka/Coding/Pythons/Licosha/Display/assets/excel/リコシャ　2023早稲田祭展　写真収集フォーム .xlsx",
        "/Users/masataka/Desktop/plate",
        "/Users/masataka/Coding/Pythons/Licosha/Display",
    )
