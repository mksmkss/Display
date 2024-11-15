import math
import platform
from PIL import Image

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

if __name__ == "__main__":
    from functions import *
else:
    from .functions import *

cards_num = (4, 5)  # カードをA4に何枚配置するか（横の枚数, 縦の枚数）
A4_mm = (210, 297)  # A4用紙のサイズをmmで指定(横、縦)
card_mm = (50, 50)  # 1枚のラベルのサイズをmmで指定(横、縦)
margin_mm = (0, 0)  # 余白をmmで指定(左右、上下)

to_px = A4[0] / A4_mm[0]  # mmをpxに変換
card = tuple((x * to_px for x in card_mm))  # カードのサイズpx
margin = tuple((x * to_px for x in margin_mm))  # 余白のサイズpx

system = platform.system()


def generate_qr_pdf(excel_path, output_path, main_path):
    # わざわざsys.argv使っているのは、pyinstallerでexe化した時のエラーを回避するため
    if system == "Darwin":
        font_path = f"{main_path}/assets/ttf/MeiryoUI-03.ttf"
    else:
        font_path = f"{main_path}\\assets\\ttf\\MeiryoUI-03.ttf"
    pdfmetrics.registerFont(TTFont("Meiryo UI", font_path))
    _instagram_data = get_id_list(excel_path, "instagram")
    _twitter_data = get_id_list(excel_path, "twitter")

    _id_list = _instagram_data + _twitter_data
    page_len = math.ceil(len(_id_list) / cards_num[0] * cards_num[1])

    isEnd = False
    for i in range(page_len):
        # iがページ数，jが各ページにおけるカード番号
        j = 0

        # A4のpdfに必要なものを描画する
        while j < cards_num[0] * cards_num[1]:
            # A4のpdfを作成する
            if system == "Darwin":
                file_name = f"{output_path}/QRcode PDF/qr_{i}.pdf"
            else:
                file_name = f"{output_path}\\QRcode PDF\\qr_{i}.pdf"
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
                    if j + 10 * i >= len(_id_list):
                        page.save()
                        isEnd = True
                        print("First loop is done!")
                        break

                    id = _id_list[j + 10 * i][0]
                    sns = _id_list[j + 10 * i][1]
                    # 画像の挿入
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
                    # to broaden image
                    page.drawInlineImage(
                        image,
                        pos[0] + card[0] / 2 - 55,
                        pos[1] + card[1] / 6,
                        width=110,
                        height=110,
                    )

                    fontsize = 10
                    # pennameの位置を取得
                    id_width = round(page.stringWidth(id, "Meiryo UI", fontsize))

                    # ペンネーム
                    page.setFont("Meiryo UI", fontsize)
                    # #16537B=22, 83, 123
                    page.setFillColorRGB(0.08, 0.32, 0.48)
                    page.drawString(
                        pos[0] + card[0] / 2 - id_width / 2, pos[1] + 10, id
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
    generate_qr_pdf(
        "/Users/masataka/Desktop/写真展フォーム　テンプレート.xlsx", "/Users/masataka/Desktop/Data"
    )
