import os
import sys
import json
import shutil
import tkinter
import platform
from tkinter import filedialog
import customtkinter
from PIL import Image

from src.components.Integeration import main as Integration
from src.components.Manupulate_PDF import main as ManupulatePDF

from src.components.Tag import main as Tag
from src.components.QRcode import main as QRcode
from src.components.Description import main as Description


customtkinter.set_appearance_mode("System")  # Modes: system (default), light, dark
customtkinter.set_default_color_theme(
    "blue"
)  # Themes: blue (default), dark-blue, green

system = platform.system()

app = customtkinter.CTk()  # create CTk window like you do with the Tk window
main_width = 640
main_height = 320
app.geometry(f"{main_width}x{main_height}")
app.title("Plate Generator")

main_path = os.path.dirname(sys.argv[0])

# 設定ファイルを読み込む．過去のログ（pathなど）を保存しておくためのもの
with open(f"{main_path}/settings.json", "r", encoding="utf-8") as _settings:
    global excel_path, outputFolder_path, year, exhibition_title
    dic = json.load(_settings)
    excel_path = dic["excel_path"]
    outputFolder_path = dic["outputFolder_path"]
    year = dic["year"]
    exhibition_title = dic["exhibition_title"]


icon = customtkinter.CTkImage(
    light_image=Image.open(f"{main_path}/assets/img/icons8-folder-48-light.png"),
    dark_image=Image.open(f"{main_path}/assets/img/icons8-folder-48-dark.png"),
    size=(22, 22),
)


class PathEntry(customtkinter.CTkFrame):
    index = 0

    # CustomTkinterのフレームを継承したクラス
    def __init__(self, master, width, height, path_text, index, path_disp):
        # super()は親クラスのメソッドを呼び出す（使えるようにする）
        super().__init__(master, width, height)
        self.index = index
        self.path_disp = customtkinter.CTkEntry(
            master=self.master,
            width=width - 65,
            height=height - 5,
            font=("Arial", 11),
        )
        self.path_disp.insert(0, path_text)
        # 挿入されてからreadonlyにしないと文字が表示されない
        self.path_disp.configure(state="readonly")
        self.path_disp.place(relx=0.02, rely=0.5, anchor=tkinter.W)
        self.button = customtkinter.CTkButton(
            master=self.master,
            image=icon,
            command=self.open_folder,
            text="",
            width=24,
            height=24,
        )
        self.button.place(relx=0.99, rely=0.5, anchor=tkinter.E)

    def open_folder(self):
        if self.index <= 0:
            if system == "Darwin":
                new_path = filedialog.askopenfilename(
                    initialdir="/",
                    filetypes=(("Excel", ".xlsx .xls"),),
                )
            else:
                new_path = filedialog.askopenfilename(
                    initialdir="/",
                    filetypes=(
                        ("Excel", ".xlsx .xls"),
                        ("ExcelMacro .xlsm"),
                    ),
                )
                print(new_path)
                new_path = new_path.replace("/", "\\")
        else:
            # フォルダーを選択する
            new_path = filedialog.askdirectory()
            if system == "Windows":
                new_path = new_path.replace("/", "\\")
        # まずは表示を変更，そうしないと値が変更できず表示されない
        self.path_disp.configure(state="normal")
        self.path_disp.delete(0, tkinter.END)
        self.path_disp.insert(0, new_path)
        self.path_disp.configure(state="readonly")
        # path_disp.append(new_path)
        dic[list(dic.keys())[self.index]] = new_path


period = 0.24
width = main_width * 0.96
height = 40
label = ["使用するExcelファイル", "出力先のフォルダー", "PDF名"]
path_disp = []

for i in enumerate(label):
    index = i[0]
    label_disp = customtkinter.CTkLabel(
        master=app, text=i[1], font=("Arial", 14, "bold")
    )
    label_disp.place(relx=0.05, rely=0.05 + period * index, anchor=tkinter.W)
    frame = customtkinter.CTkFrame(
        master=app,
        width=width,
        height=height,
    )
    frame.grid(row=0, column=0, padx=20)
    frame.place(
        relx=0.5,
        rely=period * index + 0.15,
        anchor=tkinter.CENTER,
    )
    if index < 2:
        path = list(dic.values())[index]
        PathEntry(frame, width, height, path, index, path_disp)
    else:
        # validateで入力された値を検証する
        year_disp = customtkinter.CTkEntry(
            master=frame,
            width=width * 0.2,
            height=height - 5,
            font=("Arial", 11),
            placeholder_text="Year",
        )
        year_disp.place(relx=0.02, rely=0.5, anchor=tkinter.W)
        title_disp = customtkinter.CTkEntry(
            master=frame,
            width=width * 0.75,
            height=height - 5,
            font=("Arial", 11),
            placeholder_text="Exhibition Title",
        )
        title_disp.place(relx=0.98, rely=0.5, anchor=tkinter.E)


def Process():
    toplevel = customtkinter.CTkToplevel()
    toplevel.geometry("300x200")
    toplevel.transient(app)  # これでToplevelウィンドウを親ウィンドウに関連付ける
    isFilled = True
    dic["year"] = year_disp.get()
    dic["exhibition_title"] = title_disp.get()
    print(dic)
    with open(f"{main_path}/settings.json", "w", encoding="utf-8") as f:
        json.dump(dic, f, ensure_ascii=False, indent=4)
    for i in range(len(dic)):
        if list(dic.values())[i] == "":
            isFilled = False
    ProcessLookupError = customtkinter.CTkLabel(
        master=toplevel, text="", font=("Arial", 16, "bold")
    )
    if isFilled == True:
        toplevel.title("Processing...")
        ProcessLookupError.configure(text="処理中です...")
        ProcessLookupError.place(relx=0.5, rely=0.4, anchor=tkinter.CENTER)
        toplevel.update()
        mkdir_list = [
            "QRcode",
            "Tag PDF",
            "QRcode PDF",
            "Description PDF",
            "Caption PDF",
        ]
        excel_path = dic["excel_path"]
        outputFolder_path = dic["outputFolder_path"]
        if system == "Darwin":
            for i in mkdir_list:
                # はじめに，出力先のフォルダーの中身を削除する
                shutil.rmtree(f"{outputFolder_path}/{i}", ignore_errors=True)
                os.makedirs(f"{outputFolder_path}/{i}")
        elif system == "Windows":
            for i in mkdir_list:
                # はじめに，出力先のフォルダーの中身を削除する
                shutil.rmtree(f"{outputFolder_path}\\{i}", ignore_errors=True)
                os.makedirs(f"{outputFolder_path}\\{i}")

        # メインの関数はIntegration
        try:
            Integration.generate_caption_pdf(excel_path, outputFolder_path, main_path)
        except Exception as e:
            print(e)
            ProcessLookupError.configure(text="エラーが発生しました")
            ProcessLookupError.place(relx=0.5, rely=0.3, anchor=tkinter.CENTER)
            Detail = customtkinter.CTkLabel(
                master=toplevel, text="", font=("Arial", 12)
            )

            # 改行を入れる
            __e = str(e).split(" ")
            _e = []
            # /の前後で分割するが，/は残す
            for j in enumerate(__e):
                if "/" in j[1]:
                    _text_list = j[1].split("/")
                    text_list = [text + "/" for text in _text_list]
                    _e.extend(text_list)
                else:
                    _e.append(f"{j[1]} ")
            print(_e)
            e = []
            long = 0
            line = ""
            for k in _e:
                if (long + len(k)) <= 40:
                    line += k
                    long += len(k)
                else:
                    e.append(line)
                    line = k
                    long = len(k)
            e.append(line)
            e = "\n".join(e)
            Detail.configure(text=f"{e}")
            Detail.place(relx=0.5, rely=0.5, anchor=tkinter.CENTER)
            toplevel.update()
            if str(e) == "'numpy.int64' object has no attribute 'split'":
                Detail.configure(
                    text="Excelの列名がずれています\n「お名前」,「[写真の詳細] タイトル」,「[写真の詳細] 説明」,「ペンネーム」となっていることを確認してください"
                )
                Detail.place(relx=0.5, rely=0.5, anchor=tkinter.CENTER)
                toplevel.update()
            fin = customtkinter.CTkButton(
                master=toplevel, text="終了する", command=app.destroy
            )
            fin.place(relx=0.5, rely=0.8, anchor=tkinter.CENTER)
            # エラーが発生したら終了する
            return 0
        # 頑張ったので昔の関数もついでに出力しておく
        Tag.generate_tag_pdf(excel_path, outputFolder_path, main_path)
        QRcode.generate_qr_pdf(excel_path, outputFolder_path, main_path)
        Description.generate_description_pdf(excel_path, outputFolder_path, main_path)

        ProcessLookupError.configure(text="処理が完了しました")
        ProcessLookupError.place(relx=0.5, rely=0.4, anchor=tkinter.CENTER)
        fin = customtkinter.CTkButton(master=toplevel, text="終了する", command=app.destroy)
        fin.place(relx=0.5, rely=0.8, anchor=tkinter.CENTER)
        year = year_disp.get()
        title = title_disp.get()
        file_name = f"{year}_{title}.pdf"
        ManupulatePDF.merge_pdfs(
            f"{outputFolder_path}/Caption PDF",
            file_name,
        )
    else:
        toplevel.title("Error")
        ProcessLookupError.configure(text="すべての項目を入力してください")
        ProcessLookupError.place(relx=0.5, rely=0.5, anchor=tkinter.CENTER)


button = customtkinter.CTkButton(
    master=app,
    text="Process",
    command=Process,
)
button.place(relx=0.5, rely=0.85, anchor=tkinter.CENTER)


app.mainloop()
