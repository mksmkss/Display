import os
import time
import platform
import webbrowser
import PySimpleGUI as sg

from src.components.Tag import main as Tag
from src.components.QRcode import main as QRcode
from src.components.Description import main as Description

sg.theme("Dark Blue 3")

system = platform.system()

choose_layout = [
    [sg.Text("使用するエクセルファイルを選択してください", font=("Meiryo UI", 15, "bold"))],
    [sg.Text("Excel File:"), sg.Text(), sg.FilesBrowse("Browse", key="excelFilePath")],
    # [sg.Text('ファイル名を決めてください (e.g.2022 早稲田祭)',font=('Meiryo UI',15,'bold'))],
    # [sg.Text("File Name:"), sg.Text(),sg.InputText(key='fileName')],
    [sg.Text("出力先のフォルダーを選択してください", font=("Meiryo UI", 15, "bold"))],
    [
        sg.Text("Output Folder:"),
        sg.Text(),
        sg.FolderBrowse("Browse", key="outputFolder"),
    ],
    [
        sg.Text(
            "使い方はこちら",
            font=("Meiryo UI", 10, "underline"),
            key="howToUse",
            enable_events=True,
        )
    ],
    [sg.Ok("Generate"), sg.Cancel("Cancel")],
]

choose_window = sg.Window("Plate Generator", choose_layout)  # 画面生成

while True:
    event, values = choose_window.read()  # いちいち取得しないといけない
    if event == "Generate" or event == "Cancel":
        break
    if event == "howToUse":
        webbrowser.open("https://github.com/mksmkss/Display")

print(values)

if event == "Generate":
    print("close")
    choose_window.close()
    excelFilePath = values["excelFilePath"]
    outputFolder = values["outputFolder"]

    progress_layout = [
        [sg.Text("進捗状況", font=("Meiryo UI", 15, "bold"))],
        [sg.ProgressBar(100, orientation="h", size=(20, 20), key="progressbar")],
    ]
    progress_window = sg.Window("Plate Generator", progress_layout)

    progress_bar = progress_window["progressbar"]

    mkdir_list = ["QRcode", "Tag PDF", "QRcode PDF", "Description PDF", "Sample PDF"]
    if system == "Darwin":
        for i in mkdir_list:
            os.makedirs(f"{outputFolder}/{i}", exist_ok=True)

    elif system == "Windows":
        for i in mkdir_list:
            os.makedirs(f"{outputFolder}\\{i}", exist_ok=True)

    # Tag
    event, values = progress_window.read(timeout=0)
    progress_bar.update_bar(10)
    Tag.generate_tag_pdf(excelFilePath, outputFolder)

    # QRcode
    event, values = progress_window.read(timeout=0)
    progress_bar.update_bar(40)
    QRcode.generate_qr_pdf(excelFilePath, outputFolder)

    # Description
    event, values = progress_window.read(timeout=0)
    progress_bar.update_bar(70)
    Description.generate_description_pdf(excelFilePath, outputFolder)

    event, values = progress_window.read(timeout=0)
    progress_bar.update_bar(100)
    time.sleep(1)
    progress_window.close()

    fin_layout = [
        [sg.Text("生成が完了しました", font=("Meiryo UI", 15, "bold"))],
        [sg.Text("改善点やエラーがありましたら鈴木柾孝までご連絡ください")],
        [sg.Ok("Done")],
    ]
    print("生成が完了しました")
    fin_window = sg.Window("Plate Generator", fin_layout)  # 画面生成
    while True:
        event, values = fin_window.read()
        if event == "Done":
            break
