import os
import sys
import platform
import PySimpleGUI as sg

from src.components.A4Paper import main as A4Paper
from src.components.QRcode import main as QRcode

sg.theme('Dark Blue 3')

system=platform.system()

choose_layout=[
    [sg.Text('使用するエクセルファイルを選択してください',font=('Meiryo UI',15,'bold'))], 
    [sg.Text("Excel File:"), sg.Text(),sg.FilesBrowse('Browse',key='excelFilePath')],
    # [sg.Text('ファイル名を決めてください (e.g.2022 早稲田祭)',font=('Meiryo UI',15,'bold'))],
    # [sg.Text("File Name:"), sg.Text(),sg.InputText(key='fileName')],
    [sg.Text('出力先のフォルダーを選択してください',font=('Meiryo UI',15,'bold'))],
    [sg.Text("Output Folder:"), sg.Text(),sg.FolderBrowse('Browse',key='outputFolder')],
    [sg.Ok('Generate'), sg.Cancel('Cancel')],
]

choose_window= sg.Window('Plate Generator', choose_layout) #画面生成

while True:
    event,values= choose_window.read() #いちいち取得しないといけない
    if event == 'Generate' or event == 'Cancel':
        break

print(values)

if event == 'Generate':
    print("close")
    choose_window.close()
    excelFilePath=values['excelFilePath']
    outputFolder=values['outputFolder']
    if system =='Darwin':
        os.mkdir(outputFolder+'/'+'QRCodes')
        os.mkdir(outputFolder+'/'+'Tags PDF')
        os.mkdir(outputFolder+'/'+'QRcodes PDF')
    elif system =='Windows':
        os.mkdir(outputFolder+'\\'+'QRCodes')
        os.mkdir(outputFolder+'\\'+'Tags PDF')
        os.mkdir(outputFolder+'\\'+'QRcodes PDF')

    QRcode.generate_qr_pdf(excelFilePath,outputFolder)
    A4Paper.generate_tag_pdf(excelFilePath,outputFolder)
    
    fin_layout=[
        [sg.Text('生成が完了しました',font=('Meiryo UI',15,'bold'))],
        [sg.Text('改善点やエラーがありましたら鈴木柾孝までご連絡ください')],
        [sg.Ok('Done')]
    ]
    print("生成が完了しました")
    fin_window= sg.Window('Plate Generator', fin_layout) #画面生成
    while True:
        event,values= fin_window.read()
        if event == 'Done':
            break
    