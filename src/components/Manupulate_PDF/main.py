import os
import glob
import PyPDF2
import platform
import webbrowser
import subprocess

merger = PyPDF2.PdfMerger()


def merge_pdfs(outputFolder, file_name):
    pdf_list = sorted(glob.glob(f"{outputFolder}/each PDF/*.pdf"))
    for i in pdf_list:
        merger.append(i)
        print(i)
        # os.remove(i)
    merger.write(f"{outputFolder}/{file_name}")
    merger.close()
    # 生成したPDFを開く
    if platform.system() == "Darwin":
        subprocess.run(["open", f"{outputFolder}"])
    elif platform.system() == "Windows":
        os.startfile(f"{outputFolder}")
    # webbrowser.open("https://acrobat.adobe.com/link/home/?x_api_client_id=adobe_com")


if __name__ == "__main__":
    merge_pdfs("/Users/masataka/Desktop/Plate/Caption PDF", "2022")
