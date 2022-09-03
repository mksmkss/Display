# progress bar with pysimplegui
import PySimpleGUI as sg
import time

layout = [
    [sg.Text("My Window")],
    [sg.ProgressBar(1000, orientation="h", size=(40, 20), key="progressbar")],
    [sg.Cancel()],
]
window = sg.Window("My Progress Meter", layout)
progress_bar = window["progressbar"]
for i in range(1000):
    # check to see if the cancel button was clicked and exit loop if clicked
    event, values = window.read(timeout=0)
    # if event == "Cancel" or event == sg.WIN_CLOSED:
    #     break
    # update bar with loop value +1 so that bar eventually reaches the maximum
    progress_bar.update_bar(i + 10)
window.close()
