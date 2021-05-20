# img_viewer.py

import PySimpleGUI as sg
import os.path
import openpyxl

# First the window layout in 2 columns

file_list_column = [
    [
        sg.Text("Выбор файла:"),
        sg.In(size=(25, 1), enable_events=True, key="-FOLDER-"),
        sg.FolderBrowse(button_text='ПОИСК'),
    ],
    [
        sg.Listbox(
            values=[], enable_events=True, size=(80, 20), key="-FILE LIST-"
        )
    ],
]

# For now will only show the name of the file that was chosen
image_viewer_column = [
    [sg.Text("кнопочки всякие")],
    [sg.Text(size=(40, 1), key="-TOUT-")],
    [sg.Output(size=(63, 13), key="-IMAGE-")],
]

# ----- Full layout -----
layout = [
    [
        sg.Column(file_list_column),
        sg.VSeperator(),
        sg.Column(image_viewer_column),
    ]
]

window = sg.Window("макет", layout)

# Run the Event Loop
while True:
    event, values = window.read()
    if event == "Exit" or event == sg.WIN_CLOSED:
        break
    # Folder name was filled in, make a list of files in the folder
    if event == "-FOLDER-":
        folder = values["-FOLDER-"]
        try:
            # Get list of files in folder
            file_list = os.listdir(folder)
        except:
            file_list = []

        fnames = [
            f
            for f in file_list
            if os.path.isfile(os.path.join(folder, f))
            and f.lower().endswith(".xlsx")
            #and f.lower().startswith("итоговый отчет о работе выходов по филиалу")

        ]
        window["-FILE LIST-"].update(fnames)
    elif event == "-FILE LIST-":  # A file was chosen from the listbox
        try:
            filename = os.path.join(
                values["-FOLDER-"], values["-FILE LIST-"][0]
            )
            window["-TOUT-"].update(filename)


            book = openpyxl.load_workbook(filename=filename)

            sheet = book.active

            a1 = sheet['A1']
            a2 = sheet['A2']
            a3 = sheet['A3']

            #a4 = sheet.cell(row=3, column=1)

            print(a1.value)
            print(a2.value)
            print(a3.value)
            #print(a4.value)
            window["-IMAGE-"].update(filename=filename)

        except:
            pass

window.close()