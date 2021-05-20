import PySimpleGUI as sg
import easygui
import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
import win32com.client as win32
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00
#import imwatchingyou
import time

sg.theme('DarkBlue9')

version = __version__ = "1.5.0 Beta"

WIN_W: int = 40
WIN_H: int = 2
Pause_Sleep: float = 0.01

menu_layout: list = [['File', [' Выход ']],
                     ['Help', ['About', 'debugHigh', 'debugLight']]]
# кнопки для конкретно гуи

layout: list = [[sg.Menu(menu_layout)],
                [sg.Text('Добро пожаловать!',  auto_size_text=True, justification='center', font=('Consolas', 12), size=(WIN_W, WIN_H) )],
                [sg.Text('1.для начала cконвертируйте отчет РНИСа и сохраните его.', auto_size_text=True, justification='center', font=('Consolas', 12), size=(WIN_W, WIN_H) )],
                [sg.Text('2. затем выбирайте какие листы создавать. Загрузите файл и затем сохраните.', auto_size_text=True, justification='center', font=('Consolas', 12), size=(WIN_W, WIN_H) )],
                [sg.ProgressBar(1, orientation='h', size=(20, 20), key='progress')],
                [sg.Button(' Егорьевск ', font=('Consolas', 12), size=(WIN_W, WIN_H), border_width=(5), pad=((0, 0), 0)), sg.Button(' Раменское ', font=('Consolas', 12), size=(WIN_W, WIN_H), border_width=(5), pad=((0, 0), 0)), \
                sg.Button(' Шатура ', font=('Consolas', 12), size=(WIN_W, WIN_H), border_width=(5), pad=((0, 0), 0))],
                [sg.Button(' МАП4 ', font=('Consolas', 12), size=(WIN_W, WIN_H), border_width=(5), pad=((0, 0), 0)) ],
                [sg.Button(' Конверт ', font=('Consolas', 12), size=(WIN_W, WIN_H), border_width=(5), pad=((0, 0), 0))],
                [sg.Quit(' Выход ', font=('Consolas', 12), size=(WIN_W, WIN_H), pad=((230, 0), 3))]]


# титульное окно
window = sg.Window('RNISka Reports 1.5.0 Beta', layout, size=(1280, 720))
progress_bar = window.FindElement('progress')
