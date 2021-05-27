import PySimpleGUI as sg
import configparser
import os.path
from pathlib import Path
import zipfile
from datetime import datetime
import shutil

SYMBOL_UP =    '▲'
SYMBOL_DOWN =  '▼'


def collapse(layout, key):
    """
    Helper function that creates a Column that can be later made hidden, thus appearing "collapsed"
    :param layout: The layout for the section
    :param key: Key used to make this seciton visible / invisible
    :return: A pinned column that can be placed directly into your layout
    :rtype: sg.pin
    """
    return sg.pin(sg.Column(layout, key=key))


def create_config(path):
    """
    Create a config file
    """
    config = configparser.ConfigParser()
    config.add_section("MarshNum")
    config.set("MarshNum", "Soisp2", "2 соисп")
    config.set("MarshNum", "Soisp3", "3 соисп")
    config.set("MarshNum", "Soisp4", "4 соисп")
    config.set("MarshNum", "Soisp5", "5 соисп")
    config.set("MarshNum", "Soisp6", "6 соисп")
    config.set("MarshNum", "Soisp7", "7 соисп")
    config.set("MarshNum", "Soisp8", "8 соисп")
    config.set("MarshNum", "Soisp9", "9 соисп")
    config.set("MarshNum", "Soisp10", "10 соисп")
    config.set("MarshNum", "Soisp11", "11 соисп")
    config.set("MarshNum", "Soisp12", "12 соисп")
    config.set("MarshNum", "Soisp13", "13 соисп")
    config.set("MarshNum", "Soisp14", "14 соисп")
    config.set("MarshNum", "Soisp15", "15 соисп")
    config.set("MarshNum", "Soisp16", "16 соисп")

    with open(path, "w") as config_file:
        config.write(config_file)


def get_config(path):
    """
    Returns the config object
    """
    if not os.path.exists(path):
        create_config(path)

    config = configparser.ConfigParser()
    config.read(path)
    return config


def get_setting(path, section, setting, msg=0):
    """
    Print out a setting
    """
    config = get_config(path)
    value = config.get(section, setting)
    if msg:
        msg = "{section} {setting} is {value}".format(
            section=section, setting=setting, value=value)

#    window['-STATUS-'].update(msg)
    return value


def update_setting(path, section, setting, value):
    """
    Update a setting
    """
    config = get_config(path)
    config.set(section, setting, value)
    with open(path, "w") as config_file:
        config.write(config_file)

if __name__ == "__main__":
    path = "settings.ini"
    ac2 = get_setting(path, 'MarshNum', 'Soisp2')
    ac3 = get_setting(path, 'MarshNum', 'Soisp3')
    ac4 = get_setting(path, 'MarshNum', 'Soisp4')
    ac5 = get_setting(path, 'MarshNum', 'Soisp5')
    ac6 = get_setting(path, 'MarshNum', 'Soisp6')
    ac7 = get_setting(path, 'MarshNum', 'Soisp7')
    ac8 = get_setting(path, 'MarshNum', 'Soisp8')
    ac9 = get_setting(path, 'MarshNum', 'Soisp9')
    ac10 = get_setting(path, 'MarshNum', 'Soisp10')
    ac11 = get_setting(path, 'MarshNum', 'Soisp11')
    ac12 = get_setting(path, 'MarshNum', 'Soisp12')
    ac13 = get_setting(path, 'MarshNum', 'Soisp13')
    ac14 = get_setting(path, 'MarshNum', 'Soisp14')
    ac15 = get_setting(path, 'MarshNum', 'Soisp15')
    ac16 = get_setting(path, 'MarshNum', 'Soisp16')

sg.theme('Dark Green 7')

section1 = [[sg.Txt(ac5, size=(8, 1)), sg.In('0', size=(8, 1), key='-factS5-'),
             sg.In('0', size=(8, 1), key='-shodS5-')],
            [sg.Txt(ac6, size=(8, 1)), sg.In('0', size=(8, 1), key='-factS6-'),
             sg.In('0', size=(8, 1), key='-shodS6-')],
            [sg.Txt(ac7, size=(8, 1)), sg.In('0', size=(8, 1), key='-factS7-'),
             sg.In('0', size=(8, 1), key='-shodS7-')],
            [sg.Txt(ac8, size=(8, 1)), sg.In('0', size=(8, 1), key='-factS8-'),
             sg.In('0', size=(8, 1), key='-shodS8-')],
            [sg.Txt(ac9, size=(8, 1)), sg.In('0', size=(8, 1), key='-factS9-'),
             sg.In('0', size=(8, 1), key='-shodS9-')],
            [sg.Txt(ac10, size=(8, 1)), sg.In('0', size=(8, 1), key='-factS10-'),
             sg.In('0', size=(8, 1), key='-shodS10-')]]

section2 = [[sg.Txt(ac11, size=(8, 1)), sg.In('0', size=(8, 1), key='-factS11-'),
             sg.In('0', size=(8, 1), key='-shodS11-')],
            [sg.Txt(ac12, size=(8, 1)), sg.In('0', size=(8, 1), key='-factS12-'),
             sg.In('0', size=(8, 1), key='-shodS12-')],
            [sg.Txt(ac13, size=(8, 1)), sg.In('0', size=(8, 1), key='-factS13-'),
             sg.In('0', size=(8, 1), key='-shodS13-')],
            [sg.Txt(ac14, size=(8, 1)), sg.In('0', size=(8, 1), key='-factS14-'),
             sg.In('0', size=(8, 1), key='-shodS14-')],
            [sg.Txt(ac15, size=(8, 1)), sg.In('0', size=(8, 1), key='-factS15-'),
             sg.In('0', size=(8, 1), key='-shodS15-')],
            [sg.Txt(ac16, size=(8, 1)), sg.In('0', size=(8, 1), key='-factS16-'),
             sg.In('0', size=(8, 1), key='-shodS16-')]]


layout = [ [sg.Txt('Введите значение из Отчета:', justification='center')],
           [sg.Txt('МАП (с соисполнением)', relief='sunken', justification='center')],
           [sg.Txt('план', size=(8,1), justification='center'), sg.Txt('факт', size=(8,1), justification='center'), sg.Txt('сходы', size=(8,1), justification='center')],
           [sg.Txt('0', size=(8,1), key='-plan-', relief='sunken'), sg.In('0', size=(8,1), key='-fact-'), sg.In('0', size=(8,1), key='-shod-')],
           [sg.Txt('_'  * 10), sg.Txt('_'  * 10), sg.Txt('_'  * 10)],
           [sg.Txt('соисполнение отдельно', relief='sunken')],
           [sg.Txt('план', size=(8,1), justification='center'), sg.Txt('факт', size=(8,1), justification='center'), sg.Txt('сходы', size=(8,1), justification='center')],
           [sg.Txt('0', size=(8,1), key='-planS-', relief='sunken'), sg.In('0', size=(8,1), key='-factS1-'), sg.In('0', size=(8,1), key='-shodS1-')],
           [sg.Txt(ac2, size=(8,1)), sg.In('0', size=(8,1), key='-factS2-'), sg.In('0', size=(8,1), key='-shodS2-')],
           [sg.Txt(ac3, size=(8,1)), sg.In('0', size=(8,1), key='-factS3-'), sg.In('0', size=(8,1), key='-shodS3-')],
           [sg.Txt(ac4, size=(8,1)), sg.In('0', size=(8,1), key='-factS4-'), sg.In('0', size=(8,1), key='-shodS4-')],

           [sg.Txt('_'  * 10), sg.Txt('_'  * 10), sg.Txt('_'  * 10)],
           [sg.Checkbox('Убрать доп маршруты 1', enable_events=True, key='-OPEN SEC1-CHECKBOX'), sg.Checkbox('Убрать доп маршруты 1', enable_events=True, key='-OPEN SEC2-CHECKBOX')],
            #### Section 1 part ####
            [sg.T(SYMBOL_DOWN, enable_events=True, k='-OPEN SEC1-', text_color='yellow'), sg.T('Доп маршруты 1', enable_events=True, text_color='yellow', k='-OPEN SEC1-TEXT')],
            [collapse(section1, '-SEC1-')],
            #### Section 2 part ####
            [sg.T(SYMBOL_DOWN, enable_events=True, k='-OPEN SEC2-', text_color='purple'),
             sg.T('Доп маршруты 2', enable_events=True, text_color='purple', k='-OPEN SEC2-TEXT')],
            [collapse(section2, '-SEC2-')],

           [sg.Txt('_'  * 10), sg.Txt('_'  * 10), sg.Txt('_'  * 10)],
           [sg.Txt('МАП без соисполнения', relief='sunken')],
           [sg.Txt('план', size=(8,1), justification='center'), sg.Txt('факт', size=(8,1), justification='center'), sg.Txt('сходы', size=(8,1), justification='center')],
           [sg.Txt('0', size=(8,1), key='-planwS-', relief='sunken'), sg.Txt('0', size=(8,1), key='-factwS-', relief='sunken'), sg.Txt('0', size=(8,1), key='-shodwS-', relief='sunken')],
           [sg.Button('Посчитать', bind_return_key=True)]]

window = sg.Window('Калькулятор ЦУП', layout, element_justification='c')
err = "НЕВЕРНО"
opened1, opened2 = True, True




while True:
    event, values = window.read()

    if event == 'Посчитать':
        try:
            fact = int(values['-fact-'])
            shod = int(values['-shod-'])
            shodS1 = int(values['-shodS1-'])
            shodS2 = int(values['-shodS2-'])
            shodS3 = int(values['-shodS3-'])
            shodS4 = int(values['-shodS4-'])
            shodS5 = int(values['-shodS5-'])
            shodS6 = int(values['-shodS6-'])
            shodS7 = int(values['-shodS7-'])
            shodS8 = int(values['-shodS8-'])
            shodS9 = int(values['-shodS9-'])
            shodS10 = int(values['-shodS10-'])
            shodS11 = int(values['-shodS11-'])
            shodS12 = int(values['-shodS12-'])
            shodS13 = int(values['-shodS13-'])
            shodS14 = int(values['-shodS14-'])
            shodS15 = int(values['-shodS15-'])
            shodS16 = int(values['-shodS16-'])

            factS1 = int(values['-factS1-'])
            factS2 = int(values['-factS2-'])
            factS3 = int(values['-factS3-'])
            factS4 = int(values['-factS4-'])
            factS5 = int(values['-factS5-'])
            factS6 = int(values['-factS6-'])
            factS7 = int(values['-factS7-'])
            factS8 = int(values['-factS8-'])
            factS9 = int(values['-factS9-'])
            factS10 = int(values['-factS10-'])
            factS11 = int(values['-factS11-'])
            factS12 = int(values['-factS12-'])
            factS13 = int(values['-factS13-'])
            factS14 = int(values['-factS14-'])
            factS15 = int(values['-factS15-'])
            factS16 = int(values['-factS16-'])

            shodS = shodS1 + shodS2 + shodS3 + shodS4 + shodS5 + shodS6 + shodS7 + shodS8 + shodS9 + shodS10 + shodS11 + shodS12 + shodS13 + shodS14 + shodS15 + shodS16
            factS = factS1 + factS2 + factS3 + factS4 + factS5 + factS6 + factS7 + factS8 + factS9 + factS10 + factS11 + factS12 + factS13 + factS14 + factS15 + factS16
            plan = fact + shod
            planS = factS + shodS
            factwS = fact - factS
            shodwS = shod - shodS
            planwS = factwS + shodwS
        except:
            plan = err
            planS = err
            planwS = err
            factwS = err
            shodwS = err

        window['-plan-'].update(plan)
        window['-planS-'].update(planS)
        window['-planwS-'].update(planwS)
        window['-factwS-'].update(factwS)
        window['-shodwS-'].update(shodwS)

    if event == sg.WIN_CLOSED:
        break

    if event.startswith('-OPEN SEC1-'):
        opened1 = not opened1
        window['-OPEN SEC1-'].update(SYMBOL_DOWN if opened1 else SYMBOL_UP)
        window['-OPEN SEC1-CHECKBOX'].update(not opened1)
        window['-SEC1-'].update(visible=opened1)

    if event.startswith('-OPEN SEC2-'):
        opened2 = not opened2
        window['-OPEN SEC2-'].update(SYMBOL_DOWN if opened2 else SYMBOL_UP)
        window['-OPEN SEC2-CHECKBOX'].update(not opened2)
        window['-SEC2-'].update(visible=opened2)


window.close()