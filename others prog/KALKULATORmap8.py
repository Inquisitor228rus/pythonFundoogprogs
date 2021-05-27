import PySimpleGUI as sg
import configparser
import os.path
from pathlib import Path
import zipfile
from datetime import datetime
import shutil

def create_config(path):
    """
    Create a config file
    """
    config = configparser.ConfigParser()
    config.add_section("MarshNum")
    config.set("MarshNum", "Soisp2", "300")
    config.set("MarshNum", "Soisp3", "666")
    config.set("MarshNum", "Soisp4", "888")
    config.set("MarshNum", "Soisp5", "14")
    config.set("MarshNum", "Soisp6", "9999")
    config.set("MarshNum", "Soisp7", "2220")

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

sg.theme('Dark Green 7')

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
           [sg.Txt(ac5, size=(8,1)), sg.In('0', size=(8,1), key='-factS5-'), sg.In('0', size=(8,1), key='-shodS5-')],
           [sg.Txt(ac6, size=(8,1)), sg.In('0', size=(8,1), key='-factS6-'), sg.In('0', size=(8,1), key='-shodS6-')],
           [sg.Txt(ac7, size=(8,1)), sg.In('0', size=(8,1), key='-factS7-'), sg.In('0', size=(8,1), key='-shodS7-')],
           [sg.Txt('_'  * 10), sg.Txt('_'  * 10), sg.Txt('_'  * 10)],
           [sg.Txt('МАП без соисполнения', relief='sunken')],
           [sg.Txt('план', size=(8,1), justification='center'), sg.Txt('факт', size=(8,1), justification='center'), sg.Txt('сходы', size=(8,1), justification='center')],
           [sg.Txt('0', size=(8,1), key='-planwS-', relief='sunken'), sg.Txt('0', size=(8,1), key='-factwS-', relief='sunken'), sg.Txt('0', size=(8,1), key='-shodwS-', relief='sunken')],
           [sg.Button('Посчитать', bind_return_key=True)]]

window = sg.Window('Калькулятор ЦУП', layout, element_justification='c')
err = "НЕВЕРНО"





while True:
    event, values = window.read()

    if event != sg.WIN_CLOSED:
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
            factS1 = int(values['-factS1-'])
            factS2 = int(values['-factS2-'])
            factS3 = int(values['-factS3-'])
            factS4 = int(values['-factS4-'])
            factS5 = int(values['-factS5-'])
            factS6 = int(values['-factS6-'])
            factS7 = int(values['-factS7-'])
            shodS = shodS1 + shodS2 + shodS3 + shodS4 + shodS5 + shodS6 + shodS7
            factS = factS1 + factS2 + factS3 + factS4 + factS5 + factS6 + factS7
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
    else:
        break
