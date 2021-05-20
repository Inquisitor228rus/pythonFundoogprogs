import subprocess
import configparser
import os.path
from pathlib import Path
import zipfile
from datetime import datetime
import PySimpleGUI as sg
import shutil

def find_ms():
    if os.path.isfile("C://Program Files (x86)//Microsoft Office//Office16//MSACCESS.EXE"):
        window['-MSAC-'].update("C:\Program Files (x86)\Microsoft Office\Office16\MSACCESS.EXE")
        window['-STATUS-'].update("MS Access находится по адресу MS16")
    elif os.path.isfile("C://Program Files (x86)//Microsoft Office//Office12//MSACCESS.EXE"):
        window['-MSAC-'].update("C:\Program Files (x86)\Microsoft Office\Office12\MSACCESS.EXE")
        window['-STATUS-'].update("MS Access находится по адресу MS12")
    else:
        window['-MSAC-'].update("Не найден MS Access.")
        window['-STATUS-'].update("Не найден MS Access!")


'''
CONFIG CREATE/USE
'''


def create_config(path):
    """
    Create a config file
    """
    config = configparser.ConfigParser()
    config.add_section("ACCESS")
    config.set("ACCESS", "ACPATH", r"C:\Program Files (x86)\Microsoft Office\Office16\MSACCESS.EXE")

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


def runner(ac, pathP):
    proc = subprocess.run([ac, pathP],
                          stdout=subprocess.PIPE,
                          stderr=subprocess.STDOUT,
                          universal_newlines=True)

    if proc.returncode != 0:
        window['-STATUS-'].update("В этой программе произошла ошибка: " + ac + "\n")
    else:
        window['-STATUS-'].update('done')


'''
PACK THIS CODE
'''
DATE_FORMAT = '%d%m%y'


def date_str():
    """returns the today string year, month, day. ващето день месяц год епты"""
    return '{}'.format(datetime.now().strftime(DATE_FORMAT))


def zip_name(path_back):
    """returns the zip filename as string"""
    cur_dir = Path(path_back).resolve()
    parent_dir = cur_dir.parents[0]
    zip_filename = '{}/{}_{}.zip'.format(parent_dir, cur_dir.name, date_str())
    p_zip = Path(zip_filename)
    n = 1
    while p_zip.exists():
        zip_filename = ('{}/{}_{}_{}.zip'.format(parent_dir, cur_dir.name,
                                                 date_str(), n))
        p_zip = Path(zip_filename)
        n += 1
    return zip_filename


def all_files(path_back):
    """iterator returns all files and folders from path as absolute path string
    """
    for child in Path(path_back).iterdir():
        yield str(child)
        if child.is_dir():
            for grand_child in all_files(str(child)):
                yield str(Path(grand_child))

    mta_main = 0
    mta_rs = 0
    mta_xp = 0


def zip_dir(path_back):
    """generate a zip"""
    zip_filename = zip_name(path_back)
    zip_file = zipfile.ZipFile(zip_filename, 'w')
    print('create:', zip_filename)
    for file in all_files(path_back):
        print('adding... ', file)
        zip_file.write(file)
    zip_file.close()
    print('end!')


'''
COPY DEF
'''


def copy_files():



    try:
        mta_xp = r'C:\rs\rs_mta_XP.accdb'
        mta_main = r'C:\rs\rs-main-4_39.accdb'
        mta_classic = r'C:\rs\rs_39.accdb'
        target_mta_xp = r'C:\rs\backup\rs_mta_XP.accdb'
        target_mta_main = r'C:\rs\backup\rs-main-4_39.accdb'
        target_mta_classic = r'C:\rs\backup\rs_39.accdb'

        shutil.copyfile(mta_xp, target_mta_xp)
        shutil.copyfile(mta_main, target_mta_main)
        shutil.copyfile(mta_classic, target_mta_classic)

    except Exception() as e:
        print(e)



'''
MENU CANVAS
'''
layout = [[sg.Text('НЕ ЗАКОНЧЕННАЯ ВЕРСИ\nЕЩЕ НЕ ЗАВЕРШЕННЫЙ ВИД')],

          [sg.Text('Где живет запускатор Access\nдля РС', relief='solid')],
          [sg.Text(' ', key='-MSAC-', size=(70, 1), auto_size_text=True)],
          [sg.Text('Управление:')],
          [sg.Button('КОПИРОВАНИЕ')], [sg.Button('АРХИВИРОВАНИЕ')],
          [sg.Button('СТАРТ', bind_return_key=True, button_color=('white', 'green'), size=(12,2)),
          sg.Button('Выход', button_color=('blue', 'red'), size=(12,2))],
          [sg.Text('Состояние:'), sg.Text(' ', key='-STATUS-', size=(70, 1))]]

window = sg.Window('Открывашка РС', layout, finalize=True, element_justification='c')

if __name__ == "__main__":
    pathAC = 0
    path = "settings.ini"
    find_ms()
    ac = get_setting(path, 'ACCESS', 'ACPATH')
    pathP = r"C:\rs\rs-main-4_39.accdb"
    #update_setting(path, "Settings", "font_size", "12")
    window['-STATUS-'].update(shutil.disk_usage(r'C:\rs\backup'))


while True:
    event, values = window.read()
    #    print(event, values)
    if event == 'СТАРТ':
        runner(ac, pathP)
#        zip_dir(r'C:\rs\backup\1')
    if event == 'КОПИРОВАНИЕ':
        print('ну да. сейчас ')
        copy_files()
    if event == 'АРХИВИРОВАНИЕ':
        print('ну да. сейча ')
        zip_dir(r'C:\rs\backup')
    if event == sg.WIN_CLOSED or event == 'Выход':
        break

window.close()
