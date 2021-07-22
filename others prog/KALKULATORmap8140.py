import PySimpleGUI as sg
import configparser
import os.path

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
    config = configparser.ConfigParser(allow_no_value=True)
    config.add_section("MarshNum")
    config.add_section("Settings")
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
    config.set("MarshNum", "; справа от равно название маршрута")
    config.set("MarshNum", "; не меняйте ничего что слева от равно")
    config.set("MarshNum", "справа от равно можно написать и текст и цифры.")
    config.set("Settings", "; справа от знака равно укажите")
    config.set("Settings", "; сколько строк соисполнителей будет")
    config.set("Settings", "SumofLines", "7")
    

    with open(path, "w") as config_file:
        config.write(config_file)


def get_config(path):
    """
    Returns the config object
    """
    if not os.path.exists(path):
        create_config(path)

    config = configparser.ConfigParser(allow_no_value=True)
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
    numList = int(get_setting(path, 'Settings', 'SumofLines'))

sg.theme('Dark Green 7')


baseSoisp3 = [[sg.Txt(ac3, size=(8,1)), sg.In('0', size=(8,1), key='-factS3-'),
        sg.In('0', size=(8,1), key='-shodS3-')]]
baseSoisp4 = [[sg.Txt(ac4, size=(8,1)), sg.In('0', size=(8,1), key='-factS4-'),
        sg.In('0', size=(8,1), key='-shodS4-')]]
baseSoisp5 = [[sg.Txt(ac5, size=(8,1)), sg.In('0', size=(8,1), key='-factS5-'),
        sg.In('0', size=(8,1), key='-shodS5-')]]
baseSoisp6 = [[sg.Txt(ac6, size=(8,1)), sg.In('0', size=(8,1), key='-factS6-'),
        sg.In('0', size=(8,1), key='-shodS6-')]]
baseSoisp7 = [[sg.Txt(ac7, size=(8,1)), sg.In('0', size=(8,1), key='-factS7-'),
        sg.In('0', size=(8,1), key='-shodS7-')]]
baseSoisp8 = [[sg.Txt(ac8, size=(8,1)), sg.In('0', size=(8,1), key='-factS8-'),
        sg.In('0', size=(8,1), key='-shodS8-')]]
baseSoisp9 = [[sg.Txt(ac9, size=(8,1)), sg.In('0', size=(8,1), key='-factS9-'),
        sg.In('0', size=(8,1), key='-shodS9-')]]
baseSoisp10 = [[sg.Txt(ac10, size=(8,1)), sg.In('0', size=(8,1), key='-factS10-'),
        sg.In('0', size=(8,1), key='-shodS10-')]]
baseSoisp11 = [[sg.Txt(ac11, size=(8,1)), sg.In('0', size=(8,1), key='-factS11-'),
        sg.In('0', size=(8,1), key='-shodS11-')]]
baseSoisp12 = [[sg.Txt(ac12, size=(8,1)), sg.In('0', size=(8,1), key='-factS12-'),
        sg.In('0', size=(8,1), key='-shodS12-')]]
baseSoisp13 = [[sg.Txt(ac13, size=(8,1)), sg.In('0', size=(8,1), key='-factS13-'),
        sg.In('0', size=(8,1), key='-shodS13-')]]
baseSoisp14 = [[sg.Txt(ac14, size=(8,1)), sg.In('0', size=(8,1), key='-factS14-'),
        sg.In('0', size=(8,1), key='-shodS14-')]]
baseSoisp15 = [[sg.Txt(ac15, size=(8,1)), sg.In('0', size=(8,1), key='-factS15-'),
        sg.In('0', size=(8,1), key='-shodS15-')]]
baseSoisp16 = [[sg.Txt(ac16, size=(8,1)), sg.In('0', size=(8,1), key='-factS16-'),
        sg.In('0', size=(8,1), key='-shodS16-')]]


menuLayout = [[sg.Txt('Введите значение из Отчета:', justification='center')],
           [sg.Txt('МАП (с соисполнением)', relief='sunken', justification='center')],
           [sg.Txt('план', size=(8,1), justification='center'), sg.Txt('факт', size=(8,1), justification='center'), sg.Txt('сходы', size=(8,1), justification='center')],
           [sg.Txt('0', size=(8,1), key='-plan-', relief='sunken'), sg.In('0', size=(8,1), key='-fact-'), sg.In('0', size=(8,1), key='-shod-')],
           [sg.Txt('_'  * 10), sg.Txt('_'  * 10), sg.Txt('_'  * 10)],
           [sg.Txt('соисполнение отдельно', relief='sunken')],
           [sg.Txt('план', size=(8,1), justification='center'), sg.Txt('факт', size=(8,1), justification='center'), sg.Txt('сходы', size=(8,1), justification='center')],
           [sg.Txt('0', size=(8,1), key='-planS-', relief='sunken'),
            sg.In('0', size=(8,1), key='-factS1-'),
            sg.In('0', size=(8,1), key='-shodS1-')]]

soispLayout = [[sg.Txt(ac2, size=(8,1)), sg.In('0', size=(8,1), key='-factS2-'),
            sg.In('0', size=(8,1), key='-shodS2-')]]

if numList == 3:
    
    soispLayout += baseSoisp3


if numList == 4:
    
    soispLayout += baseSoisp3 + baseSoisp4


if numList == 5:
    
    soispLayout += baseSoisp3 + baseSoisp4 + baseSoisp5


if numList == 6:
    
    soispLayout += baseSoisp3 + baseSoisp4 + baseSoisp5 + baseSoisp6


if numList == 7:
    
    soispLayout += baseSoisp3 + baseSoisp4 + baseSoisp5 + baseSoisp6 + baseSoisp7


if numList == 8:
    
    soispLayout += baseSoisp3 + baseSoisp4 + baseSoisp5 + baseSoisp6 + baseSoisp7 + baseSoisp8



if numList == 9:
    
    soispLayout += baseSoisp3 + baseSoisp4 + baseSoisp5 + baseSoisp6 + baseSoisp7 + baseSoisp8 + baseSoisp9



if numList == 10:
    
    soispLayout += baseSoisp3 + baseSoisp4 + baseSoisp5 + baseSoisp6 + baseSoisp7 + baseSoisp8 + baseSoisp9 + baseSoisp10



if numList == 11:
    
    soispLayout += baseSoisp3 + baseSoisp4 + baseSoisp5 + baseSoisp6 + baseSoisp7 + baseSoisp8 + baseSoisp9 + baseSoisp10 + baseSoisp11



if numList == 12:
    
    soispLayout += baseSoisp3 + baseSoisp4 + baseSoisp5 + baseSoisp6 + baseSoisp7 + baseSoisp8 + baseSoisp9 + baseSoisp10 + baseSoisp11 + baseSoisp12



if numList == 13:
    
    soispLayout += baseSoisp3 + baseSoisp4 + baseSoisp5 + baseSoisp6 + baseSoisp7 + baseSoisp8 + baseSoisp9 + baseSoisp10 + baseSoisp11 + baseSoisp12 + baseSoisp13


if numList == 14:
    
    soispLayout += baseSoisp3 + baseSoisp4 + baseSoisp5 + baseSoisp6 + baseSoisp7 + baseSoisp8 + baseSoisp9 + baseSoisp10 + baseSoisp11 + baseSoisp12 + baseSoisp13 + baseSoisp14


if numList == 15:
    
    soispLayout += baseSoisp3 + baseSoisp4 + baseSoisp5 + baseSoisp6 + baseSoisp7 + baseSoisp8 + baseSoisp9 + baseSoisp10 + baseSoisp11 + baseSoisp12 + baseSoisp13 + baseSoisp14 + baseSoisp15

    
if numList >= 16:

    soispLayout += baseSoisp3 + baseSoisp4 + baseSoisp5 + baseSoisp6 + baseSoisp7 + baseSoisp8 + baseSoisp9 + baseSoisp10 + baseSoisp11 + baseSoisp12 + baseSoisp13 + baseSoisp14 + baseSoisp15 + baseSoisp16


mapLayout = [[sg.Txt('_'  * 10), sg.Txt('_'  * 10), sg.Txt('_'  * 10)],
           [sg.Txt('МАП без соисполнения', relief='sunken')],
           [sg.Txt('план', size=(8,1), justification='center'), sg.Txt('факт', size=(8,1), justification='center'), sg.Txt('сходы', size=(8,1), justification='center')],
           [sg.Txt('0', size=(8,1), key='-planwS-', relief='sunken'), sg.Txt('0', size=(8,1), key='-factwS-', relief='sunken'), sg.Txt('0', size=(8,1), key='-shodwS-', relief='sunken')],
           [sg.Button('Посчитать', bind_return_key=True)]]

layout = menuLayout + soispLayout + mapLayout
window = sg.Window('Калькулятор ЦУП', layout, element_justification='c')
err = "НЕВЕРНО"
opened1, opened2 = True, True

shodS1,shodS2,shodS3,shodS4,shodS5,shodS6,shodS7,shodS8,shodS9,shodS10,shodS11,shodS12,shodS13,shodS14,shodS15,shodS16 = 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0
factS1,factS2,factS3,factS4,factS5,factS6,factS7,factS8,factS9,factS10,factS11,factS12,factS13,factS14,factS15,factS16 = 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0

while True:
    event, values = window.read()
    #print(event)
    if event == 'Посчитать':
        try:
            fact = int(values['-fact-'])
            shod = int(values['-shod-'])
            shodS1 = int(values['-shodS1-'])
            shodS2 = int(values['-shodS2-'])
            
            factS1 = int(values['-factS1-'])
            factS2 = int(values['-factS2-'])
            if numList == 3:
                
    
                shodS3 = int(values['-shodS3-'])

                factS3 = int(values['-factS3-'])

            if numList == 4:
    
                shodS3 = int(values['-shodS3-'])
                shodS3 = int(values['-shodS4-'])

                factS3 = int(values['-factS3-'])
                factS4 = int(values['-factS4-'])

            if numList == 5:
    
                shodS3 = int(values['-shodS3-'])
                shodS3 = int(values['-shodS4-'])
                shodS3 = int(values['-shodS5-'])

                factS3 = int(values['-factS3-'])
                factS4 = int(values['-factS4-'])
                factS5 = int(values['-factS5-'])

            if numList == 6:
    
                shodS3 = int(values['-shodS3-'])
                shodS3 = int(values['-shodS4-'])
                shodS3 = int(values['-shodS5-'])
                shodS3 = int(values['-shodS6-'])

                factS3 = int(values['-factS3-'])
                factS4 = int(values['-factS4-'])
                factS5 = int(values['-factS5-'])
                factS6 = int(values['-factS6-'])

            if numList == 7:
    
                shodS3 = int(values['-shodS3-'])
                shodS3 = int(values['-shodS4-'])
                shodS3 = int(values['-shodS5-'])
                shodS3 = int(values['-shodS6-'])
                shodS3 = int(values['-shodS7-'])

                factS3 = int(values['-factS3-'])
                factS4 = int(values['-factS4-'])
                factS5 = int(values['-factS5-'])
                factS6 = int(values['-factS6-'])
                factS7 = int(values['-factS7-'])

            if numList == 8:
    
                shodS3 = int(values['-shodS3-'])
                shodS3 = int(values['-shodS4-'])
                shodS3 = int(values['-shodS5-'])
                shodS3 = int(values['-shodS6-'])
                shodS3 = int(values['-shodS7-'])
                shodS3 = int(values['-shodS8-'])

                factS3 = int(values['-factS3-'])
                factS4 = int(values['-factS4-'])
                factS5 = int(values['-factS5-'])
                factS6 = int(values['-factS6-'])
                factS7 = int(values['-factS7-'])
                factS8 = int(values['-factS8-'])


            if numList == 9:
    
                shodS3 = int(values['-shodS3-'])
                shodS3 = int(values['-shodS4-'])
                shodS3 = int(values['-shodS5-'])
                shodS3 = int(values['-shodS6-'])
                shodS3 = int(values['-shodS7-'])
                shodS3 = int(values['-shodS8-'])
                shodS3 = int(values['-shodS9-'])

                factS3 = int(values['-factS3-'])
                factS4 = int(values['-factS4-'])
                factS5 = int(values['-factS5-'])
                factS6 = int(values['-factS6-'])
                factS7 = int(values['-factS7-'])
                factS8 = int(values['-factS8-'])
                factS9 = int(values['-factS9-'])


            if numList == 10:
    
                shodS3 = int(values['-shodS3-'])
                shodS3 = int(values['-shodS4-'])
                shodS3 = int(values['-shodS5-'])
                shodS3 = int(values['-shodS6-'])
                shodS3 = int(values['-shodS7-'])
                shodS3 = int(values['-shodS8-'])
                shodS3 = int(values['-shodS9-'])
                shodS3 = int(values['-shodS10-'])

                factS3 = int(values['-factS3-'])
                factS4 = int(values['-factS4-'])
                factS5 = int(values['-factS5-'])
                factS6 = int(values['-factS6-'])
                factS7 = int(values['-factS7-'])
                factS8 = int(values['-factS8-'])
                factS9 = int(values['-factS9-'])
                factS10 = int(values['-factS10-'])


            if numList == 11:
    
                shodS3 = int(values['-shodS3-'])
                shodS3 = int(values['-shodS4-'])
                shodS3 = int(values['-shodS5-'])
                shodS3 = int(values['-shodS6-'])
                shodS3 = int(values['-shodS7-'])
                shodS3 = int(values['-shodS8-'])
                shodS3 = int(values['-shodS9-'])
                shodS3 = int(values['-shodS10-'])
                shodS3 = int(values['-shodS11-'])

                factS3 = int(values['-factS3-'])
                factS4 = int(values['-factS4-'])
                factS5 = int(values['-factS5-'])
                factS6 = int(values['-factS6-'])
                factS7 = int(values['-factS7-'])
                factS8 = int(values['-factS8-'])
                factS9 = int(values['-factS9-'])
                factS10 = int(values['-factS10-'])
                factS11 = int(values['-factS11-'])


            if numList == 12:
    
                shodS3 = int(values['-shodS3-'])
                shodS3 = int(values['-shodS4-'])
                shodS3 = int(values['-shodS5-'])
                shodS3 = int(values['-shodS6-'])
                shodS3 = int(values['-shodS7-'])
                shodS3 = int(values['-shodS8-'])
                shodS3 = int(values['-shodS9-'])
                shodS3 = int(values['-shodS10-'])
                shodS3 = int(values['-shodS11-'])
                shodS3 = int(values['-shodS12-'])

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


            if numList == 13:
    
                shodS3 = int(values['-shodS3-'])
                shodS3 = int(values['-shodS4-'])
                shodS3 = int(values['-shodS5-'])
                shodS3 = int(values['-shodS6-'])
                shodS3 = int(values['-shodS7-'])
                shodS3 = int(values['-shodS8-'])
                shodS3 = int(values['-shodS9-'])
                shodS3 = int(values['-shodS10-'])
                shodS3 = int(values['-shodS11-'])
                shodS3 = int(values['-shodS12-'])
                shodS3 = int(values['-shodS13-'])

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


            if numList == 14:
                
                shodS3 = int(values['-shodS3-'])
                shodS3 = int(values['-shodS4-'])
                shodS3 = int(values['-shodS5-'])
                shodS3 = int(values['-shodS6-'])
                shodS3 = int(values['-shodS7-'])
                shodS3 = int(values['-shodS8-'])
                shodS3 = int(values['-shodS9-'])
                shodS3 = int(values['-shodS10-'])
                shodS3 = int(values['-shodS11-'])
                shodS3 = int(values['-shodS12-'])
                shodS3 = int(values['-shodS13-'])
                shodS3 = int(values['-shodS14-'])

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


            if numList == 15:
                
                shodS3 = int(values['-shodS3-'])
                shodS3 = int(values['-shodS4-'])
                shodS3 = int(values['-shodS5-'])
                shodS3 = int(values['-shodS6-'])
                shodS3 = int(values['-shodS7-'])
                shodS3 = int(values['-shodS8-'])
                shodS3 = int(values['-shodS9-'])
                shodS3 = int(values['-shodS10-'])
                shodS3 = int(values['-shodS11-'])
                shodS3 = int(values['-shodS12-'])
                shodS3 = int(values['-shodS13-'])
                shodS3 = int(values['-shodS14-'])
                shodS3 = int(values['-shodS15-'])

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

    
            if numList >= 16:
                shodS3 = int(values['-shodS3-'])
                shodS3 = int(values['-shodS4-'])
                shodS3 = int(values['-shodS5-'])
                shodS3 = int(values['-shodS6-'])
                shodS3 = int(values['-shodS7-'])
                shodS3 = int(values['-shodS8-'])
                shodS3 = int(values['-shodS9-'])
                shodS3 = int(values['-shodS10-'])
                shodS3 = int(values['-shodS11-'])
                shodS3 = int(values['-shodS12-'])
                shodS3 = int(values['-shodS13-'])
                shodS3 = int(values['-shodS14-'])
                shodS3 = int(values['-shodS15-'])
                shodS3 = int(values['-shodS16-'])

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



window.close()
