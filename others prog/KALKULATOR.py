import PySimpleGUI as sg

sg.theme('Dark Green 7')

layout = [ [sg.Txt('Введите значение из Отчета:', justification='center')],
           [sg.Txt('МАП (с соисполнением)', relief='sunken', justification='center')],
           [sg.Txt('план', size=(8,1), justification='center'), sg.Txt('факт', size=(8,1), justification='center'), sg.Txt('сходы', size=(8,1), justification='center')],
           [sg.Txt('0', size=(8,1), key='-plan-', relief='sunken'), sg.In('0', size=(8,1), key='-fact-'), sg.In('0', size=(8,1), key='-shod-')],
           [sg.Txt('_'  * 10), sg.Txt('_'  * 10), sg.Txt('_'  * 10)],
           [sg.Txt('соисполнение отдельно', relief='sunken')],
           [sg.Txt('план', size=(8,1), justification='center'), sg.Txt('факт', size=(8,1), justification='center'), sg.Txt('сходы', size=(8,1), justification='center')],
           [sg.Txt('0', size=(8,1), key='-planS-', relief='sunken'), sg.In('0', size=(8,1), key='-factS1-'), sg.In('0', size=(8,1), key='-shodS1-')],
           [sg.Txt(size=(8,1)), sg.In('0', size=(8,1), key='-factS2-'), sg.In('0', size=(8,1), key='-shodS2-')],
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
            factS1 = int(values['-factS1-'])
            factS2 = int(values['-factS2-'])
            shodS = shodS1 + shodS2
            factS = factS1 + factS2
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