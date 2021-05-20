import numpy as np
import sys
        
def Calc():
    totals = np.array(Shtuka, dtype='int32')  # don't use Sum because sum is a reserved keyword and it confusing
    resultA = np.array([], dtype='int32')

    a = np.random.random((Marsh,1))  # create random numbers
    a = a/np.sum(a, axis=0) * totals  # force them to sum to totals

    a = np.round(a)  # transform them into integers
    remainings = totals - np.sum(a, axis=0)  # check if there are corrections to be done
    for j, r in enumerate(remainings):  # implement the correction
        step = 1 if r > 0 else -1
        while r != 0:
            i = np.random.randint(44)
            if a[i,j] + step >= 0:
                a[i, j] += step
                r -= step
        if event == 'Принять':
            np.savetxt(sys.stdout, a, fmt="%i")
            window['-lolT-'].update(' ')
        if event == 'Перевернуть':
            window['-lolT-'].update('   КНОПКА СОДЕРЖИТ ОШИБКИ')
            np.flip(np.savetxt(sys.stdout, a, fmt="%i"))
#            print("есть пробитие")
numA = int(44)
numB = int()

import PySimpleGUI as sg

layout = [[sg.Text('Здеся цифрывренции вылезают.')],
          [sg.Output(size=(25, 45), key='_output_')],
          [sg.Text('Вписываем пассажиров и остановки.\n44 обычно для 39го')],
          [sg.Text('Количество остановок:'),
           sg.Input(numA, key='-IN-', enable_events=True, size=(10,1))],
          [sg.Text('Сумма пассажиров:    '),
           sg.Input(numB, key='-IN1-', enable_events=True, size=(10,1))],
          [sg.Text('Нажимаем "принять" \nи собираем циферки.')]
          , [sg.Text('                                                       ', text_color='red', key='-lolT-')],
          [sg.Button('Принять', bind_return_key=True), sg.Button('Перевернуть', button_color=('blue', 'red')), sg.Button('Выход')]]

window = sg.Window('38 и 39 считалка', layout)

while True:
    event, values = window.read()
#    print(event, values)
    if event == 'Принять':
        try:
            Marsh = int(values['-IN-'])
        except:
            Marsh = 0
            sg.popup_error("\nОшибка в первом числе!\n")
            break
        try:
            n = int(1)
            Shtuka = list(map(int,values['-IN1-'].strip().split()))[:n]
        except:
            sg.popup_error("\nОшибка во втором числе!\n")
            break
        window.FindElement('_output_').Update('')
        Calc()

    if event == 'Перевернуть':
        window.FindElement('_output_').Update('')
#        window.FindElement('lolT').Update('')
        Calc()
        
        
    if event == sg.WIN_CLOSED or event == 'Выход':
        break
    if event == '-IN-' and values['-IN-'] and values['-IN-'][-1] not in ('0123456789.'):
        window['-IN-'].update(values['-IN-'][:-1])
    if event == '-IN1-' and values['-IN1-'] and values['-IN1-'][-1] not in ('0123456789.'):
        window['-IN1-'].update(values['-I1N-'][:-1])
    if len(values['-IN-']) > 3:
        window.Element('-IN-').Update(values['-IN-'][:-1])
    if len(values['-IN1-']) > 3:
        window.Element('-IN1-').Update(values['-IN1-'][:-1])
window.close()
