## Import some packages
import PySimpleGUI as sg
import pandas as pd


## Add some color to the window
sg.theme('BlueMono')


## Input excel file (note: this file must same in name)
EXCEL_FILE = 'student-exam-score.xlsx'
df = pd.read_excel(EXCEL_FILE)


## Adding form input and output

# example code of input and output layout
# layout = [
    # [sg.Text('Please fill out the following fields:')],
    # [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    # [sg.Text('City', size=(15,1)), sg.InputText(key='City')],
    # [sg.Text('Favorite Colour', size=(15,1)), sg.Combo(['Green', 'Blue', 'Red'], key='Favorite Colour')],
    # [sg.Text('I speak', size=(15,1)),
                            # sg.Checkbox('German', key='German'),
                            # sg.Checkbox('Spanish', key='Spanish'),
                            # sg.Checkbox('English', key='English')],
    # [sg.Text('No. of Children', size=(15,1)), sg.Spin([i for i in range(0,16)],
                                                       # initial_value=0, key='Children')],
    # [sg.Submit(), sg.Button('Clear'), sg.Exit()]
# ]

# input and output layout
layout = [
    # form header
    [sg.Text('Form to enter exam scores:')],
    # adding no
    [sg.Text('No', size=(15,1)), sg.InputText(key='No')],
    # adding Student_number
    [sg.Text('Student_number', size=(15,1)), sg.InputText(key='Student_number')],
    # adding Name
    [sg.Text('Name', size=(15,1)), sg.InputText(key='Name')],
    # adding Score
    [sg.Text('Score', size=(15,1)), sg.InputText(key='Score')],
    # submit and exit button
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

## Define title and layout
window = sg.Window('Form to enter exam scores', layout)

## If has empty values
def clear_input():
    for key in values:
        window[key]('')
    return None

## Define functional button
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == 'Clear':
        clear_input()
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data saved!')
        clear_input()
window.close()