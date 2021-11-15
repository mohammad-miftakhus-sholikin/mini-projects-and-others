## Import some packages
import PySimpleGUI as sg
import pandas as pd


## Add some color to the window
sg.theme('BlueMono')


## Input excel file (note: this file must same in name)
EXCEL_FILE = 'input-form-for-meta-data.xlsx'
df = pd.read_excel(EXCEL_FILE)


## Adding form input and output
# input and output layout
layout = [
    # form header
    [sg.Text('Form to input meta data:')],
    # adding no
    [sg.Text('No', size=(15,1)), sg.InputText(key='No')],
    # adding Study
    [sg.Text('Study', size=(15,1)), sg.InputText(key='Study')],
    # adding Author
    [sg.Text('Author', size=(15,1)), sg.InputText(key='Author')],
    # adding Data1
    [sg.Text('Data1', size=(15,1)), sg.InputText(key='Data1')],
    # submit and exit button
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

## Define title and layout
window = sg.Window('Form to input meta data', layout)

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