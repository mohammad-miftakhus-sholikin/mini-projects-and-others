## Import some packages
import PySimpleGUI as sg
import pandas as pd


## Add some color to the window
sg.theme('DefaultNoMoreNagging')


## Input excel file (note: this file must same in name)
EXCEL_FILE = 'student-exam-score.xlsx'
df = pd.read_excel(EXCEL_FILE)


## Adding form input and output
# input and output layout
layout = [
    # form header
    [sg.Text('form to enter exam scores:')],
    # adding no
    [sg.Text('no', size=(15,1)), sg.InputText(key='no')],
    # adding student_id
    [sg.Text('student_id', size=(15,1)), sg.InputText(key='student_id')],
    # adding student_name
    [sg.Text('student_name', size=(15,1)), sg.InputText(key='student_name')],
    # adding score
    [sg.Text('score', size=(15,1)), sg.InputText(key='score')],
    # submit and exit button
    [sg.Submit(button_text='submit'), sg.Button('clear'), sg.Exit(button_text='exit')]
]

## Define title and layout
window = sg.Window('form to enter exam scores', layout)

## Delete all values
def clear_input():
    for key in values:
        window[key]('')
    return None

## Define functional button
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'exit':
        break
    if event == 'clear':
        clear_input()
    if event == 'submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('data has been saved')
        clear_input()
window.close()