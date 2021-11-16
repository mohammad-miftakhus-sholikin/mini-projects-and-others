## Import some packages
import PySimpleGUI as sg
import pandas as pd


## Add some color to the window
sg.theme('DefaultNoMoreNagging')


## Input excel file (note: this file must same in name)
EXCEL_FILE = 'input-form-for-meta-data.xlsx'
df = pd.read_excel(EXCEL_FILE)

## Adding symbols and others
symbol_up =    '▲'
symbol_down =  '▼'

## Adding form input and output
# adding data_i as function
def adding_data(data_i, unit_i):
    # adding data_i
    return sg.Text(data_i, size=(15,1)), sg.InputText(key=data_i, size=(10,1)), sg.Text(unit_i, size=(5,1))
# adding input_i as function
def adding_input(input_i):
    # adding input_i
    return sg.Text(input_i, size=(10,1)), sg.InputText(key=input_i, size=(5,1))
# adding section_i function
def collapse(layout, key):
    """
    Helper function that creates a Column that can be later made hidden, thus appearing "collapsed"
    :param layout: The layout for the section
    :param key: Key used to make this seciton visible / invisible
    :return: A pinned column that can be placed directly into your layout
    :rtype: sg.pin
    """
    return sg.pin(sg.Column(layout, key=key))
# data input section i
section_1 = [
    # adding performance_data_1
    adding_data('performance_data_1', 'unit_1'),
    # adding performance_data_2
    adding_data('performance_data_2', 'unit_2'),
    # adding performance_data_3
    adding_data('performance_data_3', 'unit_3'),
    # adding performance_data_4
    adding_data('performance_data_4', 'unit_4'),
    # adding performance_data_5
    adding_data('performance_data_5', 'unit_5')
    ]
section_2 = [
    # adding digestibility_data_1
    adding_data('digestibility_data_1', 'unit_1'),
    # adding digestibility_data_2
    adding_data('digestibility_data_2', 'unit_2'),
    # adding digestibility_data_3
    adding_data('digestibility_data_3', 'unit_3'),
    # adding digestibility_data_4
    adding_data('digestibility_data_4', 'unit_4'),
    # adding digestibility_data_5
    adding_data('digestibility_data_5', 'unit_5')
    ]
# input and output layout
layout = [
    # form header
    [sg.Text('form to input meta data:')],
    # adding no, study, and author
    adding_input('no'), adding_input('study'), adding_input('author'),
    # adding num_exp, exp_dsgn, and num_rep
    adding_input('num_exp'), adding_input('exp_dsgn'), adding_input('num_rep'),
    # performance data
    [sg.Text(symbol_down, enable_events=True, key='--open-sec-1--'), sg.T('performance data:', enable_events=True)], [collapse(section_1, key='--sec-1--')],
    # digestibility data
    [sg.Text(symbol_down, enable_events=True, key='--open-sec-2--'), sg.T('digestibility data:', enable_events=True)], [collapse(section_2, key='--sec-2--')],
    # submit and exit button
    [sg.Submit(button_text='submit'), sg.Button('clear'), sg.Exit(button_text='exit')]
]

## Define title and layout
window = sg.Window('form to input meta data', layout)

## Delete all values function
def clear_input():
    for key in values:
        window[key]('')
    return None

## Define all section is open
opened_1, opened_2 = True, True

## Define functional button
while True:
    event, values = window.read()
    # exit window
    if event == sg.WIN_CLOSED or event == 'exit':
        break
    # section_1
    if event.startswith('--open-sec-1--'):
        opened_1 = not opened_1
        window['--open-sec-1--'].update(symbol_down if opened_1 else symbol_up)
        window['--sec-1--'].update(visible=opened_1)
    # section_2
    if event.startswith('--open-sec-2--'):
        opened_2 = not opened_2
        window['--open-sec-2--'].update(symbol_down if opened_2 else symbol_up)
        window['--sec-2--'].update(visible=opened_2)
    # clear all input and data
    if event == 'clear':
        clear_input()
    # submit all input and data
    if event == 'submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('data has been saved')
        clear_input()
window.close()