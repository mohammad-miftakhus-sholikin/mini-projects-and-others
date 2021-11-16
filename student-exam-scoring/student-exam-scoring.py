### INPUT AND PACKAGES ###
## Import some packages
import PySimpleGUI as sg
import pandas as pd


## Add some color to the window
sg.theme('DefaultNoMoreNagging')


## Input excel file (note: this file must same in name)
EXCEL_FILE = 'student-exam-score.xlsx'
df = pd.read_excel(EXCEL_FILE)


## Adding symbols and others
symbol_up = '︿'
symbol_down = '﹀'



### LAYOUT ###
## Adding form input and output

# adding data_i as function
def adding_data(data_i):
    # adding data_i
    return sg.Text(data_i, size=(15,1)), sg.InputText(key=data_i, size=(10,1))

# adding input_i as function
def adding_input(input_i):
    # adding input_i
    return sg.Text(input_i, size=(15,1)), sg.InputText(key=input_i, size=(10,1))

# adding section_i function
def collapse(layout, key, visible):
    """
    Helper function that creates a Column that can be later made hidden, thus appearing "collapsed"
    :param layout: The layout for the section
    :param key: Key used to make this seciton visible / invisible
    :return: A pinned column that can be placed directly into your layout
    :rtype: sg.pin
    """
    return sg.pin(sg.Column(layout, key=key, visible=visible))

# section 1: general information of student
section_1 = [
    # adding no, student_id, and student_name
    adding_input('no'), adding_input('student_id'), adding_input('student_name'),
    # adding class_group and semester
    adding_input('class_group'), adding_input('semester ')
    ]

# section 2: course 1
section_2 = [
    # adding course_1
    adding_data('course_1_part_1'),
    # adding course_2
    adding_data('course_1_part_2'),
    # adding course_3
    adding_data('course_1_part_3'),
    # adding course_4
    adding_data('course_1_part_4'),
    # adding course_5
    adding_data('course_1_part_5')
    ]

# section 3: course 2
section_3 = [
    # adding course_2_1
    adding_data('course_2_part_1'),
    # adding course_2_2
    adding_data('course_2_part_2'),
    # adding course_2_3
    adding_data('course_2_part_3'),
    # adding course_2_4
    adding_data('course_2_part_4'),
    # adding course_2_5
    adding_data('course_2_part_5')
    ]

# input and output layout
layout = [
    # general information of student
    [sg.Text(symbol_down, enable_events=True, key='__open_sec_1__'), sg.Text('general information of student:', enable_events=True)], [collapse(section_1, '__sec_1__', True)],
    # course 1
    [sg.Text(symbol_down, enable_events=True, key='__open_sec_2__'), sg.Text('course 1:', enable_events=True)], [collapse(section_2, '__sec_2__', False)],
    # course 2
    [sg.Text(symbol_down, enable_events=True, key='__open_sec_3__'), sg.Text('course 2:', enable_events=True)], [collapse(section_3, '__sec_3__', False)],
    # submit and exit button
    [sg.Submit(button_text='submit'), sg.Button('clear'), sg.Exit(button_text='exit')]
    ]


## Define title and layout
# size minimal layout
wdth, hght = MIN_SIZE = (350, 350)
# layout
window = sg.Window('form to input meta data', layout, resizable=True, finalize=True)
# set up weidth and height of window 1
width, height = tuple(map(int, window.TKroot.geometry().split("+")[0].split("x")))
# set up weidth and height of window 2
window.TKroot.minsize(max(wdth, width), max(hght, height))



### LOGIG OPERATION ###
## Delete all values function
def clear_input():
    for key in values:
        window[key]('')
    return None


## Define all section is open and close
opened_1, opened_2, opened_3 = True, True, True


## Define functional button
while True:
    event, values = window.read()
    # exit window
    if event == sg.WIN_CLOSED or event == 'exit':
        break
    # section_1
    if event.startswith('__open_sec_1__'):
        opened_1 = not opened_1
        window['__open_sec_1__'].update(symbol_down if opened_1 else symbol_up)
        window['__sec_1__'].update(visible=opened_1)
    # section_2
    if event.startswith('__open_sec_2__'):
        opened_2 = not opened_2
        window['__open_sec_2__'].update(symbol_down if opened_2 else symbol_up)
        window['__sec_2__'].update(visible=opened_2)
    # section_3
    if event.startswith('__open_sec_3__'):
        opened_3 = not opened_3
        window['__open_sec_3__'].update(symbol_down if opened_3 else symbol_up)
        window['__sec_3__'].update(visible=opened_3)
    # clear all input and data
    if event == 'clear':
        clear_input()
    # submit all input and data
    if event == 'submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('data has been saved')
window.close()