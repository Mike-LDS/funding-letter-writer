import csv
import docx
from docxtpl import DocxTemplate
import datetime
import PySimpleGUI as sg
import os

content = {
    'first_name': 'Georgia',
    'last_name': 'Mckenzie-Gray',
    'month': datetime.datetime.now().strftime('%b'),
    'day': datetime.datetime.now().strftime('%d'),
    'year': datetime.datetime.now().strftime('%Y'),
    'ld_discription': 'Georgia is a dog. She does not like to be alone.',
    'program_objective': '1. Learn to spend more time alone.'
    }

stu_first = []
stu_last = []

sg.theme("DarkBlue9")
layout = [
    [sg.T("")],
    [sg.Text("First Name: "),sg.Input(key='FIRST',do_not_clear=False,size=(20, 20)),sg.Text("Last Name: "),sg.Input(key='LAST',do_not_clear=False,size=(20, 20)),sg.Button("Add Student")],
    [sg.T("")],
    [sg.Text("Choose Student Export: "), sg.Input(), sg.FileBrowse(file_types=(('CSV Files', '*.CSV'),), key='-USERS-')],
    [sg.Text('Go to TC -> System -> Exports -> Click on Users -> Select Students from drop down menu -> Click Submit')],
    [sg.T("")],
    [sg.Text("Choose Letter Template:  "), sg.Input(), sg.FileBrowse(file_types=(('DOCX Files', '*.DOCX'),), key='-TEMP-')],
    [sg.T("")],
    [sg.Button("RUN"),sg.Text("", key='-TEXT-')]]

# Building Window
window = sg.Window('FUNDING LETTER WRITER', layout, size=(600,250))

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event=="EXIT":
        break
    elif event == "Add Student":
        if values['FIRST'] == '':
            sg.Popup('No Student First Name', title='ERROR')
        elif values['LAST'] == '':
            sg.Popup('No Student Last Name', title='ERROR')
        else:
            stu_first.append(values['FIRST'])
            stu_last.append(values['LAST'])
    elif event == "RUN":
        if values['FIRST'] == '' and len(stu_first) == 0:
            sg.Popup('No Student First Name', title='ERROR')
        elif values['LAST'] == '' and len(stu_last) == 0:
            sg.Popup('No Student Last Name', title='ERROR')
        else:
            if values['FIRST'] != '' and values['LAST'] != '':
                stu_first.append(values['FIRST'])
                stu_last.append(values['LAST'])
            try:
                with open(values['-USERS-'], newline='') as csvfile:
                    reader = csv.reader(csvfile)
                    letter_count = 0
                    for row in reader:
                        for x in range(len(stu_first)):
                            if (row[4].lower() == stu_first[x].lower()) and (row[5].lower() == stu_last[x].lower()):
                                content['first_name'] = row[4]
                                content['last_name'] = row[5]
                                content['ld_discription'] = row[27]
                                content['program_objective'] = row[28]
                                template = DocxTemplate(values['-TEMP-'])
                                template.render(content)
                                path = os.path.dirname(sys.executable).split('fundingletterwriter.app',-1)[0]
                                template.save(path+'/Letters/'+content['first_name']+'_'+content['last_name']+'_'+content['year']+'_Funding_Quote.docx')
                                letter_count = letter_count + 1

                letters = len(stu_first)
                stu_first = []
                stu_last = []
                window['-TEXT-'].update(str(letter_count) +" OF "+ str(letters)+" LETTERS GENERATED")
            except Exception as e:
                sg.Popup(e, title='ERROR')

window.close()


