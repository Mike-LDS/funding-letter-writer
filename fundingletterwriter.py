import docx
from docxtpl import DocxTemplate
import datetime
import PySimpleGUI as sg
import os
import requests
import json
import sys
import time

url = "https://api.teachworks.com/v1/students"

payload={}

headers = {
  'Authorization': 'Token token=st_live_fvXWm27jS_Q_BRtwYWH5lg',
  'Content-Type': 'application/json'
}

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

sg.theme("DarkTeal7")
layout = [
    [sg.T("")],
    [sg.Text("First Name: "),sg.Input(key='FIRST',do_not_clear=False,size=(20, 20)),sg.Text("Last Name: "),sg.Input(key='LAST',do_not_clear=False,size=(20, 20)),sg.Button("Add Student")],
    [sg.T("")],
    [sg.Text("Choose Letter Template:  "), sg.Input(), sg.FileBrowse(file_types=(('DOCX Files', '*.DOCX'),), key='-TEMP-')],
    [sg.T("")],
    [sg.Button("RUN"),sg.Text("", key='-TEXT-')]]

# Building Window
window = sg.Window('FUNDING LETTER WRITER', layout, size=(550,175))

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
        elif values['-TEMP-'] == '':
            sg.Popup('No Funding Template', title='ERROR')
        else:
            if values['FIRST'] != '' and values['LAST'] != '':
                stu_first.append(values['FIRST'])
                stu_last.append(values['LAST'])
            try:
                letter_count = 0
                for x in range(0,len(stu_first)):
                    stu_first[x] = stu_first[x].replace(' ', '+')
                    stu_last[x] = stu_last[x].replace(' ', '+')
                    stu_url = url+'?first_name='+stu_first[x]+'&last_name='+stu_last[x]
                    response = requests.request("GET", stu_url, headers=headers, data=payload)
                    student = response.json()
                    content['first_name'] = student[0]["first_name"]
                    content['last_name'] = student[0]["last_name"]
                    cus = student[0]["custom_fields"]
                    for i in range(0,len(cus)):
                        if cus[i]['field_id'] == 15020:
                            content['ld_discription'] = cus[i]['value']
                        elif cus[i]['field_id'] == 15022:
                            content['program_objective'] = cus[i]['value']
                    template = DocxTemplate(values['-TEMP-'])
                    template.render(content)
                    path = os.path.dirname(sys.executable).split('fundingletterwriter.app',-1)[0]
                    template.save(path+'/Letters/'+content['first_name']+'_'+content['last_name']+'_'+content['year']+'_Funding_Quote.docx')
                    letter_count = letter_count + 1
                    letters = len(stu_first)
                    window['-TEXT-'].update(str(letter_count) +" OF "+ str(letters)+" LETTERS GENERATED")
                    time.sleep(0.5)

                stu_first = []
                stu_last = []
            except Exception as e:
                sg.Popup(e, title='ERROR')

window.close()


