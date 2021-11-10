import csv
import docx
from docxtpl import DocxTemplate
import datetime

stu_first = ['Kaden', 'Oliver']
stu_last = ['Le','Van Santen']

content = {
    'first_name': 'Georgia',
    'last_name': 'Mckenzie-Gray',
    'month': datetime.datetime.now().strftime('%b'),
    'day': datetime.datetime.now().strftime('%d'),
    'year': datetime.datetime.now().strftime('%Y'),
    'ld_discription': 'Georgia is a dog. She does not like to be alone.',
    'program_objective': '1. Learn to spend more time alone.'
    }


with open('users.csv', newline='') as csvfile:
    reader = csv.reader(csvfile)
    for row in reader:
        for x in range(len(stu_first)):
            if (row[4] == stu_first[x]) and (row[5] == stu_last[x]):
                content['first_name'] = row[4]
                content['last_name'] = row[5]
                content['ld_discription'] = row[27]
                content['program_objective'] = row[28]
                template = DocxTemplate('2021-2022_Funding_Quote_Template.docx')
                template.render(content)
                template.save(content['first_name']+'_'+content['last_name']+'_2021-2022_Funding_Quote.docx')
