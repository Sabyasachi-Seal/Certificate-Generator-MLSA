import os
import re
# os.system("pip install -r requirements.txt")

import csv
from certificate import *
from docx import Document
from docx2pdf import convert
from openpyxl import Workbook, load_workbook

mailerpath = "Data/Mail.xlsm"
htmltemplatepath = "Data/mailtemplate.html"

# create output folder if not exist
try:
    os.makedirs("Output/Doc")
    os.makedirs("Output/PDF")
except OSError:
    pass


def updatemailer(row, workbook, sheet, email, filepath, sub, body, status, cc=""):
    sheet.cell(row=row, column=1).value = email
    sheet.cell(row=row, column=2).value = cc
    sheet.cell(row=row, column=3).value = sub
    sheet.cell(row=row, column=4).value = body
    sheet.cell(row=row, column=5).value = filepath
    sheet.cell(row=row, column=6).value = status
    workbook.save(filename=mailerpath)


def getworkbook(filename):
    wb = load_workbook(filename=filename, read_only=False, keep_vba=True)
    sheet = wb.active
    return wb, sheet


def gethtmltemplate(htmltemplatepath=htmltemplatepath):
    return open(htmltemplatepath, "r").read()


def getmail(name, event, ambassador):
    sub = f"[MLSA] Certificate of Participation for {name}"
    html = gethtmltemplate(htmltemplatepath)
    body = html.format(name=name, event=event, ambassador=ambassador)
    return sub, body


def get_participants(f):
    data = []  # create empty list
    with open(f, mode="r", encoding='iso-8859-1') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            data.append(row)  # append all results
    return data


def create_docx_files(filename, list_participate):
    wb, sheet = getworkbook(mailerpath)

    event = input("Enter the event name: ")
    ambassador = input("Enter Ambassador Name: ")

    for index, participate in enumerate(list_participate):
        # use original file everytime
        doc = Document(filename)
        participate['Name'] = participate.pop(next(key for key in participate if re.search(r'\bName\b', key)))
        participate['Email'] = participate.pop(next(key for key in participate if re.search(r'\bEmail\b', key)))
        name = participate["Name"]
        email = participate["Email"]

        replace_participant_name(doc, name)
        replace_event_name(doc, event)
        replace_ambassador_name(doc, ambassador)

        doc.save('Output/Doc/{}.docx'.format(name))

        doc.save('Output/Doc/{}.docx'.format(name))

        # ! if your program working slowly, comment this two line and open other 2 line.
        print("Output/{}.pdf Creating".format(name))
        convert('Output/Doc/{}.docx'.format(name), 'Output/Pdf/{}.pdf'.format(name))

        filepath = os.path.abspath('Output/Pdf/{}.pdf'.format(name))

        sub, body = getmail(name, event, ambassador)

        updatemailer(row=index + 2, workbook=wb, sheet=sheet, email=email, filepath=filepath, sub=sub, body=body,
                     status="Send")


# get certificate temple path
certificate_file = "Data/Event Certificate Template.docx"
# get participants path
participate_file = "Data/" + ("ParticipantList.csv" if (input("Test Mode (Y/N): ").lower())[0] == "n" else "temp.csv")

# get participants
list_participate = get_participants(participate_file);

# process data
create_docx_files(certificate_file, list_participate)
