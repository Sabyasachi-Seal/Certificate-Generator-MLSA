
import os

os.system("pip install -r requirements.txt")

mailerpath="Data/Mail.xlsm"
htmltemplatepath="Data/mailtemplate.html"

import csv
from certificate import *
from docx import Document
from docx2pdf import convert
from openpyxl import Workbook, load_workbook

def updatemailer(row, workbook, sheet, email, filepath, sub, body, status, cc="", mailerpath="Data/Mail.xlsm"):

    sheet.cell(row=row, column=1).value = email
    sheet.cell(row=row, column=2).value = cc
    sheet.cell(row=row, column=3).value = sub
    sheet.cell(row=row, column=4).value = body
    sheet.cell(row=row, column=5).value = filepath
    sheet.cell(row=row, column=6).value = status
    workbook.save(filename = mailerpath)

def getworkbook(filename):
    wb = load_workbook(filename=filename, read_only=False, keep_vba=True)
    sheet = wb.active
    return wb, sheet

def gethtmltemplate(htmltemplatepath="Data/mailtemplate.html"):
    return open(htmltemplatepath, "r").read()

def getmail(name, event, ambassador):
    sub = f"[MLSA] Certificate of Participation for {name}"
    html = gethtmltemplate(htmltemplatepath)
    body = html.format(name=name, event=event, ambassador=ambassador)
    return sub, body

def get_participants(f):
    data = [] # create empty list
    with open(f, mode="r", encoding='utf-8') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            data.append(row) # append all results
    return data

def create_docx_files(filename, list_participate, event, ambassador):

    main_docx = 'Output/Doc/{}.docx'
    main_pdf = 'Output/Pdf/{}.pdf'

    wb, sheet = getworkbook(mailerpath)

    for index, participate in enumerate(list_participate):
        # use original file everytime
        doc = Document(filename)

        name = participate["Name"]
        email = participate["Email"]

        replace_participant_name(doc, name)
        replace_event_name(doc, event)
        replace_ambassador_name(doc, ambassador)

        doc.save(main_docx.format(name))

        doc.save(main_docx.format(name))

        # ! if your program working slowly, comment this two line and open other 2 line.
        print(main_pdf.format(name)+" Creating")
        #convert(main_docx.format(name), main_pdf.format(name))

        #filepath = os.path.abspath(main_pdf.format(name))

        sub, body = getmail(name, event, ambassador)

        #updatemailer(row=index+2, workbook=wb,  sheet=sheet, email=email, filepath=filepath, sub=sub, body=body, status="Send")

def main(participant_list_name, event, ambassador):
    # create output folder if not exist
    try:
        os.makedirs("Output/Doc")
        os.makedirs("Output/PDF")
    except OSError:
        pass

    # get certificate temple path
    certificate_file = "Data/Event Certificate Template.docx"
    # get participants path
    participate_file = "Data/"+ participant_list_name

    # get participants
    list_participate = get_participants(participate_file);

    # process data
    create_docx_files(certificate_file, list_participate, event, ambassador)



