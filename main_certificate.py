
import os

#os.system("pip install -r requirements.txt")

import csv
from certificate import *
from docx import Document
from docx2pdf import convert
import subprocess
import os
import platform
from openpyxl import Workbook, load_workbook

from fastapi import FastAPI, Form, File, UploadFile
from fastapi.responses import FileResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
import os

from certificate import *

app = FastAPI()

# Serve static files (e.g., CSS, JS) from the 'static' folder
app.mount("/static", StaticFiles(directory="static"), name="static")

templates = Jinja2Templates(directory="templates")

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

    workbook.save(filename = mailerpath)


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
    data = [] # create empty list
    with open(f, mode="r", encoding='utf-8') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            data.append(row) # append all results
    return data

def convert_to_pdf(input_path, output_path):
    cmd = [
        "unoconv", 
        "-f", "pdf", 
        "-o", output_path, 
        input_path
    ]
    process = subprocess.Popen(cmd, stderr=subprocess.PIPE)
    output, error = process.communicate()
    return output, error

def create_docx_files(filename, list_participate, event = input("Enter the event name: "), ambassador = input("Enter Ambassador Name: ")):

    wb, sheet = getworkbook(mailerpath)

    # print(list_participate)

    for index, participate in enumerate(list_participate):
        # use original file everytime


        name = participate["Name"]
        email = participate["Email"]

        if email == '' or name == '':
            continue

        doc = Document(filename)

        replace_participant_name(doc, name)
        replace_event_name(doc, event)
        replace_ambassador_name(doc, ambassador)

        doc.save('Output/Doc/{}.docx'.format(name))

        # ! if your program working slowly, comment this two line and open other 2 line.
        print("Output/{}.pdf Creating".format(name))

        if platform.system() == 'Windows':
            convert('Output/Doc/{}.docx'.format(name), 'Output/Pdf/{}.pdf'.format(name))
        else:
            convert_to_pdf('Output/Doc/{}.docx'.format(name), 'Output/PDF/{}.pdf'.format(name))

        filepath = os.path.abspath('Output/PDF/{}.pdf'.format(name))

        sub, body = getmail(name, event, ambassador)

        updatemailer(row=index+2, workbook=wb,  sheet=sheet, email=email, filepath=filepath, sub=sub, body=body, status="Send")

    
@app.get("/")
def read_item(request: dict):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/generate_certificates")
async def generate_certificates(event_name: str = Form(...), ambassador_name: str = Form(...)):
    
    # get certificate temple path
    certificate_file = "Data/Event_Certificate_Template.docx"
    # get participants path
    participate_file = "Data/ParticipantList.csv"

    # get participants
    list_participate = get_participants(participate_file);

    # process data
    create_docx_files(certificate_file, list_participate)
    
    create_docx_files("Data/Event_Certificate_Template.docx", list_participate, event_name, ambassador_name)



    # Zip the generated certificates
    zip_filename = "certificates.zip"

    zip_folder("Output/PDF", zip_filename)

    # Send the zip file to the user
    return FileResponse(zip_filename, media_type="application/zip", headers={"Content-Disposition": "attachment; filename=certificates.zip"})

if __name__ == '__main__':
    import uvicorn

    # Run the app with Uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
