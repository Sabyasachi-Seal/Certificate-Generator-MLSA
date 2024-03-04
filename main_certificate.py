import re
import os
import csv
import zipfile
import uvicorn
import subprocess
import aiofiles
from docx import Document
from typing import AsyncGenerator
from openpyxl import load_workbook
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.middleware.cors import CORSMiddleware
from fastapi import FastAPI, Form, File, UploadFile, Request
from fastapi.responses import FileResponse
from certificate import (
    replace_participant_name,
    replace_event_name,
    replace_ambassador_name,
)

app = FastAPI()

origins = ["*"]  # Adjust this to your frontend's actual origin(s)
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Serve static files (e.g., CSS, JS) from the 'static' folder
app.mount("/static", StaticFiles(directory="static"), name="static")

templates = Jinja2Templates(directory="templates")

mailerpath = "./Data/Mail.xlsm"
htmltemplatepath = "./Data/mailtemplate.html"
zip_filename = "./static/certificates.zip"
csv_file_path = "./Data/participants.csv"

# create output folder if not exist
try:
    os.makedirs("Output/Doc")
    os.makedirs("Output/PDF")
except OSError:
    pass


async def get_data_from_file() -> AsyncGenerator[bytes, None]:
    with open(file=zip_filename, mode="rb") as file_like:
        yield file_like.read()


def clear_mailer_file(wb, sheet):

    # Keep the first row (headers) and delete all other rows
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        for cell in row:
            cell.value = None

    wb.save(filename=mailerpath)


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


async def process_csv(csv_file):

    participant_list = []

    first = True

    with open(csv_file) as csv_file:

        csv_reader = csv.reader(csv_file, delimiter=",")

        name_index = 0
        email_index = 0

        for line in csv_reader:
            if first:
                first = False
                for index, column in enumerate(line):
                    if re.search(r"\bName\b", column):
                        name_index = index
                    elif re.search(r"\bEmail\b", column):
                        email_index = index
                continue
            name = line[name_index]
            email = line[email_index]
            participant_list.append({"Name": name.strip(), "Email": email.strip()})

    return participant_list


def convert_to_pdf(input_path, output_path):
    cmd = ["unoconv", "-f", "pdf", "-o", output_path, input_path]
    process = subprocess.Popen(cmd, stderr=subprocess.PIPE)
    output, error = process.communicate()
    # print(output, error)
    return False, error


def zip_folder(folder_path, zip_filename, additional_files):

    # print(os.listdir(folder_path))

    with zipfile.ZipFile(zip_filename, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, folder_path)
                zipf.write(file_path, arcname)

        # Include additional files
        if additional_files:
            for file_path in additional_files:
                arcname = os.path.relpath(file_path, os.path.dirname(file_path))
                zipf.write(file_path, arcname)


async def create_docx_files(filename, list_participate, event, ambassador):

    wb, sheet = getworkbook(mailerpath)

    clear_mailer_file(wb, sheet)

    os.system("rm -rf Output/Doc/*")
    os.system("rm -rf Output/PDF/*")

    for index, participate in enumerate(list_participate):

        print(participate)

        name = participate["Name"]
        email = participate["Email"]

        if email == "" or name == "":
            continue

        doc = Document(filename)

        replace_participant_name(doc, name)
        replace_event_name(doc, event)
        replace_ambassador_name(doc, ambassador)

        doc.save("./Output/Doc/{}.docx".format(name))

        convert_to_pdf(
            "./Output/Doc/{}.docx".format(name), "./Output/PDF/{}.pdf".format(name)
        )

        filepath = os.path.abspath("./Output/PDF/{}.pdf".format(name))

        sub, body = getmail(name, event, ambassador)

        updatemailer(
            row=index + 2,
            workbook=wb,
            sheet=sheet,
            email=email,
            filepath=filepath,
            sub=sub,
            body=body,
            status="Send",
        )


@app.get("/", response_class=HTMLResponse)
def read_item(request: Request):
    return templates.TemplateResponse("index.html", context={"request": request})


@app.get("/{filepath}")
def get_file(filepath: str):
    file_path = os.path.join("./static", filepath)
    # print(file_path)
    return FileResponse(file_path)


async def get_statinfo():
    with open(zip_filename, "rb") as file:
        yield file.read()


def is_valid_csv(file_content: str) -> bool:
    try:
        # Attempt to parse the CSV content
        csv.reader(file_content.splitlines())
        return True
    except csv.Error:
        return False


@app.post("/generate_certificates")
async def generate_certificates(
    event_name: str = Form(...),
    ambassador_name: str = Form(...),
    participant_file: UploadFile = File(...),
):

    async with aiofiles.open(csv_file_path, "wb") as buffer:

        # Read the content of the uploaded file
        content = await participant_file.read()

        # Check if the content is a valid CSV
        if not is_valid_csv(content.decode()):
            return {"status_code": 400, "message": "Invalid CSV file"}

        await buffer.write(content)

    # get certificate temple path
    certificate_file = "./Data/Event_Certificate_Template.docx"

    # get participants
    list_participate = await process_csv(csv_file_path)

    await create_docx_files(
        certificate_file, list_participate, event=event_name, ambassador=ambassador_name
    )

    additional_files = [mailerpath]

    zip_folder("./Output/PDF", zip_filename, additional_files)

    return FileResponse(
        zip_filename, media_type="application/octet-stream", filename="certificates.zip"
    )


if __name__ == "__main__":
    # Run the app with Uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
