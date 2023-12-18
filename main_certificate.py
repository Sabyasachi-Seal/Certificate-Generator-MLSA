
import os
import zipfile
import uvicorn
import platform
import subprocess
from typing import List
from docx import Document
from docx2pdf import convert
from openpyxl import load_workbook
from socketio import AsyncServer, ASGIApp
from fastapi.responses import HTMLResponse 
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi import FastAPI, Form, File, UploadFile, Request, WebSocket, Depends, WebSocket, WebSocketDisconnect
from certificate import replace_participant_name, replace_event_name, replace_ambassador_name



app = FastAPI()

sio = AsyncServer(async_mode="asgi", cors_allowed_origins="*")
app.add_middleware(ASGIApp, socketio=sio)

class WebSocketConnectionManager:
    def __init__(self):
        self.active_connections: List = []

    async def connect(self, websocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket):
        self.active_connections.remove(websocket)

    async def send_personal_message(self, message, websocket):
        await websocket.send_text(message)

    async def broadcast(self, message):
        for connection in self.active_connections:
            await connection.send_text(message)

manager = WebSocketConnectionManager()
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

async def process_csv(csv_file):
    participant_list = []
    content = (await csv_file.read()).decode("utf-8").splitlines()

    # Skip header line
    _ = content.pop(0)

    for line in content:
        name, email = line.split(",")
        participant_list.append({"Name": name, "Email": email})

    return participant_list

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

def zip_folder(folder_path, zip_filename, additional_files):
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
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

async def create_docx_files(filename, list_participate, event, ambassador, websocket):

    wb, sheet = getworkbook(mailerpath)

    clear_mailer_file(wb, sheet)

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

        await websocket.send_text(f"PDF {name} has been generated.")

    await websocket.send_text("All PDFs generated. Creating the zip file.")
    await websocket.send_text("Sending the zip file to the user.")
    await websocket.send_text("Done!")

# Helper function to get the WebSocket
def get_websocket():
    return WebSocketConnectionManager.get_instance().get_websocket()

@app.get("/", response_class=HTMLResponse)
def read_item(request: Request):
    return templates.TemplateResponse("index.html", context={"request": request})

@app.post("/generate_certificates")
async def generate_certificates(event_name: str = Form(...), ambassador_name: str = Form(...), participant_file: UploadFile = File(...), websocket: WebSocket = Depends(get_websocket)):
    
    await websocket.accept()
    
    # get certificate temple path
    certificate_file = "Data/Event_Certificate_Template.docx"

    # get participants
    list_participate = await process_csv(participant_file);
    
    await create_docx_files(certificate_file, list_participate, event=event_name, ambassador=ambassador_name, websocket=websocket)

    # Zip the generated certificates
    zip_filename = "certificates.zip"

    additional_files = [mailerpath]

    zip_folder("Output/PDF", zip_filename, additional_files)

    os.system("rm -rf Output/Doc/*")
    os.system("rm -rf Output/PDF/*")

    # Send the zip file to the user
    return FileResponse(zip_filename, media_type="application/zip", headers={"Content-Disposition": "attachment; filename=certificates.zip"})

# Define WebSocket endpoint
@sio.event
async def connect(sid, environ):
    print(f"WebSocket {sid} connected.")
    await manager.connect(sid)

@sio.event
async def disconnect(sid):
    print(f"WebSocket {sid} disconnected.")
    await manager.disconnect(sid)

@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    await manager.connect(websocket)
    try:
        while True:
            data = await websocket.receive_text()
            if data == "close":
                await manager.disconnect(websocket)
                break
    except WebSocketDisconnect:
        manager.disconnect(websocket)



if __name__ == '__main__':
    # Run the app with Uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
