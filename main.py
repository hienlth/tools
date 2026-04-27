from fastapi import FastAPI, UploadFile, Request, File, Form
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
import shutil
import os
from services import generate_report

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={}
    )


@app.get("/report/ra-de-duyet-de-cham-thi", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="rade_chamthi.html",
        context={}
    )


@app.get("/report/khoa-luan-tot-nghiep", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="khoaluan_dh.html",
        context={}
    )


@app.post("/hoat-dong-khac")
def upload_file(
    report_type: str = Form(default="rade_chamthi"),
    file: UploadFile = File(...)
):
    print('report type', report_type)
    my_filename = os.path.join(os.getcwd(), "data", file.filename)
    with open(my_filename, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    
    template_path=os.path.join(os.getcwd(), "templates", "Template.docx")
    data_output = generate_report.generate_report(
        xlsx_path=my_filename,
        template_path=template_path,
        report_type=report_type
    )

    return StreamingResponse(
        data_output,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={
            "Content-Disposition": "attachment; filename=processed_report.docx"
        }
    )