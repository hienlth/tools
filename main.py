from typing import Optional
from fastapi import FastAPI
from fastapi import UploadFile
from fastapi.responses import FileResponse
import shutil
import os
from services import generate_report

app = FastAPI()


@app.get("/")
async def root():
    return {"message": "FIT-HCMUE"}

@app.post("/hoat-dong-khac")
def upload_file(file: UploadFile):
    my_filename = os.path.join(os.getcwd(), "data", file.filename)
    with open(my_filename, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    
    print(my_filename)
    template_path=os.path.join(os.getcwd(), "templates", "Template.docx")
    print(template_path)
    file_export_name="FIT_TTHDK_HK1_2025_2026.docx"
    export_path=os.path.join(os.getcwd(), "outputs", file_export_name)
    generate_report.generate_report(
        xlsx_path=my_filename,
        template_path=template_path,
        output_path=export_path,
    )
    return FileResponse(
        path=export_path,
        filename=file_export_name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )