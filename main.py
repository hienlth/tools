from typing import Optional
from fastapi import FastAPI
from fastapi import UploadFile
import shutil
import os

app = FastAPI()


@app.get("/")
async def root():
    return {"message": "Hello World"}

@app.post("/hoat-dong-khac")
def upload_file(file: UploadFile):
    my_filename = os.path.join(os.getcwd(), "data", file.filename)
    with open(my_filename, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    return {"filename": file.filename}