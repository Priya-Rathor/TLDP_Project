from fastapi import FastAPI
from fastapi.responses import FileResponse
import os

app = FastAPI()

FILE_DIR = "files"  # folder where ppt is stored
FILE_NAME = "output.pptx"

@app.get("/download-ppt")
async def download_ppt():
    file_path = os.path.join(FILE_DIR, FILE_NAME)
    if os.path.exists(file_path):
        return FileResponse(file_path, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", filename=FILE_NAME)
    return {"error": "File not found"}
