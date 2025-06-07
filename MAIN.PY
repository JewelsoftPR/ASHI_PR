from fastapi import FastAPI, UploadFile, File, Form
from typing import Optional
import shutil
import os
import subprocess

app = FastAPI()

UPLOAD_DIR = "uploaded_files"
os.makedirs(UPLOAD_DIR, exist_ok=True)


# === /split-pdf endpoint ===
@app.post("/split-pdf")
async def split_pdf(file: UploadFile = File(...)):
    file_path = os.path.join(UPLOAD_DIR, file.filename)
    with open(file_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    subprocess.run(["python", "split_pdf.py", file_path])  # assumes script supports args
    return {"message": f"Split PDF executed for {file.filename}"}


# === /top-style-picture endpoint ===
@app.post("/top-style-picture")
async def top_style_picture(file: UploadFile = File(...), treated: UploadFile = File(...)):
    path1 = os.path.join(UPLOAD_DIR, file.filename)
    path2 = os.path.join(UPLOAD_DIR, treated.filename)
    with open(path1, "wb") as f:
        shutil.copyfileobj(file.file, f)
    with open(path2, "wb") as f:
        shutil.copyfileobj(treated.file, f)

    subprocess.run(["python", "Top Style Picture Report.py", path1, path2])
    return {"message": f"Report generated for {file.filename}"}


# === Add more endpoints for each script similarly ===
