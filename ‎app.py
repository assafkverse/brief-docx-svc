from fastapi import FastAPI, Body
from fastapi.responses import FileResponse
from uuid import uuid4
import os
from docx import Document   # pip install python-docx

app = FastAPI()
OUTPUT_DIR = "generated"
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.post("/docx")
def make_docx(payload: dict = Body(...)):
    md = payload["markdown"]
    filename = f"{uuid4()}.docx"
    path = os.path.join(OUTPUT_DIR, filename)

    doc = Document()
    for line in md.splitlines():
        doc.add_paragraph(line)
    doc.save(path)

    # Render מגישה קבצים ב־/static אוטומטית אם נגדיר
    return {"download_url": f"https://{os.environ['RENDER_EXTERNAL_HOSTNAME']}/static/{filename}"}
