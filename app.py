from fastapi import FastAPI, Body
from fastapi.staticfiles import StaticFiles
from docx import Document
from uuid import uuid4
import os, re

app = FastAPI()

OUTPUT_DIR = "generated"
os.makedirs(OUTPUT_DIR, exist_ok=True)
app.mount("/static", StaticFiles(directory=OUTPUT_DIR), name="static")

def safe_filename(name: str) -> str:
    """remove illegal chars + limit length"""
    name = re.sub(r"[^\w\-\.]", "_", name)      # החלף רווחים/תווים לא חוקיים
    return name[:60] if len(name) > 60 else name

@app.post("/docx")
def make_docx(payload: dict = Body(...)):
    md = payload["markdown"]
    raw_name = payload.get("filename")           # חדש
    if raw_name:
        filename = safe_filename(raw_name)
        if not filename.endswith(".docx"):
            filename += ".docx"
    else:
        filename = f"{uuid4()}.docx"

    path = os.path.join(OUTPUT_DIR, filename)

    doc = Document()
    for line in md.splitlines():
        doc.add_paragraph(line)
    doc.save(path)

    host = os.environ.get("RENDER_EXTERNAL_HOSTNAME", "localhost")
    return {"download_url": f"https://{host}/static/{filename}"}
