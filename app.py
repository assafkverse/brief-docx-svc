from fastapi import FastAPI, Body
from fastapi.staticfiles import StaticFiles
from docx import Document
from uuid import uuid4
import os

app = FastAPI()

OUTPUT_DIR = "generated"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# serve runtime files
app.mount("/static", StaticFiles(directory=OUTPUT_DIR), name="static")

@app.post("/docx")
def make_docx(payload: dict = Body(...)):
    md = payload["markdown"]
    filename = f"{uuid4()}.docx"
    path = os.path.join(OUTPUT_DIR, filename)

    doc = Document()
    for line in md.splitlines():
        doc.add_paragraph(line)
    doc.save(path)

    host = os.environ.get("RENDER_EXTERNAL_HOSTNAME", "localhost")
    return {"download_url": f"https://{host}/static/{filename}"}
