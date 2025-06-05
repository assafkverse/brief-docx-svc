from fastapi import FastAPI, Body
from fastapi.staticfiles import StaticFiles
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from uuid import uuid4
import os, re

app = FastAPI()
BASE_DIR = "generated/downloads"
os.makedirs(BASE_DIR, exist_ok=True)
app.mount("/static", StaticFiles(directory="generated"), name="static")

hdr1 = re.compile(r"^# (.*)")
hdr2 = re.compile(r"^## (.*)")
bullet = re.compile(r"^[-*] (.*)")

def md_to_docx(md: str, path: str):
    doc = Document()

    # guarantee bullet style exists
    if "List Bullet" not in [s.name for s in doc.styles]:
        doc.styles.add_style("List Bullet", WD_STYLE_TYPE.PARAGRAPH).base_style = doc.styles["Normal"]

    for line in md.splitlines():
        if m := hdr1.match(line):
            doc.add_heading(m.group(1), level=1)
        elif m := hdr2.match(line):
            doc.add_heading(m.group(1), level=2)
        elif m := bullet.match(line):
            p = doc.add_paragraph(m.group(1))
            p.style = "List Bullet"
        else:
            doc.add_paragraph(line)
    doc.save(path)

@app.post("/docx")
def make_docx(payload: dict = Body(...)):
    md = payload["markdown"]
    fname = f"{uuid4()}.docx"
    full = os.path.join(BASE_DIR, fname)

    md_to_docx(md, full)

    host = os.environ.get("RENDER_EXTERNAL_HOSTNAME", "localhost")
    return {"download_url": f"https://{host}/static/downloads/{fname}"}
