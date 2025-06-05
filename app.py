from fastapi import FastAPI, Body
from fastapi.staticfiles import StaticFiles
from docx import Document
from uuid import uuid4
import os
import re

app = FastAPI()

# ---------------------------------------------------------------------------
#  Configuration
# ---------------------------------------------------------------------------
BASE_DIR = "generated"          # persistent disk or local dir
SUB_DIR  = "downloads"          # keep docs in a nice sub‑folder
OUTPUT_DIR = os.path.join(BASE_DIR, SUB_DIR)

os.makedirs(OUTPUT_DIR, exist_ok=True)

# expose /static/** so docs are downloadable via HTTPS
app.mount("/static", StaticFiles(directory=BASE_DIR), name="static")

# ---------------------------------------------------------------------------
#  Helpers
# ---------------------------------------------------------------------------
_invalid = re.compile(r"[^A-Za-z0-9_.-]")

def safe_filename(raw: str | None) -> str:
    """Sanitise the requested filename or fall back to UUID."""
    if not raw:
        return f"{uuid4()}.docx"
    name = _invalid.sub("_", raw).strip("._")
    if not name:
        name = str(uuid4())
    if not name.lower().endswith(".docx"):
        name += ".docx"
    return name

# ---------------------------------------------------------------------------
#  Route: POST /docx
# ---------------------------------------------------------------------------
@app.post("/docx")
def make_docx(payload: dict = Body(...)):
    """Convert incoming Markdown to a Word document and return download URL."""
    md: str = payload["markdown"]
    filename = safe_filename(payload.get("filename"))
    full_path = os.path.join(OUTPUT_DIR, filename)

    # create .docx
    doc = Document()
    for line in md.splitlines():
        doc.add_paragraph(line)
    doc.save(full_path)

    host = os.environ.get("RENDER_EXTERNAL_HOSTNAME", "localhost")
    return {
        "download_url": f"https://{host}/static/{SUB_DIR}/{filename}"
    }
