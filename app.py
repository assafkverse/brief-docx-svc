from fastapi import FastAPI, Body
from fastapi.staticfiles import StaticFiles
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from uuid import uuid4
import os, re

# ---------------------------------------------------------------------------
#  FastAPI app + static mount
# ---------------------------------------------------------------------------
app = FastAPI()
BASE_DIR = "generated"
SUB_DIR = "downloads"
OUTPUT_DIR = os.path.join(BASE_DIR, SUB_DIR)

os.makedirs(OUTPUT_DIR, exist_ok=True)
app.mount("/static", StaticFiles(directory=BASE_DIR), name="static")

# ---------------------------------------------------------------------------
#  Regex helpers
# ---------------------------------------------------------------------------
_hdr1 = re.compile(r"^# (.+)")
_hdr2 = re.compile(r"^## (.+)")
_bullet = re.compile(r"^[-*] (.+)")
_table_row = re.compile(r"^\|(.+)\|$")  # lines that start+end with pipe
_table_divider = re.compile(r"^[\|\s:-]+$")
_invalid_fname = re.compile(r"[^A-Za-z0-9_.-]")

# ---------------------------------------------------------------------------
#  Markdown → DOCX
# ---------------------------------------------------------------------------

def md_to_docx(md: str, path: str):
    doc = Document()

    # ensure bullet style exists
    if "List Bullet" not in [s.name for s in doc.styles]:
        doc.styles.add_style("List Bullet", WD_STYLE_TYPE.PARAGRAPH).base_style = doc.styles["Normal"]

    lines = md.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]

        # Table detection: header |---|---|
        if _table_row.match(line):
            header_cells = [c.strip() for c in line.strip("|").split("|")]
            if i + 1 < len(lines) and _table_divider.match(lines[i+1]):
                # start collecting rows
                rows = []
                i += 2  # skip divider
                while i < len(lines) and _table_row.match(lines[i]):
                    row_cells = [c.strip() for c in lines[i].strip("|").split("|")]
                    rows.append(row_cells)
                    i += 1
                # build table
                tbl = doc.add_table(rows=len(rows)+1, cols=len(header_cells))
                # header
                for c, text in enumerate(header_cells):
                    tbl.rows[0].cells[c].text = text
                # body
                for r, row_cells in enumerate(rows, start=1):
                    for c, text in enumerate(row_cells):
                        tbl.rows[r].cells[c].text = text
                continue  # already advanced i inside loop
        # Headings & bullets
        if m := _hdr1.match(line):
            doc.add_heading(m.group(1), level=1)
        elif m := _hdr2.match(line):
            doc.add_heading(m.group(1), level=2)
        elif m := _bullet.match(line):
            p = doc.add_paragraph(m.group(1))
            p.style = "List Bullet"
        else:
            doc.add_paragraph(line)
        i += 1

    doc.save(path)

# ---------------------------------------------------------------------------
#  Utilities
# ---------------------------------------------------------------------------

def safe_filename(raw: str | None) -> str:
    if not raw:
        return f"{uuid4()}.docx"
    name = _invalid_fname.sub("_", raw).strip("._")
    if not name:
        name = str(uuid4())
    if not name.lower().endswith(".docx"):
        name += ".docx"
    return name

# ---------------------------------------------------------------------------
#  Route
# ---------------------------------------------------------------------------
@app.post("/docx")
def make_docx(payload: dict = Body(...)):
    md: str = payload["markdown"]
    fname = safe_filename(payload.get("filename"))
    full_path = os.path.join(OUTPUT_DIR, fname)

    md_to_docx(md, full_path)

    host = os.environ.get("RENDER_EXTERNAL_HOSTNAME", "localhost")
    return {"download_url": f"https://{host}/static/{SUB_DIR}/{fname}"}
