<!-- Version 3.9 -->
"""FastAPI service that converts Markdown to Word and returns a download URL.
Supports headings, bullet lists, bold/italic, and GitHub‑style tables with
borders + bold header row. Files saved under /static/downloads/ ."""

from fastapi import FastAPI, Body
from fastapi.staticfiles import StaticFiles
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.table import _Cell
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from uuid import uuid4
import os, re

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
_table_row = re.compile(r"^\|(.+)\|$")
_table_divider = re.compile(r"^[\|\s:-]+$")
_bold = re.compile(r"\*\*(.+?)\*\*")
_italic = re.compile(r"\*(.+?)\*")
_html_strong = re.compile(r"<strong>(.+?)</strong>", re.IGNORECASE)
_invalid_fname = re.compile(r"[^A-Za-z0-9_.-]")

# ---------------------------------------------------------------------------
#  Util functions
# ---------------------------------------------------------------------------

def _set_cell_border(cell: _Cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for edge in ("top", "left", "bottom", "right"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "4")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "auto")
        tcPr.append(element)


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
#  Markdown → DOCX conversion
# ---------------------------------------------------------------------------

def _apply_inline_formats(paragraph, text: str):
    """Handle **bold** and *italic* inside a paragraph."""
    pos = 0
    while pos < len(text):
        bold_m = _bold.search(text, pos)
        italic_m = _italic.search(text, pos)
        next_m = None
        if bold_m and italic_m:
            next_m = bold_m if bold_m.start() < italic_m.start() else italic_m
        else:
            next_m = bold_m or italic_m
        if not next_m:
            paragraph.add_run(text[pos:])
            break
        # plain text before format
        if next_m.start() > pos:
            paragraph.add_run(text[pos:next_m.start()])
        run = paragraph.add_run(next_m.group(1))
        if next_m.re is _bold:
            run.bold = True
        else:
            run.italic = True
        pos = next_m.end()


def md_to_docx(md: str, path: str):
    doc = Document()
    if "List Bullet" not in [s.name for s in doc.styles]:
        doc.styles.add_style("List Bullet", WD_STYLE_TYPE.PARAGRAPH).base_style = doc.styles["Normal"]

    lines = md.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        # Table detection
        if _table_row.match(line):
            header = [c.strip() for c in line.strip("|").split("|")]
            if i + 1 < len(lines) and _table_divider.match(lines[i+1]):
                rows = []
                i += 2
                while i < len(lines) and _table_row.match(lines[i]):
                    rows.append([c.strip() for c in lines[i].strip("|").split("|")])
                    i += 1
                tbl = doc.add_table(rows=len(rows)+1, cols=len(header))
                # header row bold
                for c, txt in enumerate(header):
                    cell = tbl.rows[0].cells[c]
                    cell.text = txt
                    for run in cell.paragraphs[0].runs:
                        run.bold = True
                    _set_cell_border(cell)
                for r, row in enumerate(rows, start=1):
                    for c, txt in enumerate(row):
                        cell = tbl.rows[r].cells[c]
                        cell.text = txt
                        _set_cell_border(cell)
                continue  # already advanced i
        # Regular elements
        if m := _hdr1.match(line):
            doc.add_heading(m.group(1), level=1)
        elif m := _hdr2.match(line):
            doc.add_heading(m.group(1), level=2)
        elif m := _bullet.match(line):
            p = doc.add_paragraph()
            p.style = "List Bullet"
            _apply_inline_formats(p, m.group(1))
        else:
            # strip HTML <strong>
            text = _html_strong.sub(r"**\1**", line)
            p = doc.add_paragraph()
            _apply_inline_formats(p, text)
        i += 1
    doc.save(path)

# ---------------------------------------------------------------------------
#  Route: POST /docx
# ---------------------------------------------------------------------------
@app.post("/docx")
def make_docx(payload: dict = Body(...)):
    md = payload["markdown"]
    fname = safe_filename(payload.get("filename"))
    full = os.path.join(OUTPUT_DIR, fname)
    md_to_docx(md, full)
    host = os.environ.get("RENDER_EXTERNAL_HOSTNAME", "localhost")
    return {"download_url": f"https://{host}/static/{SUB_DIR}/{fname}"}
