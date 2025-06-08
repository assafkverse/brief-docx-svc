# Version 3.9b
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

# Regex patterns
_hdr1 = re.compile(r"^# (.+)")
_hdr2 = re.compile(r"^## (.+)")
_bullet = re.compile(r"^[-*] (.+)")
_table_row = re.compile(r"^\|(.+)\|$")
_table_divider = re.compile(r"^[\|\s:-]+$")
_bold = re.compile(r"\*\*(.+?)\*\*")
_italic = re.compile(r"\*(.+?)\*")
_html_strong = re.compile(r"<strong>(.+?)</strong>", re.IGNORECASE)
_invalid_fname = re.compile(r"[^A-Za-z0-9_.-]")

# Helper functions

def _set_cell_border(cell: _Cell):
    tcPr = cell._tc.get_or_add_tcPr()
    for edge in ("top", "left", "bottom", "right"):
        ln = OxmlElement(f"w:{edge}")
        ln.set(qn("w:val"), "single")
        ln.set(qn("w:sz"), "4")
        ln.set(qn("w:space"), "0")
        ln.set(qn("w:color"), "auto")
        tcPr.append(ln)

def _safe_filename(name: str | None) -> str:
    if not name:
        return f"{uuid4()}.docx"
    name = _invalid_fname.sub("_", name).strip("._") or str(uuid4())
    if not name.lower().endswith(".docx"):
        name += ".docx"
    return name

def _inline_formats(par, text: str):
    pos = 0
    while pos < len(text):
        m_b = _bold.search(text, pos)
        m_i = _italic.search(text, pos)
        m = None
        if m_b and m_i:
            m = m_b if m_b.start() < m_i.start() else m_i
        else:
            m = m_b or m_i
        if not m:
            par.add_run(text[pos:])
            break
        if m.start() > pos:
            par.add_run(text[pos:m.start()])
        run = par.add_run(m.group(1))
        if m.re is _bold:
            run.bold = True
        else:
            run.italic = True
        pos = m.end()

def md_to_docx(md: str, path: str):
    doc = Document()
    if "List Bullet" not in [s.name for s in doc.styles]:
        doc.styles.add_style("List Bullet", WD_STYLE_TYPE.PARAGRAPH).base_style = doc.styles["Normal"]

    lines = md.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        if _table_row.match(line):
            header = [c.strip() for c in line.strip("|").split("|")]
            if i + 1 < len(lines) and _table_divider.match(lines[i+1]):
                rows = []
                i += 2
                while i < len(lines) and _table_row.match(lines[i]):
                    rows.append([c.strip() for c in lines[i].strip("|").split("|")])
                    i += 1
                tbl = doc.add_table(rows=len(rows)+1, cols=len(header))
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
                continue
        if m := _hdr1.match(line):
            doc.add_heading(m.group(1), level=1)
        elif m := _hdr2.match(line):
            doc.add_heading(m.group(1), level=2)
        elif m := _bullet.match(line):
            p = doc.add_paragraph()
            p.style = "List Bullet"
            _inline_formats(p, m.group(1))
        else:
            plain = _html_strong.sub(r"**\\1**", line)
            p = doc.add_paragraph()
            _inline_formats(p, plain)
        i += 1
    doc.save(path)

@app.post("/docx")
def make_docx(payload: dict = Body(...)):
    md = payload["markdown"]
    fname = _safe_filename(payload.get("filename"))
    full = os.path.join(OUTPUT_DIR, fname)
    md_to_docx(md, full)
    host = os.environ.get("RENDER_EXTERNAL_HOSTNAME", "localhost")
    return {"download_url": f"https://{host}/static/{SUB_DIR}/{fname}"}
