# Version 3.9a

"""FastAPI service that converts Markdown to Word and returns a download URL.
Supports headings, bullet lists, bold/italic, and GitHub-style tables with
borders + bold header row. Files saved under /static/downloads/ ."""

from fastapi import FastAPI, Body
from fastapi.staticfiles import StaticFiles
from docx import Document
from docx.enum.style import WD\_STYLE\_TYPE
from docx.table import \_Cell
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from uuid import uuid4
import os, re

app = FastAPI()
BASE\_DIR = "generated"
SUB\_DIR = "downloads"
OUTPUT\_DIR = os.path.join(BASE\_DIR, SUB\_DIR)

os.makedirs(OUTPUT\_DIR, exist\_ok=True)
app.mount("/static", StaticFiles(directory=BASE\_DIR), name="static")

# ---------------------------------------------------------------------------

# Regex helpers

# ---------------------------------------------------------------------------

\_hdr1 = re.compile(r"^# (.+)")
\_hdr2 = re.compile(r"^## (.+)")
\_bullet = re.compile(r"^\[-\*] (.+)")
\_table\_row = re.compile(r"^|(.+)|\$")
\_table\_divider = re.compile(r"^\[|\s:-]+\$")
\_bold = re.compile(r"\*\*(.+?)\*\*")
\_italic = re.compile(r"\*(.+?)\*")
\_html\_strong = re.compile(r"<strong>(.+?)</strong>", re.IGNORECASE)
*invalid\_fname = re.compile(r"\[^A-Za-z0-9*.-]")

# ---------------------------------------------------------------------------

# Util functions

# ---------------------------------------------------------------------------

def \_set\_cell\_border(cell: \_Cell):
tc = cell.\_tc
tcPr = tc.get\_or\_add\_tcPr()
for edge in ("top", "left", "bottom", "right"):
element = OxmlElement(f"w:{edge}")
element.set(qn("w\:val"), "single")
element.set(qn("w\:sz"), "4")
element.set(qn("w\:space"), "0")
element.set(qn("w\:color"), "auto")
tcPr.append(element)

def safe\_filename(raw: str | None) -> str:
if not raw:
return f"{uuid4()}.docx"
name = *invalid\_fname.sub("*", raw).strip(".\_")
if not name:
name = str(uuid4())
if not name.lower().endswith(".docx"):
name += ".docx"
return name

# ---------------------------------------------------------------------------

# Markdown → DOCX conversion

# ---------------------------------------------------------------------------

def \_apply\_inline\_formats(paragraph, text: str):
"""Handle **bold** and *italic* inside a paragraph."""
pos = 0
while pos < len(text):
bold\_m = \_bold.search(text, pos)
italic\_m = \_italic.search(text, pos)
next\_m = None
if bold\_m and italic\_m:
next\_m = bold\_m if bold\_m.start() < italic\_m.start() else italic\_m
else:
next\_m = bold\_m or italic\_m
if not next\_m:
paragraph.add\_run(text\[pos:])
break
\# plain text before format
if next\_m.start() > pos:
paragraph.add\_run(text\[pos\:next\_m.start()])
run = paragraph.add\_run(next\_m.group(1))
if next\_m.re is \_bold:
run.bold = True
else:
run.italic = True
pos = next\_m.end()

def md\_to\_docx(md: str, path: str):
doc = Document()
if "List Bullet" not in \[s.name for s in doc.styles]:
doc.styles.add\_style("List Bullet", WD\_STYLE\_TYPE.PARAGRAPH).base\_style = doc.styles\["Normal"]

```
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
        text = _html_strong.sub(r"**\1**", line)
        p = doc.add_paragraph()
        _apply_inline_formats(p, text)
    i += 1
doc.save(path)
```

# ---------------------------------------------------------------------------

# Route: POST /docx

# ---------------------------------------------------------------------------

@app.post("/docx")
def make\_docx(payload: dict = Body(...)):
md = payload\["markdown"]
fname = safe\_filename(payload.get("filename"))
full = os.path.join(OUTPUT\_DIR, fname)
md\_to\_docx(md, full)
host = os.environ.get("RENDER\_EXTERNAL\_HOSTNAME", "localhost")
return {"download\_url": f"https\://{host}/static/{SUB\_DIR}/{fname}"}
