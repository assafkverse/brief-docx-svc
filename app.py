# Version 3.9b

"""FastAPI service: Markdown → DOCX with headings, bullets, bold/italic, and tables.
Files stored under /static/downloads/ ."""

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

# Regex patterns

\_hdr1 = re.compile(r"^# (.+)")
\_hdr2 = re.compile(r"^## (.+)")
\_bullet = re.compile(r"^\[-\*] (.+)")
\_table\_row = re.compile(r"^|(.+)|\$")
\_table\_divider = re.compile(r"^\[|\s:-]+\$")
\_bold = re.compile(r"\*\*(.+?)\*\*")
\_italic = re.compile(r"\*(.+?)\*")
\_html\_strong = re.compile(r"<strong>(.+?)</strong>", re.IGNORECASE)
*invalid\_fname = re.compile(r"\[^A-Za-z0-9*.-]")

# Helper functions

def \_set\_cell\_border(cell: \_Cell):
tcPr = cell.\_tc.get\_or\_add\_tcPr()
for edge in ("top", "left", "bottom", "right"):
ln = OxmlElement(f"w:{edge}")
ln.set(qn("w\:val"), "single")
ln.set(qn("w\:sz"), "4")
ln.set(qn("w\:space"), "0")
ln.set(qn("w\:color"), "auto")
tcPr.append(ln)

def *safe\_filename(name: str | None) -> str:
if not name:
return f"{uuid4()}.docx"
name = *invalid\_fname.sub("*", name).strip(".*") or str(uuid4())
if not name.lower().endswith(".docx"):
name += ".docx"
return name

def \_inline\_formats(par, text: str):
pos = 0
while pos < len(text):
m\_b = \_bold.search(text, pos)
m\_i = \_italic.search(text, pos)
m = None
if m\_b and m\_i:
m = m\_b if m\_b.start() < m\_i.start() else m\_i
else:
m = m\_b or m\_i
if not m:
par.add\_run(text\[pos:])
break
if m.start() > pos:
par.add\_run(text\[pos\:m.start()])
run = par.add\_run(m.group(1))
if m.re is \_bold:
run.bold = True
else:
run.italic = True
pos = m.end()

def md\_to\_docx(md: str, path: str):
doc = Document()
if "List Bullet" not in \[s.name for s in doc.styles]:
doc.styles.add\_style("List Bullet", WD\_STYLE\_TYPE.PARAGRAPH).base\_style = doc.styles\["Normal"]

```
lines = md.splitlines()
i = 0
while i < len(lines):
    line = lines[i]
    # tables
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
```

app = FastAPI()

@app.post("/docx")
def make\_docx(payload: dict = Body(...)):
md = payload\["markdown"]
fname = \_safe\_filename(payload.get("filename"))
full = os.path.join(OUTPUT\_DIR, fname)
md\_to\_docx(md, full)
host = os.environ.get("RENDER\_EXTERNAL\_HOSTNAME", "localhost")
return {"download\_url": f"https\://{host}/static/{SUB\_DIR}/{fname}"}
