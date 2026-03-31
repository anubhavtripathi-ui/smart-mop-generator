import io
import re
from datetime import datetime
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ================= HEADINGS =================
SECTIONS = [
    "Objective","Activity Description","Activity Type","Domain in Scope",
    "Pre-requisites","Inventory Details","Node Connectivity Process",
    "Identity & Access Management","Activity Triggering Method",
    "Standard Operating Procedure","Acceptance Criteria","Assumptions"
]

# ================= CLEAN =================
def clean(text):
    return re.sub(r'^\d+[\.\)]\s*', '', text.strip().lower())

# ================= ACTIVITY NAME =================
def get_activity(doc):
    for p in doc.paragraphs[:15]:
        if "mop:" in p.text.lower():
            return p.text.split(":")[1].strip()
    return "Activity"

# ================= PARSE =================
def parse(doc):
    data = {k: [] for k in SECTIONS}
    current = None

    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue

        t = clean(txt)

        for sec in SECTIONS:
            if t.startswith(sec.lower()):
                current = sec
                break
        else:
            if current:
                data[current].append(txt)

    return data

# ================= BUILD =================
def build(template_path, solution_doc):

    doc = Document(template_path)

    # ---- DATE FIX ----
    today = datetime.today().strftime("%d %B %Y")
    for sec in doc.sections:
        for p in sec.header.paragraphs:
            if "{{current date}}" in p.text:
                p.text = p.text.replace("{{current date}}", today)

    activity = get_activity(solution_doc)

    # ---- ACTIVITY NAME ----
    for p in doc.paragraphs:
        if "activity name" in p.text.lower():
            p.text = activity
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ---- TOC BUILD ----
    toc_index = None
    for i,p in enumerate(doc.paragraphs):
        if "contents" in p.text.lower():
            toc_index = i
            break

    if toc_index:
        for i,sec in enumerate(SECTIONS,1):
            line = f"{i}. {sec}"
            doc.paragraphs[toc_index+1].insert_paragraph_before(line)

    # ---- PAGE BREAK ----
    doc.add_page_break()

    # ---- CONTENT INSERT ----
    sections = parse(solution_doc)

    for sec in SECTIONS:
        doc.add_paragraph(f"{sec}", style='Heading 1')

        for line in sections[sec]:
            doc.add_paragraph(line)

    # ---- SAVE ----
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)

    return out, activity
