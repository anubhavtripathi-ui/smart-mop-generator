import io
import re
from datetime import datetime
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn   # ✅ NEW

st.set_page_config(page_title="Smart MOP Generator", layout="wide")

TEMPLATE_PATH = "templates/Template.docx"

SECTIONS = [
    "objective","activity description","activity type","domain in scope",
    "pre-requisites","inventory details","node connectivity process",
    "identity & access management","activity triggering method",
    "standard operating procedure","acceptance criteria","assumptions"
]

# ---------- NORMALIZE ----------
def normalize(text):
    text = text.lower().strip()
    text = re.sub(r'^\d+[\.\)]\s*', '', text)
    return text

# ---------- ACTIVITY NAME ----------
def extract_activity(doc):
    for p in doc.paragraphs[:20]:
        txt = p.text.strip()
        if "mop:" in txt.lower():
            return txt.split(":")[1].strip()
    return doc.paragraphs[0].text.strip()

# ---------- PARSE ----------
def parse_sections(doc):
    data = {k: [] for k in SECTIONS}
    current = None

    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt:
            continue

        n = normalize(txt)

        for sec in SECTIONS:
            if sec in n:
                current = sec
                break
        else:
            if current:
                data[current].append(txt)

    # fallback
    for k in data:
        if not data[k]:
            data[k] = ["N/A"]

    return data

# ---------- SAFE INSERT ----------
def insert_after(paragraph, text, bold=False):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    run = new_para.add_run(text)
    run.bold = bold
    return new_para

# ---------- PAGE BREAK ----------
def add_page_break(paragraph):
    run = paragraph.add_run()
    run.add_break(WD_BREAK.PAGE)

# ---------- FIND TOC ----------
def find_toc_anchor(doc):
    for p in doc.paragraphs:
        txt = normalize(p.text)
        if "contents" in txt or "table of content" in txt:
            return p
    return doc.paragraphs[-1]

# ---------- CLEAR PAGE 2 ----------
def clear_template_content(doc):
    start = False
    for p in doc.paragraphs:
        txt = normalize(p.text)
        if "objective" in txt:
            start = True
        if start:
            p.text = ""

# ---------- ADD TOC FIELD (NEW) ----------
def add_toc(paragraph):
    run = paragraph.add_run()

    fldChar_begin = OxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')

    instrText = OxmlElement('w:instrText')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'

    fldChar_separate = OxmlElement('w:fldChar')
    fldChar_separate.set(qn('w:fldCharType'), 'separate')

    fldChar_end = OxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')

    run._r.append(fldChar_begin)
    run._r.append(instrText)
    run._r.append(fldChar_separate)
    run._r.append(fldChar_end)

# ---------- BUILD ----------
def build_mop(template_bytes, solution_doc):

    doc = Document(io.BytesIO(template_bytes))

    # ===== HEADER DATE UPDATE =====
    today = datetime.today().strftime("%d %B %Y")
    for sec in doc.sections:
        for p in sec.header.paragraphs:
            if "{{current date}}" in p.text.lower():
                p.text = p.text.replace("{{current date}}", today)

    # ===== ACTIVITY NAME =====
    activity = extract_activity(solution_doc)

    for p in doc.paragraphs:
        if "activity name" in normalize(p.text):
            p.text = activity
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ===== CLEAR TEMPLATE PAGE 2 =====
    clear_template_content(doc)

    # ===== FIND PAGE 1 END =====
    toc_para = find_toc_anchor(doc)

    # ✅ ADD TOC HERE (NEW)
    add_toc(toc_para)

    # ===== PAGE BREAK =====
    add_page_break(toc_para)

    # ===== PARSE DATA =====
    data = parse_sections(solution_doc)

    # ===== INSERT PAGE 2 CONTENT =====
    ref = toc_para

    for idx, sec in enumerate(SECTIONS, start=1):

        heading_text = f"{idx}. {sec.title()}"
        ref = insert_after(ref, heading_text, bold=True)

        # ✅ MAKE HEADING DETECTABLE FOR TOC (NEW)
        ref.style = "Heading 1"

        for line in data[sec]:
            ref = insert_after(ref, line)

    # ===== SAVE =====
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer.getvalue(), activity

# ---------- UI ----------
st.title("🚀 Smart MOP Generator")

st.markdown("### 📄 Template-based MOP Generator")
st.markdown("✔ Page 1 → Method of Procedure + Activity Name + TOC")
st.markdown("✔ Page 2 → Auto content mapping (12 sections)")

st.info("🔒 Privacy Notice: No data is stored. Processing is fully in-memory.")

uploaded = st.file_uploader("📤 Upload Solution Document (.docx)", type=["docx"])

if st.button("⚡ Generate MOP"):

    if not uploaded:
        st.warning("Please upload a document first")
        st.stop()

    solution_doc = Document(io.BytesIO(uploaded.read()))
    template_bytes = open(TEMPLATE_PATH, "rb").read()

    output, activity = build_mop(template_bytes, solution_doc)

    st.success("✅ MOP Generated Successfully")

    st.download_button(
        "⬇️ Download MOP",
        data=output,
        file_name=f"{activity}.docx"
    )
