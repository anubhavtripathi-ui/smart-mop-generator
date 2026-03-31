import io
import re
from datetime import datetime
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.shared import Pt

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
        if "mop:" in p.text.lower():
            return p.text.split(":")[1].strip()
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

    for k in data:
        if not data[k]:
            data[k] = ["N/A"]

    return data

# ---------- CLEAR PAGE 2 CONTENT ----------
def clear_template_content(doc):
    start = False
    for p in doc.paragraphs:
        txt = normalize(p.text)

        if "objective" in txt:
            start = True

        if start:
            p.text = ""   # clear everything from page 2 onward

# ---------- INSERT PARA ----------
def insert_para_after(paragraph, text, bold=False):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    run = new_para.add_run(text)
    run.bold = bold
    return new_para

# ---------- PAGE BREAK ----------
def add_page_break(paragraph):
    run = paragraph.add_run()
    run.add_break(1)

# ---------- BUILD ----------
def build_mop(template_bytes, solution_doc):

    doc = Document(io.BytesIO(template_bytes))

    # ===== DATE UPDATE =====
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

    # ===== REMOVE TEMPLATE BODY =====
    clear_template_content(doc)

    # ===== FIND TOC END (Page 1 end) =====
    last_para = doc.paragraphs[-1]

    # ===== FORCE PAGE BREAK =====
    add_page_break(last_para)

    # ===== PARSE DATA =====
    data = parse_sections(solution_doc)

    # ===== INSERT PAGE 2 CONTENT =====
    ref = last_para

    for idx, sec in enumerate(SECTIONS, start=1):

        # heading (with numbering)
        heading_text = f"{idx}. {sec.title()}"

        ref = insert_para_after(ref, heading_text, bold=True)

        # content
        for line in data[sec]:
            ref = insert_para_after(ref, line)

    # ===== SAVE =====
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer.getvalue(), activity

# ---------- UI ----------
st.title("🚀 Smart MOP Generator")

st.markdown("### 📄 Fixed Template Based MOP Generation")

st.info("🔒 Privacy Guaranteed: No data stored. Runs fully in-memory.")

uploaded = st.file_uploader("Upload Solution Document (.docx)", type=["docx"])

if st.button("Generate MOP"):

    if not uploaded:
        st.warning("Please upload a file")
        st.stop()

    solution_doc = Document(io.BytesIO(uploaded.read()))
    template_bytes = open(TEMPLATE_PATH, "rb").read()

    output, activity = build_mop(template_bytes, solution_doc)

    st.success("MOP Generated Successfully")

    st.download_button(
        "Download MOP",
        data=output,
        file_name=f"{activity}.docx"
    )
