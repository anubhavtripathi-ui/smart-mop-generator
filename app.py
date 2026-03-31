import io
import re
from datetime import datetime
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Smart MOP Generator", layout="wide")

TEMPLATE_PATH = "templates/Template.docx"

SECTIONS = [
    "Objective","Activity Description","Activity Type","Domain in Scope",
    "Pre-requisites","Inventory Details","Node Connectivity Process",
    "Identity & Access Management","Activity Triggering Method",
    "Standard Operating Procedure","Acceptance Criteria","Assumptions"
]

# -------- CLEAN ----------
def clean(text):
    return re.sub(r'^\d+[\.\)]\s*', '', text.strip().lower())

# -------- ACTIVITY ----------
def get_activity(doc):
    for p in doc.paragraphs[:20]:
        if "mop:" in p.text.lower():
            return p.text.split(":")[1].strip()
    return "Activity"

# -------- PARSE ----------
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

    # fallback
    for k in data:
        if not data[k]:
            data[k] = ["N/A"]

    return data

# -------- BUILD ----------
def build(template_bytes, solution_doc):

    doc = Document(io.BytesIO(template_bytes))

    # ===== FIX DATE =====
    today = datetime.today().strftime("%d %B %Y")
    for sec in doc.sections:
        for p in sec.header.paragraphs:
            if "{{current date}}" in p.text:
                p.text = p.text.replace("{{current date}}", today)

    # ===== ACTIVITY NAME =====
    activity = get_activity(solution_doc)

    for p in doc.paragraphs:
        if "activity name" in p.text.lower():
            p.text = activity
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ===== PARSE DATA =====
    data = parse(solution_doc)

    # ===== STRICT 1:1 MAPPING =====
    for i, p in enumerate(doc.paragraphs):
        heading = clean(p.text)

        for sec in SECTIONS:
            if heading == sec.lower():

                insert_index = i + 1
                content = data[sec]

                # clear existing placeholder safely
                if insert_index < len(doc.paragraphs):
                    if "sample" in doc.paragraphs[insert_index].text.lower():
                        doc.paragraphs[insert_index].text = ""

                # insert content
                for line in content:
                    doc.paragraphs[insert_index].insert_paragraph_before(line)
                    insert_index += 1

    # ===== SAVE =====
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer.getvalue(), activity

# -------- UI ----------
st.title("🚀 Smart MOP Generator")
st.info("🔒 No data stored. Processing is fully in-memory.")

uploaded = st.file_uploader("Upload Solution Document", type=["docx"])

if st.button("Generate MOP"):

    if not uploaded:
        st.warning("Upload file first")
        st.stop()

    solution_doc = Document(io.BytesIO(uploaded.read()))
    template_bytes = open(TEMPLATE_PATH, "rb").read()

    output, activity = build(template_bytes, solution_doc)

    st.success("MOP Generated Successfully")

    st.download_button(
        "Download MOP",
        data=output,
        file_name=f"{activity}.docx"
    )
