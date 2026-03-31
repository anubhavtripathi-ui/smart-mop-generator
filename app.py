import io
import re
from datetime import datetime
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

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
def insert_after(paragraph, text):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    new_para.text = text
    return new_para

# ---------- TEXT REPLACE ----------
def replace_text_preserve_format(paragraph, new_text):
    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for i in range(1, len(paragraph.runs)):
            paragraph.runs[i].text = ""
    else:
        paragraph.text = new_text

# ---------- BUILD ----------
def build_mop(template_bytes, solution_doc):

    doc = Document(io.BytesIO(template_bytes))

    # ===== DATE UPDATE (HEADER ONLY) =====
    today = datetime.today().strftime("%d %B %Y")
    for sec in doc.sections:
        for p in sec.header.paragraphs:
            if "{{current date}}" in p.text.lower():
                replace_text_preserve_format(
                    p,
                    p.text.replace("{{current date}}", today)
                )

    # ===== ACTIVITY NAME =====
    activity = extract_activity(solution_doc)

    for p in doc.paragraphs:
        if "activity name" in normalize(p.text):
            replace_text_preserve_format(p, activity)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ===== PARSE =====
    data = parse_sections(solution_doc)

    # ===== INSERT CONTENT =====
    for i, p in enumerate(doc.paragraphs):
        heading = normalize(p.text)

        for sec in SECTIONS:
            if sec in heading:   # IMPORTANT FIX

                ref_para = p

                # remove "sample"
                next_index = i + 1
                if next_index < len(doc.paragraphs):
                    if "sample" in doc.paragraphs[next_index].text.lower():
                        doc.paragraphs[next_index].text = ""

                # insert actual content
                for line in data[sec]:
                    ref_para = insert_after(ref_para, line)

    # ===== SAVE =====
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer.getvalue(), activity

# ---------- UI ----------
st.title("🚀 Smart MOP Generator")

st.markdown("### 📄 Template Preview")
st.markdown("👉 Upload solution doc → MOP will be generated based on fixed template")

st.info("🔒 Privacy: No data is stored. Everything runs in-memory only.")

uploaded = st.file_uploader("📤 Upload Solution Document (.docx)", type=["docx"])

if st.button("⚡ Generate MOP"):

    if not uploaded:
        st.warning("⚠️ Please upload a document first")
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
