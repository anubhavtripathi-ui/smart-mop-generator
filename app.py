import io
import re
from datetime import datetime
from pathlib import Path

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ================= UI =================
st.set_page_config(page_title="Smart MOP Generator", layout="wide")

st.markdown("""
<style>
body {font-family: 'Segoe UI';}
.main-box {
    background: linear-gradient(135deg,#1e293b,#0f172a);
    padding:20px;border-radius:20px;color:white;margin-bottom:20px;
}
.card {background:white;padding:20px;border-radius:15px;box-shadow:0 5px 15px rgba(0,0,0,0.1);}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-box"><h1>🚀 Smart MOP Generator</h1><p>Secure | No Data Stored | In-Memory Processing</p></div>', unsafe_allow_html=True)

st.info("🔒 Privacy: Your file is NOT stored anywhere. Everything runs in-memory.")

# ================= CONFIG =================
TEMPLATE_PATH = "templates/Template.docx"

SECTION_KEYS = [
    "objective","activity description","activity type","domain in scope",
    "pre-requisites","inventory details","node connectivity process",
    "identity & access management","activity triggering method",
    "standard operating procedure","acceptance criteria","assumptions"
]

# ================= UTILS =================
def clean(text):
    return re.sub(r'^\d+[\.\)]\s*', '', text.strip().lower())

def extract_activity_name(doc):
    for p in doc.paragraphs[:15]:
        if "mop:" in p.text.lower():
            return p.text.split(":")[1].strip()
    return "Generated_MOP"

# ================= PARSER =================
def parse_sections(doc):
    data = {k: [] for k in SECTION_KEYS}
    current = None

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue

        t = clean(text)

        # detect heading
        for key in SECTION_KEYS:
            if t.startswith(key):
                current = key
                break
        else:
            if current:
                data[current].append(text)

    # fallback: ensure no blank
    for k in data:
        if not data[k]:
            data[k] = ["N/A"]

    return data

# ================= BUILDER =================
def build_mop(template_bytes, activity, sections):
    doc = Document(io.BytesIO(template_bytes))

    # Header Date
    today = datetime.today().strftime("%d %B %Y")
    for sec in doc.sections:
        for p in sec.header.paragraphs:
            if "{{date}}" in p.text.lower():
                p.text = today

    # Activity Name
    for p in doc.paragraphs:
        if "activity name" in p.text.lower():
            p.text = activity
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Map sections
    for i, p in enumerate(doc.paragraphs):
        text = clean(p.text)

        for key in SECTION_KEYS:
            if text.startswith(key):

                insert_index = i + 1
                content = sections[key]

                # clear placeholder
                if insert_index < len(doc.paragraphs):
                    doc.paragraphs[insert_index].text = ""

                for line in content:
                    doc.paragraphs[insert_index].insert_paragraph_before(line)
                    insert_index += 1

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# ================= UI =================
uploaded = st.file_uploader("📂 Upload Solution Document", type=["docx"])

if st.button("⚡ Generate MOP"):
    if not uploaded:
        st.warning("Upload file first")
        st.stop()

    sol_doc = Document(io.BytesIO(uploaded.read()))
    template_bytes = open(TEMPLATE_PATH, "rb").read()

    activity = extract_activity_name(sol_doc)
    sections = parse_sections(sol_doc)

    output = build_mop(template_bytes, activity, sections)

    st.success("✅ MOP Generated Successfully")

    st.download_button(
        "⬇ Download MOP",
        data=output,
        file_name=f"{activity}.docx"
    )
