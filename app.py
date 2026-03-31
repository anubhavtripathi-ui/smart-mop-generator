import io
import re
from datetime import datetime
from pathlib import Path

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ================= CONFIG =================
st.set_page_config(page_title="Smart MOP Generator", layout="centered")

TEMPLATES_DIR = Path("templates")
TEMPLATES_DIR.mkdir(exist_ok=True)

# ================= HEADINGS =================
SECTION_LABELS = {
    "objective": "1. Objective",
    "activity_description": "2. Activity Description",
    "activity_type": "3. Activity Type",
    "domain_in_scope": "4. Domain in Scope",
    "prerequisites": "5. Pre-requisites",
    "inventory_details": "6. Inventory Details",
    "node_connectivity": "7. Node Connectivity Process",
    "iam": "8. Identity & Access Management",
    "triggering_method": "9. Activity Triggering Method",
    "sop": "10. Standard Operating Procedure",
    "acceptance_criteria": "11. Acceptance Criteria",
    "assumptions": "12. Assumptions",
}

SECTION_KEYS = list(SECTION_LABELS.keys())

# ================= ACTIVITY NAME =================
def extract_activity_name(doc):
    for para in doc.paragraphs[:10]:
        text = para.text.strip()
        if text.lower().startswith("mop:"):
            return text.replace("MOP:", "").strip()
    return "Activity Name"

# ================= SECTION PARSER =================
def extract_sections(doc):
    sections = {k: [] for k in SECTION_KEYS}
    current_section = None

    for para in doc.paragraphs:
        text = para.text.strip()

        if not text:
            continue

        # detect heading
        detected = None
        for key, label in SECTION_LABELS.items():
            label_text = label.split(". ", 1)[1].lower()

            if text.lower().startswith(label_text):
                detected = key
                break

        if detected:
            current_section = detected
            continue

        # store content
        if current_section:
            if not text.lower().startswith("mop"):
                sections[current_section].append(text)

    return sections

# ================= CORE ENGINE =================
def build_mop(template_bytes, activity_name, sections):

    doc = Document(io.BytesIO(template_bytes))

    # -------- HEADER DATE --------
    today = datetime.today().strftime("%d %B %Y")

    for section in doc.sections:
        for para in section.header.paragraphs:
            if "{{current date}}" in para.text:
                para.text = para.text.replace("{{current date}}", today)

    # -------- ACTIVITY NAME --------
    for para in doc.paragraphs:
        if "Actvity Name" in para.text or "Activity Name" in para.text:
            para.text = activity_name
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # -------- FIND HEADINGS --------
    heading_positions = {}

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()

        for key, label in SECTION_LABELS.items():
            if text.startswith(label):
                heading_positions[key] = i

    # -------- INSERT CONTENT --------
    for key in SECTION_KEYS:

        if key not in heading_positions:
            continue

        start_index = heading_positions[key] + 1
        content = sections.get(key, [])

        # clear placeholder
        next_para = doc.paragraphs[start_index]
        if "sample" in next_para.text.lower():
            next_para.text = ""

        insert_index = start_index

        for line in content:
            doc.paragraphs[insert_index].insert_paragraph_before(line)
            insert_index += 1

    # -------- CLEAN EXTRA TEXT --------
    for para in doc.paragraphs:
        if "sample" in para.text.lower():
            para.text = ""

    # -------- SAVE --------
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer.read()

# ================= UI =================
st.title("📋 Smart MOP Generator")

templates = list(TEMPLATES_DIR.glob("*.docx"))

if not templates:
    st.warning("Put Template.docx inside /templates folder")
    st.stop()

template_file = st.selectbox("Select Template", templates)
template_bytes = open(template_file, "rb").read()

sol_file = st.file_uploader("Upload Solution Document", type=["docx"])

if st.button("Generate MOP"):

    if not sol_file:
        st.warning("Upload file first")
        st.stop()

    sol_doc = Document(io.BytesIO(sol_file.read()))

    activity_name = extract_activity_name(sol_doc)
    sections = extract_sections(sol_doc)

    output = build_mop(template_bytes, activity_name, sections)

    st.success("✅ MOP Generated Successfully")

    st.download_button(
        "Download MOP",
        data=output,
        file_name=f"MOP_{activity_name}.docx"
    )
