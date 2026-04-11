"""
Smart MOP Generator — Premium Edition
=======================================
Upload Solution Document → MOP generated in exact Template format.
Template in templates/ folder OR root directory.
No data stored. All processing in-memory.
"""

import io
import os
import re
import time
from datetime import datetime
from pathlib import Path
from copy import deepcopy

import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ─────────────────────────────────────────────────────────────────
# PAGE CONFIG  — must be FIRST streamlit call
# ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Smart MOP Generator",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────────────────────
# CSS — Premium Redesign
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700&family=IBM+Plex+Sans:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

/* ── Reset & Base ── */
html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
    background-color: #0d0f14;
    color: #c9d1d9;
}
.block-container {
    max-width: 860px;
    padding-top: 0 !important;
    padding-bottom: 2rem;
}

/* ── Top accent bar ── */
.top-bar {
    width: 100%;
    height: 3px;
    background: linear-gradient(90deg, #1a6b4a 0%, #2d9cdb 50%, #8b5cf6 100%);
    margin-bottom: 0;
}

/* ── Hero ── */
.hero {
    background: linear-gradient(160deg, #111827 0%, #0d1f2d 60%, #0d1a12 100%);
    border: 1px solid rgba(45,156,219,.12);
    border-top: none;
    border-radius: 0 0 20px 20px;
    padding: 2.6rem 2.6rem 2rem;
    margin-bottom: 2rem;
    position: relative;
    overflow: hidden;
}
.hero::before {
    content: '';
    position: absolute;
    top: -60px; right: -60px;
    width: 280px; height: 280px;
    background: radial-gradient(circle, rgba(45,156,219,.06) 0%, transparent 70%);
    pointer-events: none;
}
.hero::after {
    content: '';
    position: absolute;
    bottom: -40px; left: -40px;
    width: 200px; height: 200px;
    background: radial-gradient(circle, rgba(26,107,74,.08) 0%, transparent 70%);
    pointer-events: none;
}
.hero-eyebrow {
    font-family: 'IBM Plex Mono', monospace;
    font-size: .65rem;
    letter-spacing: 3px;
    color: #2d9cdb;
    text-transform: uppercase;
    margin-bottom: .6rem;
}
.hero-title {
    font-family: 'Playfair Display', serif;
    font-size: 2.2rem;
    font-weight: 700;
    color: #e6edf3;
    margin: 0 0 .3rem;
    line-height: 1.2;
}
.hero-title em {
    font-style: italic;
    color: #2d9cdb;
}
.hero-sub {
    font-size: .85rem;
    color: #6e7681;
    margin: 0 0 1.4rem;
    font-weight: 300;
    letter-spacing: .3px;
}
.badge-row {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
}
.badge {
    font-family: 'IBM Plex Mono', monospace;
    font-size: .62rem;
    font-weight: 500;
    padding: 4px 10px;
    border-radius: 4px;
    letter-spacing: .5px;
}
.b-green  { background: rgba(26,107,74,.18);  color: #3fb882; border: 1px solid rgba(63,184,130,.2); }
.b-blue   { background: rgba(45,156,219,.14); color: #5ba8e0; border: 1px solid rgba(45,156,219,.25); }
.b-purple { background: rgba(139,92,246,.12); color: #a78bfa; border: 1px solid rgba(139,92,246,.22); }

/* ── Privacy notice ── */
.priv-bar {
    display: flex;
    align-items: center;
    gap: 10px;
    background: rgba(26,107,74,.06);
    border-left: 2px solid #1a6b4a;
    border-radius: 0 8px 8px 0;
    padding: .6rem 1rem;
    font-size: .76rem;
    color: #6e7681;
    margin-bottom: 1.8rem;
}
.priv-bar strong { color: #3fb882; }

/* ── Step cards ── */
.step-card {
    background: #111827;
    border: 1px solid rgba(255,255,255,.06);
    border-radius: 14px;
    padding: 1.4rem 1.6rem;
    margin-bottom: 1.1rem;
    position: relative;
    transition: border-color .2s;
}
.step-card:hover { border-color: rgba(45,156,219,.2); }
.step-header {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-bottom: 1rem;
}
.step-number {
    font-family: 'IBM Plex Mono', monospace;
    font-size: .62rem;
    font-weight: 500;
    color: #2d9cdb;
    background: rgba(45,156,219,.1);
    border: 1px solid rgba(45,156,219,.2);
    border-radius: 4px;
    padding: 3px 8px;
    letter-spacing: 1px;
}
.step-title {
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: .78rem;
    font-weight: 600;
    color: #8b949e;
    letter-spacing: 1.2px;
    text-transform: uppercase;
}

/* ── Status pills ── */
.pill-ok {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background: rgba(26,107,74,.1);
    border: 1px solid rgba(63,184,130,.18);
    border-radius: 6px;
    padding: 5px 12px;
    font-size: .75rem;
    color: #3fb882;
    margin-top: 4px;
}
.pill-warn {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background: rgba(210,105,30,.1);
    border: 1px solid rgba(210,105,30,.2);
    border-radius: 6px;
    padding: 5px 12px;
    font-size: .75rem;
    color: #e8955a;
    margin-top: 4px;
}

/* ── Primary button ── */
.stButton > button {
    background: linear-gradient(135deg, #1a6b4a 0%, #1d5f8a 100%) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 10px !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-weight: 600 !important;
    font-size: .9rem !important;
    padding: .65rem 2rem !important;
    width: 100% !important;
    letter-spacing: .4px !important;
    transition: all .2s !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #1e7d56 0%, #2270a3 100%) !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 8px 24px rgba(45,156,219,.15) !important;
}
.stButton > button:disabled {
    opacity: .35 !important;
    transform: none !important;
}

/* ── Download button ── */
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #1a5c3a 0%, #164e70 100%) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 10px !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-weight: 600 !important;
    font-size: .9rem !important;
    padding: .65rem 2rem !important;
    width: 100% !important;
}

/* ── Progress steps ── */
.prog-wrap {
    background: #0d1117;
    border: 1px solid rgba(255,255,255,.05);
    border-radius: 10px;
    padding: 1rem 1.2rem;
}
.ps {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 7px 0;
    font-size: .78rem;
    border-bottom: 1px solid rgba(255,255,255,.03);
    transition: color .3s;
}
.ps:last-child { border-bottom: none; }
.ps.done  { color: #3fb882; }
.ps.doing { color: #2d9cdb; }
.ps.wait  { color: #30363d; }
.pd {
    width: 6px; height: 6px;
    border-radius: 50%;
    flex-shrink: 0;
}
.pd.done  { background: #3fb882; }
.pd.doing { background: #2d9cdb; animation: pulse 1.1s infinite; }
.pd.wait  { background: #21262d; }

@keyframes pulse {
    0%, 100% { opacity: 1; transform: scale(1); }
    50% { opacity: .4; transform: scale(.7); }
}

/* ── Success card ── */
.success-wrap {
    background: linear-gradient(135deg, rgba(26,107,74,.1) 0%, rgba(45,156,219,.07) 100%);
    border: 1px solid rgba(63,184,130,.2);
    border-radius: 14px;
    padding: 1.6rem;
    margin: 1rem 0;
    text-align: center;
}
.success-icon { font-size: 2rem; margin-bottom: .4rem; }
.success-title {
    font-family: 'Playfair Display', serif;
    font-size: 1.15rem;
    color: #3fb882;
    margin-bottom: .2rem;
}
.success-sub { font-size: .78rem; color: #6e7681; }
.success-name { color: #5ba8e0; font-weight: 600; }

/* ── Summary metrics ── */
.metric-row {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 10px;
    margin-top: 1rem;
}
.metric-box {
    background: #0d1117;
    border: 1px solid rgba(255,255,255,.06);
    border-radius: 10px;
    padding: .9rem;
    text-align: center;
}
.metric-val {
    font-family: 'Playfair Display', serif;
    font-size: 1.6rem;
    color: #e6edf3;
    line-height: 1;
}
.metric-lbl {
    font-size: .68rem;
    color: #6e7681;
    margin-top: 4px;
    letter-spacing: .5px;
    text-transform: uppercase;
}

/* ── Divider ── */
hr { border-color: rgba(255,255,255,.05) !important; }

/* ── Expander ── */
.streamlit-expanderHeader {
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-size: .78rem !important;
    color: #6e7681 !important;
}

/* ── Selectbox & file uploader labels ── */
label { color: #8b949e !important; font-size: .8rem !important; }

/* ── Footer ── */
.footer {
    text-align: center;
    font-size: .66rem;
    color: #21262d;
    padding: 1rem 0 .5rem;
    border-top: 1px solid rgba(255,255,255,.04);
    font-family: 'IBM Plex Mono', monospace;
    letter-spacing: .5px;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────
HEADING_MAP = {
    "objective":            ["objective"],
    "activity_description": ["activity description"],
    "activity_type":        ["activity type"],
    "domain_in_scope":      ["domain in scope", "domain"],
    "prerequisites":        ["pre-requisites", "prerequisites"],
    "inventory_details":    ["inventory details", "inventory"],
    "node_connectivity":    ["node connectivity", "node connectivity process"],
    "iam":                  ["identity & access", "identity and access", "identity"],
    "triggering_method":    ["activity triggering", "triggering method"],
    "sop":                  ["standard operating procedure", "sop"],
    "acceptance_criteria":  ["acceptance criteria"],
    "assumptions":          ["assumptions"],
    "connectivity_diagram": ["connectivity diagram"],
}

SECTION_KEYS = [
    "objective", "activity_description", "activity_type", "domain_in_scope",
    "prerequisites", "inventory_details", "node_connectivity", "iam",
    "triggering_method", "sop", "acceptance_criteria", "assumptions",
    "connectivity_diagram",
]

SECTION_LABELS = {
    "objective":            "1. Objective",
    "activity_description": "2. Activity Description",
    "activity_type":        "3. Activity Type",
    "domain_in_scope":      "4. Domain in Scope",
    "prerequisites":        "5. Pre-requisites",
    "inventory_details":    "6. Inventory Details",
    "node_connectivity":    "7. Node Connectivity Process",
    "iam":                  "8. Identity & Access Management",
    "triggering_method":    "9. Activity Triggering Method",
    "sop":                  "10. Standard Operating Procedure",
    "acceptance_criteria":  "11. Acceptance Criteria",
    "assumptions":          "12. Assumptions",
    "connectivity_diagram": "Connectivity Diagram:",
}

TOC_PAGES = {
    "objective": 2, "activity_description": 2, "activity_type": 2,
    "domain_in_scope": 2, "prerequisites": 2, "inventory_details": 3,
    "node_connectivity": 3, "iam": 3, "triggering_method": 3,
    "sop": 3, "acceptance_criteria": 4, "assumptions": 4,
}

PARA_SECTIONS     = {"objective", "activity_description", "activity_type",
                     "domain_in_scope", "inventory_details", "assumptions"}
BULLET_SECTIONS   = {"prerequisites", "node_connectivity", "iam",
                     "triggering_method", "acceptance_criteria"}
NUMBERED_SECTIONS = {"sop"}

# ─────────────────────────────────────────────────────────────────
# TEMPLATE DISCOVERY
# ─────────────────────────────────────────────────────────────────
def discover_templates() -> list[Path]:
    found = []
    tmpl_dir = Path("templates")
    if tmpl_dir.exists():
        found += sorted(tmpl_dir.glob("*.docx"))
    found += sorted(p for p in Path(".").glob("*.docx")
                    if p.name not in [t.name for t in found])
    return found


def load_template_bytes(path: Path) -> bytes:
    with open(path, "rb") as f:
        return f.read()


# ─────────────────────────────────────────────────────────────────
# SOLUTION DOC PARSER
# ─────────────────────────────────────────────────────────────────
def normalize_heading(text: str) -> str | None:
    t = re.sub(r'^\d+[\.\)]\s*', '', text).strip().lower()
    t = re.sub(r'\s+', ' ', t)
    for key, aliases in HEADING_MAP.items():
        for alias in aliases:
            if alias in t:
                return key
    return None


def extract_activity_name(doc: Document) -> str:
    for para in doc.paragraphs[:8]:
        if para.style.name.startswith("Heading 1"):
            name = para.text.strip()
            name = re.sub(r'^MOP\s*:\s*', '', name, flags=re.IGNORECASE)
            name = re.sub(r'^UC\s*:\s*', '', name, flags=re.IGNORECASE)
            if name:
                return name
    for para in doc.paragraphs[:10]:
        for run in para.runs:
            if run.italic and run.underline and para.text.strip():
                return para.text.strip()
    return "Activity Name"


def extract_sections(doc: Document) -> dict:
    sections   = {k: [] for k in SECTION_KEYS}
    sections["connectivity_diagram"] = []
    current_key = None
    image_rels  = {}

    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            try:
                ext = rel.target_part.content_type.split("/")[-1]
                if ext == "jpeg":
                    ext = "jpg"
                image_rels[rel.rId] = (rel.target_part.blob, ext)
            except Exception:
                pass

    _BLIP = "http://schemas.openxmlformats.org/drawingml/2006/main"
    _REL  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    for para in doc.paragraphs:
        style = para.style.name
        text  = para.text.strip()

        if style.startswith("Heading"):
            key = normalize_heading(text)
            if key:
                current_key = key
            continue

        if current_key is None:
            continue

        has_image = False
        for blip in para._p.findall(f".//{{{_BLIP}}}blip"):
            embed = blip.get(f"{{{_REL}}}embed")
            if embed and embed in image_rels:
                sections["connectivity_diagram"].append(image_rels[embed])
                has_image = True
        if has_image:
            continue

        if text in ("METHOD OF PROCEDURE", "CONTENTS:", "CONTENTS", ""):
            continue
        if re.match(r'^\d+\.\s+\w.*Page\s+\d+', text):
            continue
        if text == "sample...":
            continue

        if current_key in sections:
            clean = re.sub(r'^[-–•]\s*', '', text)
            clean = re.sub(r'^\d+[\.\)]\s*', '', clean)
            clean = clean.strip()
            if clean:
                sections[current_key].append(clean)

    return sections


# ─────────────────────────────────────────────────────────────────
# DOCX BUILDER HELPERS
# ─────────────────────────────────────────────────────────────────
def _r(para, text, font="Calibri", size=None, bold=False,
       italic=False, underline=False, color=None):
    run = para.add_run(text)
    run.font.name      = font
    if size:  run.font.size      = Pt(size)
    run.font.bold      = bold
    run.font.italic    = italic
    run.font.underline = underline
    if color: run.font.color.rgb = RGBColor(*color)
    return run


def _set_rtab(para, pos=8400):
    pPr = para._p.get_or_add_pPr()
    for old in pPr.findall(qn("w:tabs")):
        pPr.remove(old)
    tabs = OxmlElement("w:tabs")
    tab  = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:pos"), str(pos))
    tabs.append(tab)
    pPr.append(tabs)


def _pgbr(doc):
    p   = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    run = p.add_run()
    br  = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._r.append(br)


def _h2(doc, text):
    p = doc.add_paragraph()
    p.style = doc.styles["Heading 2"]
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(2)
    run = p.add_run(text)
    run.font.name  = "Calibri"
    run.font.size  = Pt(13)
    run.font.bold  = True
    run.font.color.rgb = RGBColor(0x4F, 0x81, 0xBD)


def _body(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    _r(p, text, size=11)


def _bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    _r(p, text, size=11)


def _numbered(doc, text):
    p = doc.add_paragraph(style="List Number")
    p.paragraph_format.space_after = Pt(3)
    _r(p, text, size=11)


def _img(doc, img_bytes, ext):
    p   = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    run = p.add_run()
    run.add_picture(io.BytesIO(img_bytes), width=Inches(5))


# ─────────────────────────────────────────────────────────────────
# HEADER DATE UPDATER
# ─────────────────────────────────────────────────────────────────
def _update_header_date(doc: Document, today_str: str):
    for section in doc.sections:
        for para in section.header.paragraphs:
            for run in para.runs:
                if "{{current date}}" in run.text:
                    run.text = run.text.replace("{{current date}}", today_str)


# ─────────────────────────────────────────────────────────────────
# BODY CLEAR — preserve headerReference + footerReference
# ─────────────────────────────────────────────────────────────────
def _clear_and_prep_body(doc: Document):
    body = doc.element.body

    header_refs = []
    footer_refs = []
    for child in body:
        for nested_sectPr in child.findall(".//" + qn("w:sectPr")):
            for elem in nested_sectPr:
                if elem.tag == qn("w:headerReference"):
                    header_refs.append(deepcopy(elem))
                elif elem.tag == qn("w:footerReference"):
                    footer_refs.append(deepcopy(elem))

    body_sectPr = body.find(qn("w:sectPr"))

    for child in list(body):
        body.remove(child)

    if body_sectPr is None:
        body_sectPr = OxmlElement("w:sectPr")

    for elem in list(body_sectPr):
        if elem.tag in (qn("w:headerReference"), qn("w:footerReference")):
            body_sectPr.remove(elem)

    insert_pos = 0
    for ref in header_refs + footer_refs:
        body_sectPr.insert(insert_pos, ref)
        insert_pos += 1

    body.append(body_sectPr)


# ─────────────────────────────────────────────────────────────────
# MAIN MOP BUILDER
# ─────────────────────────────────────────────────────────────────
def build_mop(template_bytes: bytes, activity_name: str,
              sections: dict, today_str: str) -> bytes:
    doc = Document(io.BytesIO(template_bytes))
    _update_header_date(doc, today_str)
    _clear_and_prep_body(doc)

    # Cover page
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    _r(p, "METHOD OF PROCEDURE", size=18, bold=True, color=(0x7F, 0x7F, 0x7F))

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    _r(p, activity_name, size=14, italic=True, underline=True)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    _r(p, "CONTENTS:", size=12, bold=True, underline=True)

    for key in SECTION_KEYS[:-1]:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        _set_rtab(p, 8400)
        _r(p, SECTION_LABELS[key], size=11)
        _r(p, f"\tPage {TOC_PAGES.get(key, 2)}", size=11)

    _pgbr(doc)

    # Body sections
    for key in SECTION_KEYS[:-1]:
        content = sections.get(key, [])
        _h2(doc, SECTION_LABELS[key])

        if key in PARA_SECTIONS:
            _body(doc, " ".join(content).strip() if content else "")
        elif key in BULLET_SECTIONS:
            if content:
                for item in content:
                    _bullet(doc, item)
            else:
                _body(doc, "")
        elif key in NUMBERED_SECTIONS:
            if content:
                for item in content:
                    _numbered(doc, item)
            else:
                _body(doc, "")

    images = sections.get("connectivity_diagram", [])
    if images:
        _h2(doc, SECTION_LABELS["connectivity_diagram"])
        for img_bytes, ext in images:
            _img(doc, img_bytes, ext)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────
# STREAMLIT UI — Premium Layout
# ─────────────────────────────────────────────────────────────────

# Top accent bar
st.markdown('<div class="top-bar"></div>', unsafe_allow_html=True)

# Hero section
st.markdown("""
<div class="hero">
  <p class="hero-eyebrow">// TELECOM AUTOMATION TOOLKIT</p>
  <h1 class="hero-title">Smart <em>MOP</em> Generator</h1>
  <p class="hero-sub">Upload your Solution Document and get a perfectly structured,<br>
  customer-submission ready Method of Procedure — instantly.</p>
  <div class="badge-row">
    <span class="badge b-green">⚡ IN-MEMORY ONLY</span>
    <span class="badge b-blue">📋 12-SECTION AUTO STRUCTURE</span>
    <span class="badge b-purple">🖼 IMAGES PRESERVED</span>
    <span class="badge b-green">🔒 ZERO STORAGE</span>
  </div>
</div>
""", unsafe_allow_html=True)

# Privacy notice
st.markdown("""
<div class="priv-bar">
  🔒&nbsp; <strong>Zero Data Storage Policy:</strong>&nbsp;
  All files are processed in-memory. Nothing is written to disk.
  No data is logged or retained. Session clears automatically on close.
</div>
""", unsafe_allow_html=True)

# ── STEP 1: Template ─────────────────────────────────────────────
st.markdown("""
<div class="step-card">
  <div class="step-header">
    <span class="step-number">STEP 01</span>
    <span class="step-title">Select or Upload Template</span>
  </div>
</div>
""", unsafe_allow_html=True)

# Render step 1 content OUTSIDE the HTML div (Streamlit widgets can't go inside markdown)
templates = discover_templates()
selected_template = None
template_bytes    = None

with st.container():
    if not templates:
        st.markdown("""
        <div class="pill-warn">
          ⚠️ No template found. Place <strong>Template.docx</strong>
          in the <code>templates/</code> folder or root directory, then restart.
        </div>""", unsafe_allow_html=True)
    else:
        col_sel, col_up = st.columns([2, 1])
        with col_sel:
            names = [t.name for t in templates]
            sel   = st.selectbox("Choose existing template", names, label_visibility="visible")
            selected_template = next(t for t in templates if t.name == sel)
            template_bytes    = load_template_bytes(selected_template)
            st.markdown(f'<div class="pill-ok">✅ <strong>{sel}</strong> — ready to use</div>',
                        unsafe_allow_html=True)
        with col_up:
            new_tmpl = st.file_uploader(
                "Upload new template (.docx)",
                type=["docx"],
                key="tmpl_up",
                label_visibility="visible"
            )
            if new_tmpl:
                save_dir = Path("templates")
                save_dir.mkdir(exist_ok=True)
                dest = save_dir / new_tmpl.name
                with open(dest, "wb") as f:
                    f.write(new_tmpl.read())
                st.success(f"✅ Saved: {new_tmpl.name}")
                st.rerun()

st.markdown("<br>", unsafe_allow_html=True)

# ── STEP 2: Solution Document ─────────────────────────────────────
st.markdown("""
<div class="step-card">
  <div class="step-header">
    <span class="step-number">STEP 02</span>
    <span class="step-title">Upload Solution Document</span>
  </div>
</div>
""", unsafe_allow_html=True)

sol_file = st.file_uploader(
    "Drop your Solution Document here (.docx)",
    type=["docx"],
    key="sol_up",
    label_visibility="visible"
)
if sol_file:
    size_kb = sol_file.size / 1024
    st.markdown(
        f'<div class="pill-ok">✅ <strong>{sol_file.name}</strong>'
        f' &nbsp;·&nbsp; {size_kb:.1f} KB &nbsp;·&nbsp; Ready for processing</div>',
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

# ── STEP 3: Generate ─────────────────────────────────────────────
st.markdown("""
<div class="step-card">
  <div class="step-header">
    <span class="step-number">STEP 03</span>
    <span class="step-title">Generate MOP Document</span>
  </div>
</div>
""", unsafe_allow_html=True)

can_go  = bool(sol_file and templates)
gen_btn = st.button("🚀  Generate MOP Document", disabled=not can_go)

if not can_go:
    missing = []
    if not templates:    missing.append("template")
    if not sol_file:     missing.append("solution document")
    if missing:
        st.markdown(
            f'<div class="pill-warn">Waiting for: {" + ".join(missing)}</div>',
            unsafe_allow_html=True
        )

# ── Processing ───────────────────────────────────────────────────
if gen_btn and can_go:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("""
    <div class="step-card">
      <div class="step-header">
        <span class="step-number">PROCESSING</span>
        <span class="step-title">Building your MOP</span>
      </div>
    </div>
    """, unsafe_allow_html=True)

    steps = [
        ("Loading template into memory",            "📂"),
        ("Reading solution document",               "📖"),
        ("Extracting activity name",                "🏷️"),
        ("Parsing all 12 MOP sections",             "🔍"),
        ("Detecting embedded images",               "🖼️"),
        ("Rebuilding cover page & TOC",             "📋"),
        ("Injecting section content",               "✍️"),
        ("Preserving header & footer from template","🔗"),
        ("Finalising & packaging document",         "📦"),
    ]

    st.markdown('<div class="prog-wrap">', unsafe_allow_html=True)
    phs = [st.empty() for _ in steps]
    for ph, (s, icon) in zip(phs, steps):
        ph.markdown(
            f'<div class="ps wait"><div class="pd wait"></div>{icon} {s}</div>',
            unsafe_allow_html=True
        )
    st.markdown('</div>', unsafe_allow_html=True)

    try:
        activity_name = ""
        sections      = {}
        today_str     = ""
        output_bytes  = b""

        for i, (ph, (step, icon)) in enumerate(zip(phs, steps)):
            ph.markdown(
                f'<div class="ps doing"><div class="pd doing"></div>{icon} {step}…</div>',
                unsafe_allow_html=True,
            )
            time.sleep(0.18)

            if i == 0:
                tmpl_b = load_template_bytes(selected_template)
            elif i == 1:
                sol_bytes = sol_file.read()
                sol_doc   = Document(io.BytesIO(sol_bytes))
            elif i == 2:
                activity_name = extract_activity_name(sol_doc)
                today_str     = datetime.today().strftime("%d %B %Y")
            elif i == 3:
                sections = extract_sections(sol_doc)
            elif i == 8:
                output_bytes = build_mop(tmpl_b, activity_name, sections, today_str)

            ph.markdown(
                f'<div class="ps done"><div class="pd done"></div>{icon} {step} ✓</div>',
                unsafe_allow_html=True,
            )
            time.sleep(0.05)

        # Success
        st.markdown(f"""
        <div class="success-wrap">
          <div class="success-icon">✅</div>
          <div class="success-title">MOP Generated Successfully</div>
          <div class="success-sub">
            Activity: <span class="success-name">{activity_name}</span>
            &nbsp;·&nbsp; {today_str}
          </div>
        </div>
        """, unsafe_allow_html=True)

        safe_name = re.sub(r'[^\w\s-]', '', activity_name).strip().replace(' ', '_')[:60]
        st.download_button(
            label="📥  Download MOP Document (.docx)",
            data=output_bytes,
            file_name=f"MOP_{safe_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        # Summary metrics
        filled = sum(1 for k in SECTION_KEYS[:-1] if sections.get(k))
        images_found = len(sections.get("connectivity_diagram", []))
        total_lines  = sum(len(v) for k, v in sections.items() if k != "connectivity_diagram")

        st.markdown(f"""
        <div class="metric-row">
          <div class="metric-box">
            <div class="metric-val">{filled}<span style="font-size:.9rem;color:#6e7681;">/12</span></div>
            <div class="metric-lbl">Sections Filled</div>
          </div>
          <div class="metric-box">
            <div class="metric-val">{images_found}</div>
            <div class="metric-lbl">Images Found</div>
          </div>
          <div class="metric-box">
            <div class="metric-val">{total_lines}</div>
            <div class="metric-lbl">Content Lines</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Preview expander
        with st.expander("📋  Preview extracted section content"):
            for key in SECTION_KEYS[:-1]:
                content = sections.get(key, [])
                label   = SECTION_LABELS[key]
                if content:
                    st.markdown(f"**{label}**")
                    for line in content[:3]:
                        st.markdown(
                            f"<span style='color:#8b949e;font-size:.76rem;'>→ {line[:130]}</span>",
                            unsafe_allow_html=True,
                        )
                    if len(content) > 3:
                        st.caption(f"  … +{len(content)-3} more lines")
                else:
                    st.markdown(
                        f"<span style='color:#30363d;font-size:.75rem;'>◌ {label} — no content extracted</span>",
                        unsafe_allow_html=True,
                    )

    except Exception as e:
        st.error(f"❌ Error during generation: {e}")
        import traceback
        st.code(traceback.format_exc())

elif gen_btn:
    st.warning("⚠️ Please upload a Solution Document and ensure a template is available.")

# Footer
st.markdown("""
<div class="footer">
  🔒 ZERO DATA STORAGE &nbsp;·&nbsp; IN-MEMORY PROCESSING &nbsp;·&nbsp;
  SESSION CLEARS ON CLOSE &nbsp;·&nbsp; SMART MOP GENERATOR
</div>
""", unsafe_allow_html=True)
