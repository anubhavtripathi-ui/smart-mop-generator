"""
Smart MOP Generator
====================
Upload Solution Document → MOP generated in exact Template format.
Template stays on your server/GitHub (templates/ folder).
No data stored. All processing in-memory.
"""

import io
import os
import re
import copy
import time
import zipfile
import tempfile
import shutil
from datetime import datetime
from pathlib import Path

import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

# ─────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Smart MOP Generator",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@600;700;800&family=DM+Sans:wght@300;400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
.block-container { max-width: 820px; padding-top: 1.5rem; }

.hero {
    background: linear-gradient(135deg, #0b1829 0%, #0f2640 55%, #091f18 100%);
    border: 1px solid rgba(99,179,237,.15);
    border-radius: 18px; padding: 2rem 2rem 1.6rem;
    margin-bottom: 1.6rem; position: relative; overflow: hidden;
}
.hero::before {
    content:''; position:absolute; inset:0; pointer-events:none;
    background: radial-gradient(ellipse at 15% 50%, rgba(56,178,172,.07) 0%,transparent 60%),
                radial-gradient(ellipse at 85% 25%, rgba(99,179,237,.05) 0%,transparent 55%);
}
.hero-title { font-family:'Syne',sans-serif; font-size:1.9rem; font-weight:800;
              color:#e2e8f0; margin:0 0 .2rem; letter-spacing:-.4px; }
.hero-title span { color:#63b3ed; }
.hero-sub { font-size:.85rem; color:#718096; margin:0; }
.badges { display:flex; gap:7px; margin-top:1rem; flex-wrap:wrap; }
.badge { font-size:.68rem; font-weight:500; padding:3px 9px; border-radius:20px; }
.bg { background:rgba(56,178,172,.13); color:#38b2ac; border:1px solid rgba(56,178,172,.28); }
.bb { background:rgba(99,179,237,.1);  color:#63b3ed; border:1px solid rgba(99,179,237,.22); }
.bo { background:rgba(237,137,54,.1);  color:#ed8936; border:1px solid rgba(237,137,54,.22); }

.card { background:#111827; border:1px solid rgba(255,255,255,.07);
        border-radius:12px; padding:1.2rem 1.4rem; margin-bottom:.9rem; }
.card h3 { font-family:'Syne',sans-serif; font-size:.75rem; font-weight:700;
            color:#63b3ed; letter-spacing:1.3px; text-transform:uppercase; margin:0 0 .75rem; }

.priv { background:rgba(56,178,172,.06); border-left:3px solid #38b2ac;
        border-radius:0 8px 8px 0; padding:.65rem 1rem; font-size:.76rem;
        color:#68d391; margin-bottom:1.3rem; }
.priv strong { color:#9ae6b4; }

.pill-ok   { display:inline-flex; align-items:center; gap:5px; background:rgba(56,178,172,.1);
             border:1px solid rgba(56,178,172,.2); border-radius:6px;
             padding:4px 10px; font-size:.76rem; color:#81e6d9; margin:3px 0; }
.pill-info { display:inline-flex; align-items:center; gap:5px; background:rgba(99,179,237,.1);
             border:1px solid rgba(99,179,237,.2); border-radius:6px;
             padding:4px 10px; font-size:.76rem; color:#90cdf4; margin:3px 0; }
.pill-warn { display:inline-flex; align-items:center; gap:5px; background:rgba(237,137,54,.1);
             border:1px solid rgba(237,137,54,.2); border-radius:6px;
             padding:4px 10px; font-size:.76rem; color:#f6ad55; margin:3px 0; }

.stButton>button {
    background:linear-gradient(135deg,#2b6cb0,#2c7a7b)!important;
    color:#fff!important; border:none!important; border-radius:10px!important;
    font-family:'Syne',sans-serif!important; font-weight:700!important;
    font-size:.92rem!important; padding:.6rem 2rem!important; width:100%!important;
    transition:all .18s!important;
}
.stButton>button:hover { background:linear-gradient(135deg,#3182ce,#319795)!important;
    transform:translateY(-1px)!important; box-shadow:0 6px 20px rgba(99,179,237,.2)!important; }
.stButton>button:disabled { opacity:.4!important; transform:none!important; }

[data-testid="stDownloadButton"]>button {
    background:linear-gradient(135deg,#276749,#285e61)!important;
    color:#fff!important; border:none!important; border-radius:10px!important;
    font-family:'Syne',sans-serif!important; font-weight:700!important;
    font-size:.92rem!important; padding:.6rem 2rem!important; width:100%!important;
}

.ps { display:flex; align-items:center; gap:9px; padding:6px 0;
      font-size:.78rem; border-bottom:1px solid rgba(255,255,255,.04); }
.ps.done  { color:#68d391; } .ps.doing { color:#63b3ed; } .ps.wait { color:#4a5568; }
.pd { width:7px; height:7px; border-radius:50%; flex-shrink:0; }
.pd.done  { background:#68d391; }
.pd.doing { background:#63b3ed; animation:blink 1s infinite; }
.pd.wait  { background:#2d3748; }
@keyframes blink { 0%,100%{opacity:1} 50%{opacity:.3} }
hr { border-color:rgba(255,255,255,.06)!important; }

[data-testid="stFileUploader"] {
    background:#0d1520!important;
    border:1.5px dashed rgba(99,179,237,.28)!important;
    border-radius:10px!important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# CONSTANTS — Heading mapping: solution doc heading → section key
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

# Section keys in order
SECTION_KEYS = [
    "objective", "activity_description", "activity_type", "domain_in_scope",
    "prerequisites", "inventory_details", "node_connectivity", "iam",
    "triggering_method", "sop", "acceptance_criteria", "assumptions",
    "connectivity_diagram",
]

# Template section labels (must match template headings exactly)
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

# TOC page numbers (fixed layout matching Final MOP)
TOC_PAGES = {
    "objective": 2, "activity_description": 2, "activity_type": 2,
    "domain_in_scope": 2, "prerequisites": 2, "inventory_details": 3,
    "node_connectivity": 3, "iam": 3, "triggering_method": 3,
    "sop": 3, "acceptance_criteria": 4, "assumptions": 4,
}

# Paragraph vs bullet vs numbered
PARA_SECTIONS    = {"objective","activity_description","activity_type",
                    "domain_in_scope","inventory_details","assumptions"}
BULLET_SECTIONS  = {"prerequisites","node_connectivity","iam",
                    "triggering_method","acceptance_criteria"}
NUMBERED_SECTIONS = {"sop"}

# ─────────────────────────────────────────────────────────────────
# TEMPLATE LOADER
# ─────────────────────────────────────────────────────────────────
TEMPLATES_DIR = Path("templates")
TEMPLATES_DIR.mkdir(exist_ok=True)

def list_templates():
    return sorted([f for f in TEMPLATES_DIR.glob("*.docx")])

def load_template_bytes(path: Path) -> bytes:
    with open(path, "rb") as f:
        return f.read()

# ─────────────────────────────────────────────────────────────────
# SOLUTION DOCUMENT PARSER
# ─────────────────────────────────────────────────────────────────
def normalize_heading(text: str):
    t = re.sub(r'^\d+[\.\)]\s*', '', text).strip().lower()
    t = re.sub(r'\s+', ' ', t)
    for key, aliases in HEADING_MAP.items():
        for alias in aliases:
            if alias in t:
                return key
    return None

def extract_activity_name(doc: Document) -> str:
    for para in doc.paragraphs[:6]:
        if para.style.name.startswith("Heading 1"):
            name = re.sub(r'^MOP:\s*', '', para.text.strip(), flags=re.IGNORECASE)
            return name
    for para in doc.paragraphs[:8]:
        if para.runs and para.runs[0].italic and para.runs[0].underline:
            return para.text.strip()
    return "Activity Name"

def extract_sections(doc: Document) -> dict:
    sections = {k: [] for k in SECTION_KEYS}
    sections["connectivity_diagram"] = []
    current_key = None
    image_rels = {}

    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            try:
                ext = rel.target_part.content_type.split("/")[-1]
                if ext == "jpeg": ext = "jpg"
                image_rels[rel.rId] = (rel.target_part.blob, ext)
            except Exception:
                pass

    _BLIP = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    _REL  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

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

        # Image check FIRST (image paras have empty text)
        has_image = False
        for blip in para._p.findall(f'.//{{{_BLIP}}}blip'):
            embed = blip.get(f'{{{_REL}}}embed')
            if embed and embed in image_rels:
                sections["connectivity_diagram"].append(image_rels[embed])
                has_image = True
        if has_image:
            continue

        # Skip cover/TOC lines and blanks
        if text in ("METHOD OF PROCEDURE", "CONTENTS:", "CONTENTS", ""):
            continue
        if re.match(r'^\d+\.\s+\w.*Page\s+\d+', text):
            continue

        if current_key in sections:
            # Clean leading markers
            clean = re.sub(r'^[-–•]\s*', '', text)
            clean = re.sub(r'^\d+[\.\)]\s*', '', clean)
            sections[current_key].append(clean.strip())

    return sections

# ─────────────────────────────────────────────────────────────────
# MOP BUILDER — injects content into template copy
# ─────────────────────────────────────────────────────────────────
def _add_run(para, text, font="Calibri", size=None, bold=False,
             italic=False, underline=False, color=None):
    run = para.add_run(text)
    run.font.name = font
    if size: run.font.size = Pt(size)
    run.font.bold      = bold
    run.font.italic    = italic
    run.font.underline = underline
    if color: run.font.color.rgb = RGBColor(*color)
    return run

def _set_right_tab(para, pos=8400):
    pPr = para._p.get_or_add_pPr()
    for old in pPr.findall(qn("w:tabs")): pPr.remove(old)
    tabs = OxmlElement("w:tabs")
    tab  = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:pos"), str(pos))
    tabs.append(tab); pPr.append(tabs)

def _page_break(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    run = p.add_run()
    br  = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._r.append(br)

def _update_header_date(doc: Document, today_str: str):
    """Replace {{current date}} placeholder in header."""
    for section in doc.sections:
        header = section.header
        for para in header.paragraphs:
            for run in para.runs:
                if "{{current date}}" in run.text:
                    run.text = run.text.replace("{{current date}}", today_str)
            # Also check full para text for split runs
            if "{{current date}}" in para.text:
                full = para.text
                # Clear and rebuild
                for run in para.runs:
                    if "{{current date}}" in run.text:
                        run.text = run.text.replace("{{current date}}", today_str)

def _clear_body(doc: Document):
    """Remove all body paragraphs (keep section properties)."""
    body = doc.element.body
    # Save sectPr
    sectPr = body.find(qn("w:sectPr"))
    # Remove all children
    for child in list(body):
        body.remove(child)
    # Add back sectPr
    if sectPr is not None:
        body.append(sectPr)

def _add_heading2(doc: Document, text: str):
    """Add Heading 2 styled paragraph."""
    p = doc.add_paragraph()
    p.style = doc.styles["Heading 2"]
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(2)
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x4F, 0x81, 0xBD)

def _add_body_para(doc: Document, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    _add_run(p, text, size=11)

def _add_bullet(doc: Document, text: str):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    _add_run(p, text, size=11)

def _add_numbered(doc: Document, text: str):
    p = doc.add_paragraph(style="List Number")
    p.paragraph_format.space_after = Pt(3)
    _add_run(p, text, size=11)

def _add_image(doc: Document, img_bytes: bytes, ext: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    run = p.add_run()
    run.add_picture(io.BytesIO(img_bytes), width=Inches(5))

def build_mop(template_bytes: bytes, activity_name: str,
              sections: dict, today_str: str) -> bytes:
    """
    Core engine:
    1. Load template into memory
    2. Update header date
    3. Clear body, rebuild from scratch with template styles intact
    """
    # Load template from bytes (in-memory, nothing saved to disk)
    doc = Document(io.BytesIO(template_bytes))

    # Step 1: Update header date
    _update_header_date(doc, today_str)

    # Step 2: Clear existing body content
    _clear_body(doc)

    # ── PAGE 1: Cover ─────────────────────────────────────────────

    # "METHOD OF PROCEDURE"
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(4)
    _add_run(p_title, "METHOD OF PROCEDURE",
             size=18, bold=True, color=(0x7F, 0x7F, 0x7F))

    # Activity Name — underlined, italic, centered
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_after = Pt(4)
    _add_run(p_name, activity_name, size=14, italic=True, underline=True)

    # Blank line
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # "CONTENTS:"
    p_cont = doc.add_paragraph()
    p_cont.paragraph_format.space_after = Pt(4)
    _add_run(p_cont, "CONTENTS:", size=12, bold=True, underline=True)

    # TOC entries with right tab
    for key in SECTION_KEYS[:-1]:
        tp = doc.add_paragraph()
        tp.paragraph_format.space_after = Pt(2)
        _set_right_tab(tp, 8400)
        _add_run(tp, SECTION_LABELS[key], size=11)
        _add_run(tp, f"\tPage {TOC_PAGES.get(key, 2)}", size=11)

    # Page break
    _page_break(doc)

    # ── PAGE 2+: Sections ─────────────────────────────────────────
    for key in SECTION_KEYS[:-1]:
        content = sections.get(key, [])
        _add_heading2(doc, SECTION_LABELS[key])

        if key in PARA_SECTIONS:
            _add_body_para(doc, " ".join(content).strip() if content else "")
        elif key in BULLET_SECTIONS:
            for item in content:
                _add_bullet(doc, item)
        elif key in NUMBERED_SECTIONS:
            for item in content:
                _add_numbered(doc, item)

    # Connectivity Diagram (if present)
    images = sections.get("connectivity_diagram", [])
    if images:
        _add_heading2(doc, SECTION_LABELS["connectivity_diagram"])
        for img_bytes, ext in images:
            _add_image(doc, img_bytes, ext)

    # Save to bytes — nothing written to disk
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

# ─────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────

# Hero
st.markdown("""
<div class="hero">
  <p class="hero-title">Smart <span>MOP</span> Generator</p>
  <p class="hero-sub">Upload Solution Document → Instantly get a perfectly formatted MOP</p>
  <div class="badges">
    <span class="badge bg">⚡ In-Memory Only</span>
    <span class="badge bb">📋 Auto TOC + Page Numbers</span>
    <span class="badge bo">🖼️ Images Preserved</span>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="priv">
  <strong>🔒 Zero Data Storage:</strong> Everything processed in-memory.
  No files written to disk. No data logged. Session clears on close.
</div>
""", unsafe_allow_html=True)

# ── Template selector ─────────────────────────────────────────────
st.markdown('<div class="card"><h3>📂 Step 1 — Select Template</h3>', unsafe_allow_html=True)

templates = list_templates()

if not templates:
    st.markdown("""
    <div class="pill-warn">⚠️ No templates found in <code>templates/</code> folder.
    Add your Template.docx file there and restart.</div>
    """, unsafe_allow_html=True)
    template_bytes = None
    selected_template = None
else:
    template_names = [t.name for t in templates]
    col_sel, col_up = st.columns([2, 1])
    with col_sel:
        selected_name = st.selectbox(
            "Choose template",
            template_names,
            label_visibility="collapsed"
        )
        selected_template = TEMPLATES_DIR / selected_name
        template_bytes = load_template_bytes(selected_template)
        st.markdown(f'<div class="pill-ok">✅ Template: <strong>{selected_name}</strong></div>',
                    unsafe_allow_html=True)
    with col_up:
        new_tmpl = st.file_uploader("Add new template", type=["docx"],
                                     key="tmpl_upload", label_visibility="collapsed")
        if new_tmpl:
            save_path = TEMPLATES_DIR / new_tmpl.name
            with open(save_path, "wb") as f:
                f.write(new_tmpl.read())
            st.success(f"Saved: {new_tmpl.name}")
            st.rerun()

st.markdown('</div>', unsafe_allow_html=True)

# ── Solution document upload ──────────────────────────────────────
st.markdown('<div class="card"><h3>📤 Step 2 — Upload Solution Document</h3>',
            unsafe_allow_html=True)
sol_file = st.file_uploader(
    "Upload Solution Document (.docx)",
    type=["docx"], key="sol_upload",
    label_visibility="collapsed"
)
if sol_file:
    st.markdown(
        f'<div class="pill-ok">✅ Loaded: <strong>{sol_file.name}</strong>'
        f' &nbsp;·&nbsp; {sol_file.size/1024:.1f} KB</div>',
        unsafe_allow_html=True
    )
st.markdown('</div>', unsafe_allow_html=True)

# ── Generate ──────────────────────────────────────────────────────
st.markdown('<div class="card"><h3>⚡ Step 3 — Generate MOP</h3>', unsafe_allow_html=True)
can_generate = bool(sol_file and templates)
gen_btn = st.button("🚀 Generate MOP Document", disabled=not can_generate)
st.markdown('</div>', unsafe_allow_html=True)

if gen_btn and sol_file and templates:
    st.markdown('<div class="card"><h3>⚙️ Processing</h3>', unsafe_allow_html=True)

    steps = [
        "Loading template",
        "Reading solution document",
        "Extracting activity name",
        "Parsing all 12 sections",
        "Detecting images",
        "Building cover page & TOC",
        "Injecting section content",
        "Finalising document",
    ]
    phs = [st.empty() for _ in steps]
    for ph, s in zip(phs, steps):
        ph.markdown(f'<div class="ps wait"><div class="pd wait"></div>{s}</div>',
                    unsafe_allow_html=True)

    try:
        activity_name = ""; sections = {}
        today_str = ""; output_bytes = b""

        for i, (ph, step) in enumerate(zip(phs, steps)):
            ph.markdown(f'<div class="ps doing"><div class="pd doing"></div>{step}</div>',
                        unsafe_allow_html=True)
            time.sleep(0.18)

            if i == 0:
                tmpl_bytes = load_template_bytes(selected_template)
            elif i == 1:
                sol_bytes = sol_file.read()
                sol_doc   = Document(io.BytesIO(sol_bytes))
            elif i == 2:
                activity_name = extract_activity_name(sol_doc)
                today_str     = datetime.today().strftime("%d %B %Y")
            elif i == 3:
                sections = extract_sections(sol_doc)
            elif i == 7:
                output_bytes = build_mop(tmpl_bytes, activity_name, sections, today_str)

            ph.markdown(f'<div class="ps done"><div class="pd done"></div>{step} ✓</div>',
                        unsafe_allow_html=True)
            time.sleep(0.05)

        st.markdown('</div>', unsafe_allow_html=True)

        # Success
        st.markdown(f"""
        <div style="background:rgba(56,178,172,.08);border:1px solid rgba(56,178,172,.22);
             border-radius:12px;padding:1.3rem;margin:.8rem 0;text-align:center;">
          <div style="font-family:'Syne',sans-serif;font-size:1rem;font-weight:700;
               color:#9ae6b4;margin-bottom:.25rem;">✅ MOP Generated Successfully</div>
          <div style="font-size:.78rem;color:#68d391;">
            Activity: <strong style="color:#9ae6b4;">{activity_name}</strong>
            &nbsp;·&nbsp; Template: <strong style="color:#9ae6b4;">{selected_name}</strong>
          </div>
        </div>""", unsafe_allow_html=True)

        safe = re.sub(r'[^\w\s-]', '', activity_name).strip().replace(' ', '_')[:60]
        st.download_button(
            label="📥 Download MOP.docx",
            data=output_bytes,
            file_name=f"MOP_{safe}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        # Summary
        st.markdown('<div class="card"><h3>📊 Summary</h3>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        filled = sum(1 for k in SECTION_KEYS[:-1] if sections.get(k))
        with c1: st.metric("Sections Filled", f"{filled}/12")
        with c2: st.metric("Images Found", len(sections.get("connectivity_diagram", [])))
        with c3:
            total = sum(len(v) for k, v in sections.items()
                        if k != "connectivity_diagram")
            st.metric("Content Lines", total)
        st.markdown('</div>', unsafe_allow_html=True)

        with st.expander("📋 Preview extracted content"):
            for key in SECTION_KEYS[:-1]:
                content = sections.get(key, [])
                label   = SECTION_LABELS[key]
                if content:
                    st.markdown(f"**{label}**")
                    for line in content[:3]:
                        st.markdown(
                            f"<span style='color:#a0aec0;font-size:.76rem;'>• {line[:120]}</span>",
                            unsafe_allow_html=True
                        )
                    if len(content) > 3:
                        st.caption(f"... +{len(content)-3} more")
                else:
                    st.markdown(
                        f"<span style='color:#4a5568;font-size:.76rem;'>"
                        f"{label} — empty</span>",
                        unsafe_allow_html=True
                    )

    except Exception as e:
        st.markdown('</div>', unsafe_allow_html=True)
        st.error(f"❌ Error: {e}")
        import traceback
        st.code(traceback.format_exc())

elif gen_btn:
    st.warning("⚠️ Please upload a Solution Document and ensure a template is available.")

# Footer
st.markdown("""<br>
<div style="text-align:center;font-size:.68rem;color:#2d3748;padding:.7rem 0;
     border-top:1px solid rgba(255,255,255,.04);">
  🔒 No data stored &nbsp;·&nbsp; In-memory processing only &nbsp;·&nbsp;
  Session cleared on close &nbsp;·&nbsp; Smart MOP Generator
</div>""", unsafe_allow_html=True)
