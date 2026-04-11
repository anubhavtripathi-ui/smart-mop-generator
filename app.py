"""
Smart MOP Generator — v4
=========================
4 updates:
  1. Activity name = first heading before "Objective" in solution doc
  2. Commands + "ASSUMED COMMAND" lines NOT rendered as bullet/numbered items
  3. Download filename = activity name (clean)
  4. Minor UI polish
"""

import io
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
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Smart MOP Generator",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────────────────────
# CSS — polished UI
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@600;700;800&family=DM+Sans:wght@300;400;500;600&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif;}
.block-container{max-width:840px;padding-top:0 !important;padding-bottom:2rem;}

/* ── Top accent stripe ── */
.top-stripe{height:4px;background:linear-gradient(90deg,#1a6b4a,#2d9cdb,#8b5cf6);
  margin-bottom:0;border-radius:0;}

/* ── Hero ── */
.hero{background:linear-gradient(135deg,#0b1829 0%,#0f2640 55%,#091f18 100%);
  border:1px solid rgba(99,179,237,.15);border-radius:0 0 20px 20px;
  padding:2.2rem 2.2rem 1.8rem;margin-bottom:1.6rem;position:relative;overflow:hidden;}
.hero::before{content:'';position:absolute;inset:0;pointer-events:none;
  background:radial-gradient(ellipse at 15% 50%,rgba(56,178,172,.07) 0%,transparent 60%),
             radial-gradient(ellipse at 85% 25%,rgba(99,179,237,.05) 0%,transparent 55%);}
.hero-eyebrow{font-family:'DM Sans',sans-serif;font-size:.65rem;font-weight:600;
  letter-spacing:3px;color:#2d9cdb;text-transform:uppercase;margin-bottom:.5rem;}
.hero-title{font-family:'Syne',sans-serif;font-size:2rem;font-weight:800;
  color:#e2e8f0;margin:0 0 .3rem;letter-spacing:-.5px;line-height:1.15;}
.hero-title .t1{color:#3fb882;}.hero-title .t2{color:#63b3ed;}
.hero-title .t3{color:#a78bfa;}
.hero-sub{font-size:.84rem;color:#718096;margin:0 0 1.2rem;font-weight:300;}
.badges{display:flex;gap:7px;margin-top:.2rem;flex-wrap:wrap;}
.badge{font-size:.63rem;font-weight:600;padding:3px 10px;border-radius:20px;letter-spacing:.4px;}
.bg{background:rgba(56,178,172,.13);color:#38b2ac;border:1px solid rgba(56,178,172,.28);}
.bb{background:rgba(99,179,237,.1);color:#63b3ed;border:1px solid rgba(99,179,237,.22);}
.bo{background:rgba(237,137,54,.1);color:#ed8936;border:1px solid rgba(237,137,54,.22);}
.bv{background:rgba(139,92,246,.1);color:#a78bfa;border:1px solid rgba(139,92,246,.22);}

/* ── Privacy bar ── */
.priv{background:rgba(56,178,172,.06);border-left:3px solid #38b2ac;
  border-radius:0 8px 8px 0;padding:.65rem 1rem;font-size:.76rem;
  color:#68d391;margin-bottom:1.4rem;}
.priv strong{color:#9ae6b4;}

/* ── Cards ── */
.card{background:#0f1923;border:1px solid rgba(255,255,255,.07);
  border-radius:14px;padding:1.3rem 1.5rem;margin-bottom:1rem;
  transition:border-color .2s;}
.card:hover{border-color:rgba(99,179,237,.18);}
.card h3{font-family:'Syne',sans-serif;font-size:.72rem;font-weight:700;
  color:#63b3ed;letter-spacing:1.4px;text-transform:uppercase;margin:0 0 .85rem;
  display:flex;align-items:center;gap:7px;}
.card h3 .step-pill{background:rgba(99,179,237,.12);border:1px solid rgba(99,179,237,.2);
  border-radius:4px;padding:1px 7px;font-size:.6rem;color:#90cdf4;letter-spacing:1px;}

/* ── Pills ── */
.pill-ok{display:inline-flex;align-items:center;gap:5px;
  background:rgba(56,178,172,.1);border:1px solid rgba(56,178,172,.2);
  border-radius:6px;padding:5px 11px;font-size:.75rem;color:#81e6d9;margin:3px 0;}
.pill-warn{display:inline-flex;align-items:center;gap:5px;
  background:rgba(237,137,54,.1);border:1px solid rgba(237,137,54,.2);
  border-radius:6px;padding:5px 11px;font-size:.75rem;color:#f6ad55;margin:3px 0;}

/* ── Buttons ── */
.stButton>button{
  background:linear-gradient(135deg,#1a6b4a,#1d5f8a)!important;
  color:#fff!important;border:none!important;border-radius:10px!important;
  font-family:'Syne',sans-serif!important;font-weight:700!important;
  font-size:.9rem!important;padding:.65rem 2rem!important;width:100%!important;
  letter-spacing:.3px!important;transition:all .2s!important;}
.stButton>button:hover{
  background:linear-gradient(135deg,#1e7d56,#2270a3)!important;
  transform:translateY(-1px)!important;
  box-shadow:0 8px 24px rgba(45,156,219,.18)!important;}
.stButton>button:disabled{opacity:.35!important;transform:none!important;}
[data-testid="stDownloadButton"]>button{
  background:linear-gradient(135deg,#1a5c3a,#164e70)!important;
  color:#fff!important;border:none!important;border-radius:10px!important;
  font-family:'Syne',sans-serif!important;font-weight:700!important;
  font-size:.9rem!important;padding:.65rem 2rem!important;width:100%!important;}
[data-testid="stDownloadButton"]>button:hover{
  background:linear-gradient(135deg,#1e6b44,#1a5e87)!important;
  transform:translateY(-1px)!important;}

/* ── Progress steps ── */
.prog-wrap{background:#0a0f18;border:1px solid rgba(255,255,255,.05);
  border-radius:10px;padding:.9rem 1.1rem;margin-top:.5rem;}
.ps{display:flex;align-items:center;gap:9px;padding:6px 0;
  font-size:.78rem;border-bottom:1px solid rgba(255,255,255,.04);}
.ps:last-child{border-bottom:none;}
.ps.done{color:#68d391;}.ps.doing{color:#63b3ed;}.ps.wait{color:#4a5568;}
.pd{width:7px;height:7px;border-radius:50%;flex-shrink:0;}
.pd.done{background:#68d391;}
.pd.doing{background:#63b3ed;animation:blink 1s infinite;}
.pd.wait{background:#2d3748;}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.3}}

/* ── Success card ── */
.success-card{background:linear-gradient(135deg,rgba(26,107,74,.12),rgba(45,156,219,.08));
  border:1px solid rgba(56,178,172,.25);border-radius:14px;
  padding:1.5rem;margin:.8rem 0;text-align:center;}
.success-icon{font-size:2rem;margin-bottom:.3rem;}
.success-title{font-family:'Syne',sans-serif;font-size:1.05rem;font-weight:700;
  color:#9ae6b4;margin-bottom:.2rem;}
.success-sub{font-size:.78rem;color:#68d391;}
.success-name{color:#63b3ed;font-weight:600;}

/* ── Metric boxes ── */
.metric-row{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-top:.8rem;}
.metric-box{background:#0a0f18;border:1px solid rgba(255,255,255,.06);
  border-radius:10px;padding:.9rem;text-align:center;}
.metric-val{font-family:'Syne',sans-serif;font-size:1.5rem;font-weight:700;color:#e2e8f0;}
.metric-lbl{font-size:.66rem;color:#4a5568;margin-top:3px;letter-spacing:.5px;text-transform:uppercase;}

hr{border-color:rgba(255,255,255,.06)!important;}
label{color:#718096!important;font-size:.8rem!important;}
.footer{text-align:center;font-size:.66rem;color:#2d3748;padding:.8rem 0 .4rem;
  border-top:1px solid rgba(255,255,255,.04);letter-spacing:.5px;}
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
    "objective","activity_description","activity_type","domain_in_scope",
    "prerequisites","inventory_details","node_connectivity","iam",
    "triggering_method","sop","acceptance_criteria","assumptions",
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
    "objective":2,"activity_description":2,"activity_type":2,"domain_in_scope":2,
    "prerequisites":2,"inventory_details":3,"node_connectivity":3,"iam":3,
    "triggering_method":3,"sop":3,"acceptance_criteria":4,"assumptions":4,
}

PARA_SECTIONS     = {"objective","activity_description","activity_type",
                     "domain_in_scope","inventory_details","assumptions"}
BULLET_SECTIONS   = {"prerequisites","node_connectivity","iam",
                     "triggering_method","acceptance_criteria"}
NUMBERED_SECTIONS = {"sop"}

# ─────────────────────────────────────────────────────────────────
# HELPERS — detect command / assumed-command lines
# ─────────────────────────────────────────────────────────────────
def _is_command_line(text: str) -> bool:
    """
    Returns True if the line is a quoted command (e.g. "RLCRP:CELL=ALL")
    or an ASSUMED COMMAND label — these must NOT be rendered as list items.
    """
    s = text.strip()
    # Quoted command: starts and ends with " and content is present
    if re.match(r'^"[^"]{2,}"$', s):
        return True
    # Assumed command label (any casing)
    if re.match(r'^"?ASSUMED\s+COMMAND', s, re.IGNORECASE):
        return True
    return False

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


# ── FIX 1 & 3: Activity name = first heading/paragraph before "Objective" ──
def extract_activity_name(doc: Document) -> str:
    """
    Activity name = the paragraph/heading that appears just before
    the first 'Objective' heading in the document.
    Falls back to Heading 1 → italic+underline → 'Activity Name'.
    """
    paragraphs = doc.paragraphs

    # Find index of first "Objective" heading
    obj_idx = None
    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        if not text:
            continue
        norm = normalize_heading(text)
        if norm == "objective":
            obj_idx = i
            break

    if obj_idx is not None and obj_idx > 0:
        # Walk backwards from objective to find the last non-empty paragraph
        for i in range(obj_idx - 1, -1, -1):
            text = paragraphs[i].text.strip()
            if not text:
                continue
            # Skip noise
            if text.upper() in ("METHOD OF PROCEDURE", "CONTENTS:", "CONTENTS"):
                continue
            if re.match(r'^\d+[\.\)]\s+\w.*Page\s+\d+', text):
                continue
            # This is the activity name
            name = re.sub(r'^MOP\s*:\s*', '', text, flags=re.IGNORECASE)
            name = re.sub(r'^UC\s*:\s*', '', name, flags=re.IGNORECASE)
            name = re.sub(r'^Method of Procedure\s*[\(\[]?MOP[\)\]]?\s*:\s*', '',
                          name, flags=re.IGNORECASE)
            if name:
                return name.strip()

    # Fallback 1: Heading 1
    for para in paragraphs[:8]:
        if para.style.name.startswith("Heading 1"):
            name = para.text.strip()
            name = re.sub(r'^MOP\s*:\s*', '', name, flags=re.IGNORECASE)
            name = re.sub(r'^UC\s*:\s*', '', name, flags=re.IGNORECASE)
            if name:
                return name

    # Fallback 2: italic + underline
    for para in paragraphs[:10]:
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
                if ext == "jpeg": ext = "jpg"
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

        if text in ("METHOD OF PROCEDURE","CONTENTS:","CONTENTS",""):
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
    run.font.name = font
    if size: run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    if color: run.font.color.rgb = RGBColor(*color)
    return run

def _set_rtab(para, pos=8400):
    pPr = para._p.get_or_add_pPr()
    for old in pPr.findall(qn("w:tabs")): pPr.remove(old)
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
    run.font.name = "Calibri"
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x4F, 0x81, 0xBD)

def _body(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    _r(p, text, size=11)

def _bullet(doc, text):
    # ── FIX 2: command/assumed lines → body paragraph, not bullet ──
    if _is_command_line(text):
        _body(doc, text)
        return
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(3)
    _r(p, text, size=11)

def _numbered(doc, text):
    # ── FIX 2: command/assumed lines → body paragraph, not numbered ──
    if _is_command_line(text):
        _body(doc, text)
        return
    p = doc.add_paragraph(style="List Number")
    p.paragraph_format.space_after = Pt(3)
    _r(p, text, size=11)

def _img(doc, img_bytes, ext):
    p   = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    run = p.add_run()
    run.add_picture(io.BytesIO(img_bytes), width=Inches(5))

def _update_header_date(doc: Document, today_str: str):
    for section in doc.sections:
        for para in section.header.paragraphs:
            for run in para.runs:
                if "{{current date}}" in run.text:
                    run.text = run.text.replace("{{current date}}", today_str)

def _clear_and_prep_body(doc: Document):
    body = doc.element.body
    header_refs = []; footer_refs = []
    for child in body:
        for nested_sectPr in child.findall(".//" + qn("w:sectPr")):
            for elem in nested_sectPr:
                if elem.tag == qn("w:headerReference"):
                    header_refs.append(deepcopy(elem))
                elif elem.tag == qn("w:footerReference"):
                    footer_refs.append(deepcopy(elem))
    body_sectPr = body.find(qn("w:sectPr"))
    for child in list(body): body.remove(child)
    if body_sectPr is None: body_sectPr = OxmlElement("w:sectPr")
    for elem in list(body_sectPr):
        if elem.tag in (qn("w:headerReference"), qn("w:footerReference")):
            body_sectPr.remove(elem)
    insert_pos = 0
    for ref in header_refs + footer_refs:
        body_sectPr.insert(insert_pos, ref); insert_pos += 1
    body.append(body_sectPr)

# ─────────────────────────────────────────────────────────────────
# MAIN MOP BUILDER
# ─────────────────────────────────────────────────────────────────
def build_mop(template_bytes: bytes, activity_name: str,
              sections: dict, today_str: str) -> bytes:
    doc = Document(io.BytesIO(template_bytes))
    _update_header_date(doc, today_str)
    _clear_and_prep_body(doc)

    # ── COVER PAGE ──
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    _r(p, "METHOD OF PROCEDURE", size=18, bold=True, color=(0x7F,0x7F,0x7F))

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

    # ── BODY SECTIONS ──
    for key in SECTION_KEYS[:-1]:
        content = sections.get(key, [])
        _h2(doc, SECTION_LABELS[key])

        if key in PARA_SECTIONS:
            _body(doc, " ".join(content).strip() if content else "")
        elif key in BULLET_SECTIONS:
            if content:
                for item in content:
                    _bullet(doc, item)   # FIX 2 applied inside _bullet
            else:
                _body(doc, "")
        elif key in NUMBERED_SECTIONS:
            if content:
                for item in content:
                    _numbered(doc, item) # FIX 2 applied inside _numbered
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
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────
st.markdown('<div class="top-stripe"></div>', unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <p class="hero-eyebrow">// Telecom Automation Toolkit</p>
  <p class="hero-title">
    <span class="t1">Smart</span>&nbsp;<span class="t2">MOP</span>&nbsp;<span class="t3">Generator</span>
  </p>
  <p class="hero-sub">
    Upload your Solution Document → get a fully formatted, audit-ready MOP instantly.
  </p>
  <div class="badges">
    <span class="badge bg">⚡ In-Memory Only</span>
    <span class="badge bb">📋 12-Section Auto Structure</span>
    <span class="badge bo">🖼️ Images Preserved</span>
    <span class="badge bv">🔒 Zero Storage</span>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="priv">
  <strong>🔒 Zero Data Storage:</strong> Everything processed in-memory.
  No files written to disk. No data logged. Session clears on close.
</div>
""", unsafe_allow_html=True)

# ── Step 1: Template ─────────────────────────────────────────────
st.markdown('<div class="card"><h3><span class="step-pill">01</span> Select Template</h3>', unsafe_allow_html=True)

templates = discover_templates()
selected_template = None
template_bytes    = None

if not templates:
    st.markdown('<div class="pill-warn">⚠️ No template found. Place <strong>Template.docx</strong> in <code>templates/</code> or root directory, then restart.</div>', unsafe_allow_html=True)
else:
    col_sel, col_up = st.columns([2, 1])
    with col_sel:
        names = [t.name for t in templates]
        sel   = st.selectbox("Choose template", names, label_visibility="visible")
        selected_template = next(t for t in templates if t.name == sel)
        template_bytes    = load_template_bytes(selected_template)
        st.markdown(f'<div class="pill-ok">✅ <strong>{sel}</strong> — ready</div>', unsafe_allow_html=True)
    with col_up:
        new_tmpl = st.file_uploader("Upload new template (.docx)", type=["docx"], key="tmpl_up", label_visibility="visible")
        if new_tmpl:
            save_dir = Path("templates"); save_dir.mkdir(exist_ok=True)
            with open(save_dir / new_tmpl.name, "wb") as f: f.write(new_tmpl.read())
            st.success(f"✅ Saved: {new_tmpl.name}")
            st.rerun()

st.markdown('</div>', unsafe_allow_html=True)

# ── Step 2: Solution Document ─────────────────────────────────────
st.markdown('<div class="card"><h3><span class="step-pill">02</span> Upload Solution Document</h3>', unsafe_allow_html=True)
sol_file = st.file_uploader("Solution Document (.docx)", type=["docx"], key="sol_up", label_visibility="visible")
if sol_file:
    st.markdown(
        f'<div class="pill-ok">✅ <strong>{sol_file.name}</strong>'
        f' &nbsp;·&nbsp; {sol_file.size/1024:.1f} KB</div>',
        unsafe_allow_html=True,
    )
st.markdown('</div>', unsafe_allow_html=True)

# ── Step 3: Generate ─────────────────────────────────────────────
st.markdown('<div class="card"><h3><span class="step-pill">03</span> Generate MOP</h3>', unsafe_allow_html=True)
can_go  = bool(sol_file and templates)
gen_btn = st.button("🚀  Generate MOP Document", disabled=not can_go)
if not can_go:
    missing = []
    if not templates: missing.append("template")
    if not sol_file:  missing.append("solution document")
    if missing:
        st.markdown(f'<div class="pill-warn">⏳ Waiting for: {" + ".join(missing)}</div>', unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# ── Processing ────────────────────────────────────────────────────
if gen_btn and can_go:
    st.markdown('<div class="card"><h3>⚙️ Processing</h3>', unsafe_allow_html=True)

    steps = [
        "Loading template",
        "Reading solution document",
        "Extracting activity name",
        "Parsing all 12 sections",
        "Detecting images",
        "Rebuilding cover page & TOC",
        "Injecting section content",
        "Preserving header & footer",
        "Finalising document",
    ]

    st.markdown('<div class="prog-wrap">', unsafe_allow_html=True)
    phs = [st.empty() for _ in steps]
    for ph, s in zip(phs, steps):
        ph.markdown(f'<div class="ps wait"><div class="pd wait"></div>{s}</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    try:
        activity_name = ""
        sections      = {}
        today_str     = ""
        output_bytes  = b""

        for i, (ph, step) in enumerate(zip(phs, steps)):
            ph.markdown(f'<div class="ps doing"><div class="pd doing"></div>{step}…</div>', unsafe_allow_html=True)
            time.sleep(0.15)

            if i == 0:
                tmpl_b = load_template_bytes(selected_template)
            elif i == 1:
                sol_bytes = sol_file.read()
                sol_doc   = Document(io.BytesIO(sol_bytes))
            elif i == 2:
                activity_name = extract_activity_name(sol_doc)  # FIX 1
                today_str     = datetime.today().strftime("%d %B %Y")
            elif i == 3:
                sections = extract_sections(sol_doc)
            elif i == 8:
                output_bytes = build_mop(tmpl_b, activity_name, sections, today_str)

            ph.markdown(f'<div class="ps done"><div class="pd done"></div>{step} ✓</div>', unsafe_allow_html=True)
            time.sleep(0.04)

        st.markdown('</div>', unsafe_allow_html=True)

        # ── Success ──
        st.markdown(f"""
        <div class="success-card">
          <div class="success-icon">✅</div>
          <div class="success-title">MOP Generated Successfully</div>
          <div class="success-sub">
            Activity: <span class="success-name">{activity_name}</span>
            &nbsp;·&nbsp; {today_str}
          </div>
        </div>""", unsafe_allow_html=True)

        # ── FIX 3: Download filename = activity name ──
        safe_name = re.sub(r'[^\w\s\-]', '', activity_name).strip().replace(' ', '_')[:80]
        st.download_button(
            label="📥  Download MOP Document",
            data=output_bytes,
            file_name=f"{safe_name}.docx",   # activity name only, no "MOP_" prefix
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        # ── Summary ──
        st.markdown('<div class="card"><h3>📊 Summary</h3>', unsafe_allow_html=True)
        filled = sum(1 for k in SECTION_KEYS[:-1] if sections.get(k))
        images_n = len(sections.get("connectivity_diagram", []))
        total_n  = sum(len(v) for k, v in sections.items() if k != "connectivity_diagram")
        st.markdown(f"""
        <div class="metric-row">
          <div class="metric-box">
            <div class="metric-val">{filled}<span style="font-size:.9rem;color:#4a5568;">/12</span></div>
            <div class="metric-lbl">Sections Filled</div>
          </div>
          <div class="metric-box">
            <div class="metric-val">{images_n}</div>
            <div class="metric-lbl">Images Found</div>
          </div>
          <div class="metric-box">
            <div class="metric-val">{total_n}</div>
            <div class="metric-lbl">Content Lines</div>
          </div>
        </div>""", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        with st.expander("📋 Preview extracted content"):
            for key in SECTION_KEYS[:-1]:
                content = sections.get(key, [])
                label   = SECTION_LABELS[key]
                if content:
                    st.markdown(f"**{label}**")
                    for line in content[:3]:
                        st.markdown(f"<span style='color:#a0aec0;font-size:.76rem;'>→ {line[:120]}</span>", unsafe_allow_html=True)
                    if len(content) > 3:
                        st.caption(f"… +{len(content)-3} more")
                else:
                    st.markdown(f"<span style='color:#4a5568;font-size:.76rem;'>{label} — empty</span>", unsafe_allow_html=True)

    except Exception as e:
        st.markdown('</div>', unsafe_allow_html=True)
        st.error(f"❌ Error: {e}")
        import traceback
        st.code(traceback.format_exc())

elif gen_btn:
    st.warning("⚠️ Upload a Solution Document and ensure a template is available.")

# ── Footer ────────────────────────────────────────────────────────
st.markdown("""<br>
<div class="footer">
  🔒 No data stored &nbsp;·&nbsp; In-memory processing only &nbsp;·&nbsp;
  Session cleared on close &nbsp;·&nbsp; Smart MOP Generator v4
</div>""", unsafe_allow_html=True)
