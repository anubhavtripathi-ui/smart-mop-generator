"""
Smart MOP Generator — v5
=========================
Changes from v4:
  1. New elegant, professional UI — Ericsson-branded (navy + signal blue + white)
  2. Template content cleared before injection (all heading bodies blanked)
  3. Both solution doc types supported (all-Normal style + Heading-style)
  4. Revision History table updated with today's date + activity name
  5. Header date updated via {{current date}} placeholder
  6. Ericsson logo / branding in sidebar
  7. Refined progress animation & summary
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
    page_title="Smart MOP Generator · Ericsson",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Session state init (persists across re-runs caused by download button) ──
for _key, _val in {
    "output_bytes":  b"",
    "activity_name": "",
    "today_str":     "",
    "sections":      {},
    "filled":        0,
    "images_n":      0,
    "total_n":       0,
}.items():
    if _key not in st.session_state:
        st.session_state[_key] = _val

# ─────────────────────────────────────────────────────────────────
# CSS — Ericsson corporate palette: navy #003366, signal blue #0082C8
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Lato:wght@300;400;700;900&family=Source+Code+Pro:wght@400;600&display=swap');

/* ── Reset & Base ── */
html, body, [class*="css"] {
  font-family: 'Lato', sans-serif;
  background-color: #F4F6F9;
  color: #1A1A2E;
}
.block-container { padding-top: 1.5rem !important; padding-bottom: 2rem; max-width: 100%; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #001f4d 0%, #003366 40%, #004080 100%) !important;
  border-right: 1px solid rgba(0,130,200,0.2);
}
[data-testid="stSidebar"] * { color: #e8f0ff !important; }
[data-testid="stSidebar"] .stMarkdown h1,
[data-testid="stSidebar"] .stMarkdown h2,
[data-testid="stSidebar"] .stMarkdown h3 { color: #ffffff !important; }
[data-testid="stSidebar"] hr { border-color: rgba(0,130,200,0.3) !important; }
[data-testid="stSidebar"] label { color: #90b8e0 !important; font-size: .78rem !important; }

/* ── Top header bar ── */
.eri-topbar {
  background: linear-gradient(90deg, #001f4d, #003366, #004080);
  border-bottom: 3px solid #0082C8;
  padding: 1rem 2rem 0.8rem;
  border-radius: 12px;
  margin-bottom: 1.5rem;
  display: flex;
  align-items: center;
  justify-content: space-between;
}
.eri-logo-text {
  font-family: 'Lato', sans-serif;
  font-weight: 900;
  font-size: 1.5rem;
  letter-spacing: 3px;
  color: #0082C8;
  text-transform: uppercase;
}
.eri-logo-sub {
  font-size: 0.7rem;
  letter-spacing: 1.5px;
  color: rgba(255,255,255,0.4);
  text-transform: uppercase;
}
.eri-app-title {
  font-size: 1.15rem;
  font-weight: 700;
  color: #ffffff;
  letter-spacing: 0.3px;
}
.eri-app-sub {
  font-size: 0.72rem;
  color: rgba(255,255,255,0.45);
  letter-spacing: 0.5px;
  margin-top: 2px;
}
.eri-version {
  background: rgba(0,130,200,0.15);
  border: 1px solid rgba(0,130,200,0.3);
  border-radius: 20px;
  padding: 3px 12px;
  font-size: 0.65rem;
  color: #0082C8;
  font-weight: 700;
  letter-spacing: 1px;
}

/* ── Privacy bar ── */
.priv-bar {
  background: rgba(0, 64, 128, 0.06);
  border: 1px solid rgba(0, 130, 200, 0.18);
  border-left: 4px solid #0082C8;
  border-radius: 0 8px 8px 0;
  padding: 0.6rem 1rem;
  font-size: 0.76rem;
  color: #003366;
  margin-bottom: 1.2rem;
}
.priv-bar strong { color: #0082C8; }

/* ── Section cards ── */
.eri-card {
  background: #ffffff;
  border: 1px solid #dde4ed;
  border-radius: 12px;
  padding: 1.4rem 1.6rem;
  margin-bottom: 1rem;
  box-shadow: 0 2px 8px rgba(0,51,102,0.06);
  transition: box-shadow 0.2s, border-color 0.2s;
}
.eri-card:hover { border-color: #0082C8; box-shadow: 0 4px 16px rgba(0,130,200,0.10); }
.eri-card-title {
  font-size: 0.68rem;
  font-weight: 900;
  letter-spacing: 1.8px;
  text-transform: uppercase;
  color: #003366;
  margin-bottom: 0.9rem;
  display: flex;
  align-items: center;
  gap: 8px;
}
.step-badge {
  background: #003366;
  color: #ffffff;
  font-size: 0.58rem;
  font-weight: 700;
  padding: 2px 8px;
  border-radius: 4px;
  letter-spacing: 0.5px;
}

/* ── Status pills ── */
.pill-ok {
  display: inline-flex; align-items: center; gap: 6px;
  background: rgba(0, 150, 80, 0.08);
  border: 1px solid rgba(0, 150, 80, 0.22);
  border-radius: 6px; padding: 5px 12px;
  font-size: 0.76rem; color: #006633; margin: 3px 0;
}
.pill-warn {
  display: inline-flex; align-items: center; gap: 6px;
  background: rgba(200, 80, 0, 0.07);
  border: 1px solid rgba(200, 80, 0, 0.2);
  border-radius: 6px; padding: 5px 12px;
  font-size: 0.76rem; color: #7a3500; margin: 3px 0;
}

/* ── Buttons ── */
.stButton > button {
  background: linear-gradient(135deg, #003366, #0082C8) !important;
  color: #ffffff !important;
  border: none !important;
  border-radius: 8px !important;
  font-family: 'Lato', sans-serif !important;
  font-weight: 700 !important;
  font-size: 0.9rem !important;
  padding: 0.6rem 2rem !important;
  width: 100% !important;
  letter-spacing: 0.4px !important;
  transition: all 0.2s !important;
}
.stButton > button:hover {
  background: linear-gradient(135deg, #004080, #0099E6) !important;
  transform: translateY(-1px) !important;
  box-shadow: 0 6px 20px rgba(0,130,200,0.28) !important;
}
.stButton > button:disabled { opacity: 0.38 !important; transform: none !important; }
[data-testid="stDownloadButton"] > button {
  background: linear-gradient(135deg, #00527A, #006FA8) !important;
  color: #ffffff !important;
  border: none !important;
  border-radius: 8px !important;
  font-family: 'Lato', sans-serif !important;
  font-weight: 700 !important;
  font-size: 0.9rem !important;
  padding: 0.6rem 2rem !important;
  width: 100% !important;
}
[data-testid="stDownloadButton"] > button:hover {
  background: linear-gradient(135deg, #006090, #0082C8) !important;
  transform: translateY(-1px) !important;
}

/* ── Progress tracker ── */
.prog-wrap {
  background: #f8fafc;
  border: 1px solid #dde4ed;
  border-radius: 10px;
  padding: 1rem 1.2rem;
}
.ps {
  display: flex; align-items: center; gap: 10px;
  padding: 6px 0; font-size: 0.78rem;
  border-bottom: 1px solid rgba(0,0,0,0.04);
}
.ps:last-child { border-bottom: none; }
.ps.done { color: #006633; font-weight: 600; }
.ps.doing { color: #0082C8; font-weight: 600; }
.ps.wait  { color: #9aaab8; }
.pd { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
.pd.done  { background: #00a050; }
.pd.doing { background: #0082C8; animation: pulse 1s infinite; }
.pd.wait  { background: #c8d4e0; }
@keyframes pulse { 0%,100%{opacity:1;transform:scale(1)} 50%{opacity:.5;transform:scale(0.85)} }

/* ── Success card ── */
.success-card {
  background: linear-gradient(135deg, rgba(0,51,102,0.04), rgba(0,130,200,0.06));
  border: 1px solid rgba(0,130,200,0.25);
  border-top: 4px solid #0082C8;
  border-radius: 12px;
  padding: 1.6rem;
  margin: 0.8rem 0;
  text-align: center;
}
.success-icon { font-size: 2.2rem; margin-bottom: 0.4rem; }
.success-title { font-size: 1.1rem; font-weight: 900; color: #003366; margin-bottom: 0.25rem; }
.success-sub   { font-size: 0.78rem; color: #0082C8; }
.success-name  { color: #003366; font-weight: 700; }

/* ── Metrics ── */
.metric-row {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: 12px;
  margin-top: 1rem;
}
.metric-box {
  background: #f0f5fb;
  border: 1px solid #dde4ed;
  border-radius: 10px;
  padding: 1rem;
  text-align: center;
}
.metric-val {
  font-family: 'Lato', sans-serif;
  font-size: 1.7rem;
  font-weight: 900;
  color: #003366;
}
.metric-sub { font-size: 0.58rem; color: #9aaab8; margin-top: 1px; font-weight: 700; letter-spacing: .6px; text-transform: uppercase; }

/* ── Sidebar widgets ── */
.sidebar-info {
  background: rgba(0,130,200,0.08);
  border: 1px solid rgba(0,130,200,0.18);
  border-radius: 8px;
  padding: 0.8rem 1rem;
  font-size: 0.74rem;
  color: #b8d4f0 !important;
  margin: 0.5rem 0;
}
.sidebar-logo {
  text-align: center;
  padding: 1.2rem 0 0.8rem;
}
.eri-big-logo {
  font-family: 'Lato', sans-serif;
  font-weight: 900;
  font-size: 2rem;
  letter-spacing: 6px;
  color: #0082C8;
  display: block;
}
.eri-tagline {
  font-size: 0.6rem;
  letter-spacing: 2px;
  color: rgba(255,255,255,0.3);
  text-transform: uppercase;
  display: block;
  margin-top: 2px;
}

/* ── Misc ── */
hr { border-color: rgba(0,51,102,0.1) !important; }
.stFileUploader label { color: #003366 !important; font-weight: 600 !important; font-size: 0.82rem !important; }
.stSelectbox label   { color: #003366 !important; font-weight: 600 !important; font-size: 0.82rem !important; }
.footer {
  text-align: center;
  font-size: 0.65rem;
  color: #9aaab8;
  padding: 1rem 0 0.4rem;
  border-top: 1px solid #dde4ed;
  letter-spacing: 0.5px;
  margin-top: 1rem;
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
    "inventory_details":    "6. Inventory Details (Node Name, Type, Count, Vendor)",
    "node_connectivity":    "7. Node Connectivity Process",
    "iam":                  "8. Identity and Access Management",
    "triggering_method":    "9. Activity Triggering Method",
    "sop":                  "10. Standard Operating Procedure (Attach the detailed SOP)",
    "acceptance_criteria":  "11. Acceptance Criteria (UAT scenarios)",
    "assumptions":          "12. Assumptions",
    "connectivity_diagram": "Connectivity Diagram",
}

PARA_SECTIONS     = {"objective", "activity_description", "activity_type",
                     "domain_in_scope", "inventory_details", "assumptions"}
BULLET_SECTIONS   = {"prerequisites", "node_connectivity", "iam",
                     "triggering_method", "acceptance_criteria"}
NUMBERED_SECTIONS = {"sop"}

# ─────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────
def _is_command_line(text: str) -> bool:
    s = text.strip()
    if re.match(r'^"[^"]{2,}"$', s):
        return True
    if re.match(r'^\[?ASSUMED\s+COMMAND', s, re.IGNORECASE):
        return True
    return False

def discover_templates() -> list:
    """Only look in templates/ folder. Never pick up root-dir .docx files."""
    tmpl_dir = Path("templates")
    if tmpl_dir.exists():
        return sorted(tmpl_dir.glob("*.docx"))
    return []

def load_template_bytes(path: Path) -> bytes:
    with open(path, "rb") as f:
        return f.read()

# ─────────────────────────────────────────────────────────────────
# SOLUTION DOC PARSER
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
    """
    Activity name = the paragraph/heading that appears just before
    the first recognised section heading (Objective) in the document.
    Works for both Type-A (Heading style) and Type-B (Normal style) solution docs.
    """
    paragraphs = doc.paragraphs

    # Find index of first recognised section heading
    obj_idx = None
    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        if not text:
            continue
        # Match by style OR by text content
        if para.style.name.startswith("Heading") and normalize_heading(text) == "objective":
            obj_idx = i
            break
        # Type-B: Normal style but text like "1. OBJECTIVE"
        if normalize_heading(text) == "objective":
            obj_idx = i
            break

    if obj_idx is not None and obj_idx > 0:
        for i in range(obj_idx - 1, -1, -1):
            text = paragraphs[i].text.strip()
            if not text:
                continue
            upper = text.upper()
            if upper in ("METHOD OF PROCEDURE", "METHOD OF PROCEDURE (MOP)",
                         "CONTENTS:", "CONTENTS", "TITLE PAGE"):
                continue
            if re.match(r'^\d+[\.\)]\s+\w.*Page\s+\d+', text):
                continue
            # ── Multi-line metadata block: "Customer: ...\nActivity Title: ..." ──
            # Check this BEFORE normalize_heading because metadata text may contain
            # domain/activity keywords that would falsely match section names
            if "\n" in text or re.match(
                    r'^(Customer|Activity Title|Document Reference|Domain|Vendor)[\s]*:',
                    text, re.IGNORECASE):
                for line in text.split("\n"):
                    line = line.strip()
                    m = re.match(r'^Activity\s+Title\s*:\s*(.+)', line, re.IGNORECASE)
                    if m:
                        return m.group(1).strip()
                continue
            # Skip if it looks like another section heading
            if normalize_heading(text) is not None:
                continue
            if re.match(r'^(Customer|Header|Footer|Document)[\s]*:', text, re.IGNORECASE):
                continue
            # Strip common prefixes
            name = re.sub(r'^MOP\s*:\s*', '', text, flags=re.IGNORECASE)
            name = re.sub(r'^UC\s*:\s*', '', name, flags=re.IGNORECASE)
            name = re.sub(r'^Activity\s+Title\s*:\s*', '', name, flags=re.IGNORECASE)
            name = re.sub(r'^Method of Procedure\s*[\(\[]?MOP[\)\]]?\s*[:\-]?\s*',
                          '', name, flags=re.IGNORECASE)
            name = name.strip()
            if name and len(name) > 3:
                return name

    # Fallback 1: Heading 1
    for para in paragraphs[:10]:
        if para.style.name.startswith("Heading 1"):
            name = para.text.strip()
            name = re.sub(r'^MOP\s*:\s*', '', name, flags=re.IGNORECASE)
            name = re.sub(r'^UC\s*:\s*', '', name, flags=re.IGNORECASE)
            if name and normalize_heading(name) is None:
                return name

    # Fallback 2: italic + underline run
    for para in paragraphs[:15]:
        for run in para.runs:
            if run.italic and run.underline and para.text.strip():
                return para.text.strip()

    return "Activity Name"


def extract_sections(doc: Document) -> dict:
    """
    Supports both doc types:
    - Type A: headings use Heading style
    - Type B: headings use Normal style but text matches section names (e.g. "1. OBJECTIVE")
    """
    sections = {k: [] for k in SECTION_KEYS}
    sections["connectivity_diagram"] = []
    current_key = None
    image_rels   = {}

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

        # Detect heading — either by style OR by normalized text matching a section key
        is_heading_style = style.startswith("Heading")
        key_from_text    = normalize_heading(text) if text else None

        if is_heading_style or (key_from_text and len(text) < 80):
            key = key_from_text or normalize_heading(text)
            if key:
                current_key = key
                continue  # Don't add heading text as content

        if current_key is None:
            continue

        # Images
        has_image = False
        for blip in para._p.findall(f".//{{{_BLIP}}}blip"):
            embed = blip.get(f"{{{_REL}}}embed")
            if embed and embed in image_rels:
                sections["connectivity_diagram"].append(image_rels[embed])
                has_image = True
        if has_image:
            continue

        # Skip noise
        if text.upper() in ("METHOD OF PROCEDURE", "METHOD OF PROCEDURE (MOP)",
                            "CONTENTS:", "CONTENTS", ""):
            continue
        if re.match(r'^\d+\.\s+\w.*Page\s+\d+', text):
            continue
        if re.match(r'^(Customer|Header|Footer|Activity Title|Document)[\s]*:', text, re.IGNORECASE):
            continue
        if text == "sample...":
            continue

        if current_key in sections:
            clean = re.sub(r'^[-–•]\s*', '', text)
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
    if size:
        run.font.size = Pt(size)
    run.font.bold      = bold
    run.font.italic    = italic
    run.font.underline = underline
    if color:
        run.font.color.rgb = RGBColor(*color)
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
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    run = p.add_run()
    br  = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._r.append(br)

def _h1(doc, text):
    """Heading 1 matching template style."""
    p = doc.add_paragraph()
    try:
        p.style = doc.styles["Heading 1"]
    except Exception:
        pass
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(text)
    run.font.name  = "Calibri"
    run.font.size  = Pt(13)
    run.font.bold  = True

def _body(doc, text):
    p = doc.add_paragraph()
    try:
        p.style = doc.styles["Body Text"]
    except Exception:
        pass
    p.paragraph_format.space_after = Pt(4)
    _r(p, text, size=11)

def _bullet(doc, text):
    if _is_command_line(text):
        _body(doc, text)
        return
    p = doc.add_paragraph(style="List Paragraph")
    p.paragraph_format.space_after = Pt(3)
    _r(p, text, size=11)

def _numbered(doc, text):
    if _is_command_line(text):
        _body(doc, text)
        return
    try:
        p = doc.add_paragraph(style="List Number")
    except Exception:
        p = doc.add_paragraph()
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

# ─────────────────────────────────────────────────────────────────
# MAIN MOP BUILDER — inject into template (preserve format)
# ─────────────────────────────────────────────────────────────────
def _insert_para_after(ref_elem, new_para_elem):
    """Insert new_para_elem immediately after ref_elem in the body."""
    ref_elem.addnext(new_para_elem)

def _new_body_para(doc, text, style_name=None, size=11,
                   bold=False, italic=False, underline=False, color=None):
    """Create a paragraph element (not added to doc yet)."""
    from docx.oxml import OxmlElement as _OE
    p = doc.add_paragraph()
    if style_name:
        try:
            p.style = doc.styles[style_name]
        except Exception:
            pass
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    if color:
        run.font.color.rgb = RGBColor(*color)
    # detach from doc body so we can reinsert at correct position
    p._element.getparent().remove(p._element)
    return p._element

def _update_revision_table(doc: Document, activity_name: str, today_str: str):
    """Update Revision History table: date + activity description."""
    for table in doc.tables:
        # Check if this looks like the revision table
        header_texts = [c.text.strip() for c in table.rows[0].cells]
        if "Version No." not in header_texts:
            continue
        if len(table.rows) >= 2:
            row = table.rows[1]
            # col 1 = Revision Date
            for para in row.cells[1].paragraphs:
                for run in para.runs:
                    run.text = ""
                if para.runs:
                    para.runs[0].text = today_str
                else:
                    run = para.add_run(today_str)
                    run.font.name = "Calibri"
            # col 3 = Description
            for para in row.cells[3].paragraphs:
                for run in para.runs:
                    run.text = ""
                if para.runs:
                    para.runs[0].text = f"Auto-generated MOP — {activity_name}"
                else:
                    run = para.add_run(f"Auto-generated MOP — {activity_name}")
                    run.font.name = "Calibri"
        break

def build_mop(template_bytes: bytes, activity_name: str,
              sections: dict, today_str: str) -> bytes:
    """
    CORRECT APPROACH:
    1. Load template as-is (preserves ALL formatting, header, footer, styles)
    2. Update header date placeholder if present
    3. Update revision table date + activity name
    4. Find each Heading 1 paragraph in the template body
    5. Remove all existing content paragraphs under that heading
    6. Insert solution doc content paragraphs right after the heading
    7. Save and return
    """
    doc = Document(io.BytesIO(template_bytes))
    body = doc.element.body

    # ── 1. Update header date ────────────────────────────────────
    _update_header_date(doc, today_str)

    # ── 2. Update title paragraph (Title style) ──────────────────
    # Find "METHOD OF PROCEDURE" title and add activity name subtitle below it
    title_elem = None
    for child in body:
        tag = child.tag.split("}")[-1]
        if tag != "p":
            continue
        style_el = child.find(".//" + qn("w:pStyle"))
        if style_el is not None and style_el.get(qn("w:val")) == "Title":
            title_elem = child
            break

    if title_elem is not None:
        subtitle = doc.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitle.paragraph_format.space_after = Pt(6)
        run = subtitle.add_run(activity_name)
        run.font.name = "Calibri"
        run.font.size = Pt(14)
        run.font.italic = True
        run.font.underline = True
        sub_elem = subtitle._element
        sub_elem.getparent().remove(sub_elem)
        title_elem.addnext(sub_elem)

    # ── 3. Update revision table ─────────────────────────────────
    _update_revision_table(doc, activity_name, today_str)

    # ── 4 & 5 & 6. For each Heading 1, clear old content & inject new ──
    # Build a map: heading_key -> list of body children between this heading and next
    # We iterate over body children to identify heading positions

    # Template heading text → section key mapping
    TMPL_HEADING_MAP = {}
    for child in list(body):
        tag = child.tag.split("}")[-1]
        if tag != "p":
            continue
        style_el = child.find(".//" + qn("w:pStyle"))
        if style_el is None:
            continue
        style_val = style_el.get(qn("w:val"), "")
        if "Heading1" not in style_val and style_val != "Heading1":
            continue
        runs = child.findall(".//" + qn("w:t"))
        text = "".join(r.text or "" for r in runs).strip()
        key = normalize_heading(text)
        if key:
            TMPL_HEADING_MAP[child] = key

    # For each heading element, remove all body children until next heading/table/sectPr
    # then insert content paragraphs
    body_children = list(body)
    for h_elem, sec_key in TMPL_HEADING_MAP.items():
        # Collect elements to remove (between this heading and next heading/table/sectPr)
        to_remove = []
        found_heading = False
        for child in body_children:
            if child is h_elem:
                found_heading = True
                continue
            if not found_heading:
                continue
            child_tag = child.tag.split("}")[-1]
            # Stop at next heading, table or sectPr
            if child_tag in ("tbl", "sectPr"):
                break
            if child_tag == "p":
                style_el = child.find(".//" + qn("w:pStyle"))
                if style_el is not None:
                    sv = style_el.get(qn("w:val"), "")
                    if "Heading" in sv:
                        break
                to_remove.append(child)

        for elem in to_remove:
            body.remove(elem)

        # Refresh body_children after removal
        body_children = list(body)

        # Now insert content after h_elem
        content = sections.get(sec_key, [])
        insert_after = h_elem

        if sec_key in PARA_SECTIONS:
            combined = " ".join(content).strip() if content else ""
            pe = _new_body_para(doc, combined, style_name="Body Text", size=11)
            insert_after.addnext(pe)
            insert_after = pe

        elif sec_key in BULLET_SECTIONS:
            # Insert in reverse so addnext keeps correct order
            items = content if content else [""]
            for item in reversed(items):
                if _is_command_line(item):
                    pe = _new_body_para(doc, item, style_name="Body Text", size=11)
                else:
                    pe = _new_body_para(doc, item, style_name="List Paragraph", size=11)
                insert_after.addnext(pe)

        elif sec_key in NUMBERED_SECTIONS:
            items = content if content else [""]
            for item in reversed(items):
                if _is_command_line(item):
                    pe = _new_body_para(doc, item, style_name="Body Text", size=11)
                else:
                    try:
                        pe = _new_body_para(doc, item, style_name="List Number", size=11)
                    except Exception:
                        pe = _new_body_para(doc, item, size=11)
                insert_after.addnext(pe)

        # Refresh again
        body_children = list(body)

    # ── 7. Connectivity diagram ───────────────────────────────────
    images = sections.get("connectivity_diagram", [])
    if images:
        for img_bytes, ext in images:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            run = p.add_run()
            run.add_picture(io.BytesIO(img_bytes), width=Inches(5))

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

# ─────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div class="sidebar-logo">
      <span class="eri-big-logo">ERICSSON</span>
      <span class="eri-tagline">Technology For Good</span>
    </div>
    <hr/>
    """, unsafe_allow_html=True)

    st.markdown("### 📋 Smart MOP Generator")
    st.markdown("""
    <div class="sidebar-info">
      Automates MOP document generation from Solution Documents.<br><br>
      <strong>Supported formats:</strong><br>
      · Normal-style solution docs<br>
      · Heading-style solution docs
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("**How to use:**")
    st.markdown("""
    <div style="font-size:0.78rem; color:#90b8e0; line-height:1.8;">
    1️⃣ &nbsp;Templates managed via <strong>GitHub</strong><br>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(<code>templates/</code> folder)<br>
    2️⃣ &nbsp;Upload your Solution Document<br>
    3️⃣ &nbsp;Click <strong>Generate MOP</strong><br>
    4️⃣ &nbsp;Click <strong>Download</strong> to save .docx
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("""
    <div class="sidebar-info">
      🔒 <strong>Zero Data Storage</strong><br>
      All processing in-memory.<br>
      No files written to disk.<br>
      Session clears on close.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("""
    <div style="font-size:0.62rem; color:rgba(255,255,255,0.2); text-align:center; letter-spacing:.5px;">
      Smart MOP Generator v5<br>
      Ericsson Internal Tool
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# MAIN AREA
# ─────────────────────────────────────────────────────────────────

# Top bar
st.markdown("""
<div class="eri-topbar">
  <div>
    <div class="eri-logo-text">ERICSSON</div>
    <div class="eri-logo-sub">Telecom Automation Toolkit</div>
  </div>
  <div style="text-align:center;">
    <div class="eri-app-title">Smart MOP Generator</div>
    <div class="eri-app-sub">Solution Document → Formatted MOP · Instant · Audit-Ready</div>
  </div>
  <div>
    <span class="eri-version">v5</span>
  </div>
</div>
""", unsafe_allow_html=True)

# Privacy bar
st.markdown("""
<div class="priv-bar">
  <strong>🔒 Zero Data Storage:</strong> Everything processed in-memory only.
  No files written to disk. No data logged. Session clears on browser close.
</div>
""", unsafe_allow_html=True)

# ── Layout: two columns ──────────────────────────────────────────
col_left, col_right = st.columns([1.1, 1], gap="large")

with col_left:

    # ── Step 1: Template ────────────────────────────────────────
    st.markdown('<div class="eri-card"><div class="eri-card-title"><span class="step-badge">STEP 01</span> Select MOP Template</div>', unsafe_allow_html=True)

    templates = discover_templates()
    selected_template = None
    template_bytes    = None

    if not templates:
        st.markdown('<div class="pill-warn">⚠ No template found. Place <strong>Automation_MOP_Template.docx</strong> in <code>templates/</code> folder, then restart.</div>', unsafe_allow_html=True)
    else:
        names = [t.name for t in templates]
        sel   = st.selectbox("Template file", names, label_visibility="visible")
        selected_template = next(t for t in templates if t.name == sel)
        template_bytes    = load_template_bytes(selected_template)
        st.markdown(f'<div class="pill-ok">✔ &nbsp;<strong>{sel}</strong> &nbsp;ready</div>', unsafe_allow_html=True)

    st.markdown("""
    <div style="font-size:0.72rem; color:#5a7a9a; background:rgba(0,82,130,0.05);
         border:1px solid rgba(0,130,200,0.15); border-left:3px solid #0082C8;
         border-radius:0 6px 6px 0; padding:0.5rem 0.8rem; margin-top:0.5rem;">
      📁 &nbsp;<strong>To add or update a template</strong>, commit the <code>.docx</code>
      file to the <code>templates/</code> folder in the GitHub repository.
      It will automatically appear here on next deployment.
    </div>
    """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Step 2: Solution Document ────────────────────────────────
    st.markdown('<div class="eri-card"><div class="eri-card-title"><span class="step-badge">STEP 02</span> Upload Solution Document</div>', unsafe_allow_html=True)

    sol_file = st.file_uploader("Solution Document (.docx)", type=["docx"],
                                key="sol_up", label_visibility="visible")
    if sol_file:
        size_kb = sol_file.size / 1024
        st.markdown(
            f'<div class="pill-ok">✔ &nbsp;<strong>{sol_file.name}</strong>'
            f' &nbsp;·&nbsp; {size_kb:.1f} KB</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown('<div class="pill-warn">⏳ Waiting for solution document…</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Step 3: Generate ─────────────────────────────────────────
    st.markdown('<div class="eri-card"><div class="eri-card-title"><span class="step-badge">STEP 03</span> Generate MOP Document</div>', unsafe_allow_html=True)

    can_go  = bool(sol_file and templates)
    gen_btn = st.button("⚡  Generate MOP", disabled=not can_go)

    if not can_go:
        missing = []
        if not templates:
            missing.append("MOP template")
        if not sol_file:
            missing.append("solution document")
        if missing:
            st.markdown(f'<div class="pill-warn">⏳ Still needed: {" + ".join(missing)}</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

with col_right:

    if gen_btn and can_go:
        st.markdown('<div class="eri-card"><div class="eri-card-title">⚙ Processing</div>', unsafe_allow_html=True)

        steps = [
            "Loading template",
            "Reading solution document",
            "Extracting activity name",
            "Parsing all 12 sections",
            "Detecting embedded images",
            "Clearing template content",
            "Injecting section content",
            "Updating revision table & header",
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
                time.sleep(0.12)

                if i == 0:
                    tmpl_b = load_template_bytes(selected_template)
                elif i == 1:
                    sol_bytes = sol_file.read()
                    sol_doc   = Document(io.BytesIO(sol_bytes))
                elif i == 2:
                    activity_name = extract_activity_name(sol_doc)
                    today_str     = datetime.today().strftime("%d-%m-%Y")
                elif i == 3:
                    sections = extract_sections(sol_doc)
                elif i == 9:
                    output_bytes = build_mop(tmpl_b, activity_name, sections, today_str)

                ph.markdown(f'<div class="ps done"><div class="pd done"></div>{step} ✓</div>', unsafe_allow_html=True)
                time.sleep(0.04)

            # ── Store everything in session_state so download survives re-run ──
            st.session_state["output_bytes"]  = output_bytes
            st.session_state["activity_name"] = activity_name
            st.session_state["today_str"]     = today_str
            st.session_state["sections"]      = sections
            st.session_state["filled"]        = sum(1 for k in SECTION_KEYS[:-1] if sections.get(k))
            st.session_state["images_n"]      = len(sections.get("connectivity_diagram", []))
            st.session_state["total_n"]       = sum(len(v) for k, v in sections.items() if k != "connectivity_diagram")

            st.markdown('</div>', unsafe_allow_html=True)

        except Exception as e:
            st.markdown('</div>', unsafe_allow_html=True)
            st.error(f"❌ Error: {e}")
            import traceback
            st.code(traceback.format_exc())

    # ── Show result panel whenever session_state has output ──────
    if st.session_state.get("output_bytes"):
        activity_name = st.session_state["activity_name"]
        today_str     = st.session_state["today_str"]
        sections      = st.session_state["sections"]
        output_bytes  = st.session_state["output_bytes"]
        filled        = st.session_state["filled"]
        images_n      = st.session_state["images_n"]
        total_n       = st.session_state["total_n"]

        # ── Success card ─────────────────────────────────────
        st.markdown(f"""
        <div class="success-card">
          <div class="success-icon">✅</div>
          <div class="success-title">MOP Generated Successfully</div>
          <div class="success-sub">
            <strong class="success-name">{activity_name}</strong>
            &nbsp;·&nbsp; {today_str}
          </div>
        </div>""", unsafe_allow_html=True)

        safe_name = re.sub(r'[^\w\s\-]', '', activity_name).strip().replace(' ', '_')[:80]
        # Unique key per generation prevents Streamlit re-run from resetting the button
        # and fixes the "Site wasn\'t available / download fails" issue.
        _dl_key = f"dl_{abs(hash(safe_name + today_str)) % 10_000_000}"
        st.download_button(
            label="📥  Download MOP Document",
            data=io.BytesIO(output_bytes),
            file_name=f"{safe_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=_dl_key,
            use_container_width=True,
        )

        # ── Summary metrics ───────────────────────────────────
        st.markdown('<div class="eri-card"><div class="eri-card-title">📊 Summary</div>', unsafe_allow_html=True)
        st.markdown(f"""
        <div class="metric-row">
          <div class="metric-box">
            <div class="metric-val">{filled}<span style="font-size:.85rem;color:#9aaab8;">/12</span></div>
            <div class="metric-sub">Sections Filled</div>
          </div>
          <div class="metric-box">
            <div class="metric-val">{images_n}</div>
            <div class="metric-sub">Images Found</div>
          </div>
          <div class="metric-box">
            <div class="metric-val">{total_n}</div>
            <div class="metric-sub">Content Lines</div>
          </div>
        </div>""", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # ── Content preview ───────────────────────────────────
        with st.expander("📋 Preview extracted content"):
            for key in SECTION_KEYS[:-1]:
                content = sections.get(key, [])
                label   = SECTION_LABELS[key]
                if content:
                    st.markdown(f"**{label}**")
                    for line in content[:3]:
                        st.markdown(
                            f"<span style='color:#003366;font-size:.75rem;'>"
                            f"→ {line[:130]}</span>",
                            unsafe_allow_html=True,
                        )
                    if len(content) > 3:
                        st.caption(f"… +{len(content) - 3} more lines")
                else:
                    st.markdown(
                        f"<span style='color:#9aaab8;font-size:.74rem;'>"
                        f"○ {label} — empty</span>",
                        unsafe_allow_html=True,
                    )

    elif gen_btn:
        st.warning("⚠ Upload a Solution Document and ensure a template is available.")

    # Show placeholder only if nothing has been generated yet in this session
    if not st.session_state.get("output_bytes"):
        st.markdown("""
        <div class="eri-card" style="border:2px dashed #dde4ed; background:#fafbfc; min-height:340px;
          display:flex; flex-direction:column; align-items:center; justify-content:center; text-align:center;">
          <div style="font-size:3rem; margin-bottom:1rem;">📄</div>
          <div style="font-size:0.95rem; font-weight:700; color:#003366; margin-bottom:0.4rem;">
            Output will appear here
          </div>
          <div style="font-size:0.76rem; color:#9aaab8; max-width:240px; line-height:1.6;">
            Upload your solution document and click Generate MOP to get started.
          </div>
        </div>
        """, unsafe_allow_html=True)

# ── Footer ────────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
  🔒 No data stored &nbsp;·&nbsp; In-memory processing only &nbsp;·&nbsp;
  Session cleared on close &nbsp;·&nbsp; Smart MOP Generator v5 &nbsp;·&nbsp; Ericsson Internal
</div>
""", unsafe_allow_html=True)
