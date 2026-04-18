"""
Smart MOP Generator — v7
=========================
Changes from v6:
  1. Page orientation: Portrait (A4)
  2. Date (header + revision table): always today's date
  3. Description of Change: "Initial draft — <Activity Name>"
  4. Heading 1 color: Blue Accent 1 Dark 25% (#2F5496)
  5. Output filename: Activity Name
  6. "METHOD OF PROCEDURE" + Activity Name below it on cover
  7. Activity Name = solution doc filename (stem)
"""

import io
import re
import time
from datetime import datetime
from pathlib import Path
from copy import deepcopy

import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ─────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Smart MOP Generator · Ericsson",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)

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
# CSS
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Lato:wght@300;400;700;900&display=swap');
html,body,[class*="css"]{font-family:'Lato',sans-serif;background:#F4F6F9;color:#1A1A2E;}
.block-container{padding-top:1.5rem!important;padding-bottom:2rem;max-width:100%;}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#001f4d,#003366 40%,#004080)!important;border-right:1px solid rgba(0,130,200,.2);}
[data-testid="stSidebar"] *{color:#e8f0ff!important;}
[data-testid="stSidebar"] hr{border-color:rgba(0,130,200,.3)!important;}
[data-testid="stSidebar"] label{color:#90b8e0!important;font-size:.78rem!important;}
.eri-topbar{background:linear-gradient(90deg,#001f4d,#003366,#004080);border-bottom:3px solid #0082C8;padding:1rem 2rem .8rem;border-radius:12px;margin-bottom:1.5rem;display:flex;align-items:center;justify-content:space-between;}
.eri-logo-text{font-weight:900;font-size:1.5rem;letter-spacing:3px;color:#0082C8;text-transform:uppercase;}
.eri-logo-sub{font-size:.7rem;letter-spacing:1.5px;color:rgba(255,255,255,.4);text-transform:uppercase;}
.eri-app-title{font-size:1.15rem;font-weight:700;color:#fff;letter-spacing:.3px;}
.eri-app-sub{font-size:.72rem;color:rgba(255,255,255,.45);letter-spacing:.5px;margin-top:2px;}
.eri-version{background:rgba(0,130,200,.15);border:1px solid rgba(0,130,200,.3);border-radius:20px;padding:3px 12px;font-size:.65rem;color:#0082C8;font-weight:700;letter-spacing:1px;}
.priv-bar{background:rgba(0,64,128,.06);border:1px solid rgba(0,130,200,.18);border-left:4px solid #0082C8;border-radius:0 8px 8px 0;padding:.6rem 1rem;font-size:.76rem;color:#003366;margin-bottom:1.2rem;}
.priv-bar strong{color:#0082C8;}
.eri-card{background:#fff;border:1px solid #dde4ed;border-radius:12px;padding:1.4rem 1.6rem;margin-bottom:1rem;box-shadow:0 2px 8px rgba(0,51,102,.06);transition:box-shadow .2s,border-color .2s;}
.eri-card:hover{border-color:#0082C8;box-shadow:0 4px 16px rgba(0,130,200,.10);}
.eri-card-title{font-size:.68rem;font-weight:900;letter-spacing:1.8px;text-transform:uppercase;color:#003366;margin-bottom:.9rem;display:flex;align-items:center;gap:8px;}
.step-badge{background:#003366;color:#fff;font-size:.58rem;font-weight:700;padding:2px 8px;border-radius:4px;letter-spacing:.5px;}
.pill-ok{display:inline-flex;align-items:center;gap:6px;background:rgba(0,150,80,.08);border:1px solid rgba(0,150,80,.22);border-radius:6px;padding:5px 12px;font-size:.76rem;color:#006633;margin:3px 0;}
.pill-warn{display:inline-flex;align-items:center;gap:6px;background:rgba(200,80,0,.07);border:1px solid rgba(200,80,0,.2);border-radius:6px;padding:5px 12px;font-size:.76rem;color:#7a3500;margin:3px 0;}
.stButton>button{background:linear-gradient(135deg,#003366,#0082C8)!important;color:#fff!important;border:none!important;border-radius:8px!important;font-family:'Lato',sans-serif!important;font-weight:700!important;font-size:.9rem!important;padding:.6rem 2rem!important;width:100%!important;letter-spacing:.4px!important;transition:all .2s!important;}
.stButton>button:hover{background:linear-gradient(135deg,#004080,#0099E6)!important;transform:translateY(-1px)!important;box-shadow:0 6px 20px rgba(0,130,200,.28)!important;}
.stButton>button:disabled{opacity:.38!important;transform:none!important;}
[data-testid="stDownloadButton"]>button{background:linear-gradient(135deg,#00527A,#006FA8)!important;color:#fff!important;border:none!important;border-radius:8px!important;font-family:'Lato',sans-serif!important;font-weight:700!important;font-size:.9rem!important;padding:.6rem 2rem!important;width:100%!important;}
[data-testid="stDownloadButton"]>button:hover{background:linear-gradient(135deg,#006090,#0082C8)!important;transform:translateY(-1px)!important;}
.prog-wrap{background:#f8fafc;border:1px solid #dde4ed;border-radius:10px;padding:1rem 1.2rem;}
.ps{display:flex;align-items:center;gap:10px;padding:6px 0;font-size:.78rem;border-bottom:1px solid rgba(0,0,0,.04);}
.ps:last-child{border-bottom:none;}
.ps.done{color:#006633;font-weight:600;}.ps.doing{color:#0082C8;font-weight:600;}.ps.wait{color:#9aaab8;}
.pd{width:8px;height:8px;border-radius:50%;flex-shrink:0;}
.pd.done{background:#00a050;}.pd.doing{background:#0082C8;animation:pulse 1s infinite;}.pd.wait{background:#c8d4e0;}
@keyframes pulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.5;transform:scale(.85)}}
.success-card{background:linear-gradient(135deg,rgba(0,51,102,.04),rgba(0,130,200,.06));border:1px solid rgba(0,130,200,.25);border-top:4px solid #0082C8;border-radius:12px;padding:1.6rem;margin:.8rem 0;text-align:center;}
.success-icon{font-size:2.2rem;margin-bottom:.4rem;}
.success-title{font-size:1.1rem;font-weight:900;color:#003366;margin-bottom:.25rem;}
.success-sub{font-size:.78rem;color:#0082C8;}
.success-name{color:#003366;font-weight:700;}
.metric-row{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-top:1rem;}
.metric-box{background:#f0f5fb;border:1px solid #dde4ed;border-radius:10px;padding:1rem;text-align:center;}
.metric-val{font-size:1.7rem;font-weight:900;color:#003366;}
.metric-sub{font-size:.58rem;color:#9aaab8;margin-top:1px;font-weight:700;letter-spacing:.6px;text-transform:uppercase;}
.sidebar-info{background:rgba(0,130,200,.08);border:1px solid rgba(0,130,200,.18);border-radius:8px;padding:.8rem 1rem;font-size:.74rem;color:#b8d4f0!important;margin:.5rem 0;}
.sidebar-logo{text-align:center;padding:1.2rem 0 .8rem;}
.eri-big-logo{font-weight:900;font-size:2rem;letter-spacing:6px;color:#0082C8;display:block;}
.eri-tagline{font-size:.6rem;letter-spacing:2px;color:rgba(255,255,255,.3);text-transform:uppercase;display:block;margin-top:2px;}
hr{border-color:rgba(0,51,102,.1)!important;}
.footer{text-align:center;font-size:.65rem;color:#9aaab8;padding:1rem 0 .4rem;border-top:1px solid #dde4ed;letter-spacing:.5px;margin-top:1rem;}
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

ALLOWED_TEMPLATE = "Automation_MOP_Template.docx"

# Blue Accent 1 Dark 25% — Word standard value
HEADING_COLOR = RGBColor(0x2F, 0x54, 0x96)   # #2F5496

# A4 Portrait dimensions in EMU (1 inch = 914400 EMU; 1 twip = 635 EMU)
# A4: 210mm x 297mm = 11906 x 16838 twips (landscape) → portrait = 11906 wide x 16838 tall
# In twips: portrait width=11906, height=16838  (these are A4 values Word uses)
A4_W = 11906   # twips  (~210mm)
A4_H = 16838   # twips  (~297mm)

# ─────────────────────────────────────────────────────────────────
# TEMPLATE DISCOVERY
# ─────────────────────────────────────────────────────────────────
def discover_templates() -> list:
    candidates = []
    tmpl_dir = Path("templates")
    if tmpl_dir.exists():
        candidates += list(tmpl_dir.glob("*.docx"))
    candidates += [p for p in Path(".").glob("*.docx")
                   if p.name not in [c.name for c in candidates]]
    return [p for p in candidates if p.name == ALLOWED_TEMPLATE]

def load_template_bytes(path: Path) -> bytes:
    with open(path, "rb") as f:
        return f.read()

# ─────────────────────────────────────────────────────────────────
# HEADING NORMALIZER
# ─────────────────────────────────────────────────────────────────
def normalize_heading(text: str):
    t = re.sub(r'^\d+[\.\)]\s*', '', text).strip().lower()
    t = re.sub(r'\s+', ' ', t)
    for key, aliases in HEADING_MAP.items():
        for alias in aliases:
            if alias in t:
                return key
    return None

# ─────────────────────────────────────────────────────────────────
# ACTIVITY NAME — always = solution doc filename stem
# ─────────────────────────────────────────────────────────────────
def get_activity_name(sol_file_stem: str) -> str:
    """
    Activity Name = solution document filename (without extension).
    Underscores/hyphens replaced with spaces for readability.
    """
    name = sol_file_stem.replace("_", " ").replace("-", " ").strip()
    return name if name else "Activity Name"

# ─────────────────────────────────────────────────────────────────
# SECTION EXTRACTOR
# ─────────────────────────────────────────────────────────────────
def extract_sections(doc: Document) -> dict:
    sections = {k: [] for k in SECTION_KEYS}
    sections["connectivity_diagram"] = []
    current_key = None
    image_rels = {}

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

        is_heading    = style.startswith("Heading")
        key_from_text = normalize_heading(text) if text else None

        if is_heading or (key_from_text and len(text) < 100):
            key = normalize_heading(text) if text else None
            if key:
                current_key = key
            continue  # always skip heading lines

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

        if text.upper() in ("METHOD OF PROCEDURE", "METHOD OF PROCEDURE (MOP)",
                            "CONTENTS:", "CONTENTS", ""):
            continue
        if re.match(r'^\d+\.\s+\w.*Page\s+\d+', text):
            continue
        if re.match(r'^(Customer|Header|Footer|Activity Title|Document)\s*:',
                    text, re.IGNORECASE):
            continue
        if text == "sample...":
            continue

        if current_key in sections:
            sections[current_key].append({"text": text, "style": style})

    return sections

# ─────────────────────────────────────────────────────────────────
# BUILD MOP — template-preserving + all 7 fixes
# ─────────────────────────────────────────────────────────────────
def build_mop(template_bytes: bytes, activity_name: str,
              sections: dict, today_str: str) -> bytes:

    doc = Document(io.BytesIO(template_bytes))
    body = doc.element.body

    # ── FIX 1: Set page to Portrait A4 ──────────────────────────
    sectPr = body.find(qn("w:sectPr"))
    if sectPr is not None:
        pgSz = sectPr.find(qn("w:pgSz"))
        if pgSz is None:
            pgSz = OxmlElement("w:pgSz")
            sectPr.insert(0, pgSz)
        # Set portrait A4: width < height, remove orient=landscape
        pgSz.set(qn("w:w"), str(A4_W))
        pgSz.set(qn("w:h"), str(A4_H))
        # Remove landscape orientation attribute
        orient_attr = qn("w:orient")
        if orient_attr in pgSz.attrib:
            del pgSz.attrib[orient_attr]
        # Remove the code attribute (printer paper code) to avoid conflicts
        code_attr = qn("w:code")
        if code_attr in pgSz.attrib:
            del pgSz.attrib[code_attr]
        # Update margins for portrait (swap top/left if needed — keep current margins)
        pgMar = sectPr.find(qn("w:pgMar"))
        if pgMar is not None:
            # Portrait sensible margins: top=1440, bottom=1440, left=1800, right=1440
            pgMar.set(qn("w:top"),    "1440")
            pgMar.set(qn("w:bottom"), "1440")
            pgMar.set(qn("w:left"),   "1800")
            pgMar.set(qn("w:right"),  "1440")
            pgMar.set(qn("w:header"), "720")
            pgMar.set(qn("w:footer"), "720")

    # Also fix header table width to match portrait content width
    # Portrait content width = A4_W - left_margin - right_margin = 11906 - 1800 - 1440 = 8666 twips
    PORTRAIT_CONTENT_W = 8666

    # ── FIX 4: Set Heading 1 color to Blue Accent 1 Dark 25% ────
    for style in doc.styles:
        if style.name == "Heading 1":
            rPr = style.element.find(qn("w:rPr"))
            if rPr is None:
                rPr = OxmlElement("w:rPr")
                style.element.append(rPr)
            # Remove existing color elements
            for old_color in rPr.findall(qn("w:color")):
                rPr.remove(old_color)
            color_el = OxmlElement("w:color")
            color_el.set(qn("w:val"), "2F5496")
            color_el.set(qn("w:themeColor"), "accent1")
            color_el.set(qn("w:themeTint"), "BF")  # ~75% tint = dark 25%
            rPr.append(color_el)
            break

    # ── FIX 2 & 3: Update header table date + revision table ─────
    # Header table: find Date cell and update
    for section in doc.sections:
        hdr = section.header
        for tbl in hdr.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        full_text = para.text.strip()
                        # Date value cells
                        if re.match(r'^\d{4}-\d{2}-\d{2}$', full_text) or \
                           re.match(r'^\d{2}-\d{2}-\d{4}$', full_text) or \
                           full_text in ("DD-MM-YYYY", "{{date}}", "{{current date}}"):
                            for run in para.runs:
                                if run.text.strip():
                                    run.text = today_str
                            # If no runs, add one
                            if not any(r.text.strip() for r in para.runs):
                                run = para.add_run(today_str)

    # Body revision table: update date + description
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                full_text = cell.text.strip()
                # Update date cells
                if re.match(r'^\d{4}-\d{2}-\d{2}$', full_text) or \
                   re.match(r'^\d{2}-\d{2}-\d{4}$', full_text) or \
                   full_text in ("DD-MM-YYYY", "17-04-2026", "16-04-2026", "{{date}}"):
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.text = today_str
                # Update Description of Change — FIX 3
                if "Initial draft" in full_text or full_text in ("Activity Name",):
                    for para in cell.paragraphs:
                        for run in para.runs:
                            if "Initial draft" in run.text or run.text.strip() == "Activity Name":
                                run.text = f"Initial draft \u2014 {activity_name}"

    # ── FIX 6: Insert Activity Name below METHOD OF PROCEDURE ────
    # Find the Title paragraph ("METHOD OF PROCEDURE")
    # and insert Activity Name paragraph right after it
    title_elem = None
    for child in list(body):
        if child.tag == qn("w:p"):
            style_el = child.find(f".//{qn('w:pStyle')}")
            style_val = style_el.get(qn("w:val"), "") if style_el is not None else ""
            text = "".join(t.text or "" for t in child.findall(f".//{qn('w:t')}")).strip()
            if "Title" in style_val or "METHOD OF PROCEDURE" in text.upper():
                title_elem = child
                break

    if title_elem is not None:
        # Build Activity Name paragraph (centered, italic, 14pt)
        act_p = OxmlElement("w:p")
        act_pPr = OxmlElement("w:pPr")
        act_jc = OxmlElement("w:jc")
        act_jc.set(qn("w:val"), "center")
        act_pPr.append(act_jc)
        act_p.append(act_pPr)
        act_r = OxmlElement("w:r")
        act_rPr = OxmlElement("w:rPr")
        # Italic
        i_el = OxmlElement("w:i")
        act_rPr.append(i_el)
        # Size 14pt = 28 half-points
        sz_el = OxmlElement("w:sz")
        sz_el.set(qn("w:val"), "28")
        act_rPr.append(sz_el)
        # Color matching heading
        c_el = OxmlElement("w:color")
        c_el.set(qn("w:val"), "2F5496")
        act_rPr.append(c_el)
        act_r.append(act_rPr)
        act_t = OxmlElement("w:t")
        act_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        act_t.text = activity_name
        act_r.append(act_t)
        act_p.append(act_r)

        # Insert after title_elem
        title_elem.addnext(act_p)

    # ── Clear placeholder content + Inject solution doc content ──
    body_children = list(body)
    heading_positions = {}
    for idx, child in enumerate(body_children):
        if child.tag == qn("w:p"):
            style_el = child.find(f".//{qn('w:pStyle')}")
            if style_el is not None:
                style_val = style_el.get(qn("w:val"), "")
                if "Heading1" in style_val or style_val == "1":
                    text = "".join(
                        t.text or "" for t in child.findall(f".//{qn('w:t')}")
                    ).strip()
                    key = normalize_heading(text)
                    if key and key not in heading_positions:
                        heading_positions[key] = idx

    # Clear placeholder content under each heading (skip Revision History)
    sorted_keys = sorted(heading_positions.items(), key=lambda x: x[1])
    for i, (key, start_idx) in enumerate(sorted_keys):
        end_idx = sorted_keys[i + 1][1] if i + 1 < len(sorted_keys) else len(body_children)
        heading_elem = body_children[start_idx]
        heading_text = "".join(
            t.text or "" for t in heading_elem.findall(f".//{qn('w:t')}")
        ).strip().lower()
        if "revision" in heading_text:
            continue
        to_remove = []
        for j in range(start_idx + 1, end_idx):
            if j < len(body_children):
                elem = body_children[j]
                if elem.tag in (qn("w:sectPr"), qn("w:tbl")):
                    continue
                to_remove.append(elem)
        for elem in to_remove:
            try:
                body.remove(elem)
            except Exception:
                pass

    # Re-discover heading positions after removal
    body_children = list(body)
    heading_positions = {}
    for idx, child in enumerate(body_children):
        if child.tag == qn("w:p"):
            style_el = child.find(f".//{qn('w:pStyle')}")
            if style_el is not None:
                style_val = style_el.get(qn("w:val"), "")
                if "Heading1" in style_val or style_val == "1":
                    text = "".join(
                        t.text or "" for t in child.findall(f".//{qn('w:t')}")
                    ).strip()
                    key = normalize_heading(text)
                    if key and key not in heading_positions:
                        heading_positions[key] = idx

    # Inject content (reverse order to maintain positions)
    sorted_keys = sorted(heading_positions.items(), key=lambda x: x[1], reverse=True)
    for key, heading_idx in sorted_keys:
        if key == "connectivity_diagram":
            continue
        content_items = sections.get(key, [])
        heading_elem  = body_children[heading_idx]
        heading_parent = heading_elem.getparent()
        if heading_parent is None:
            continue
        try:
            h_idx_in_parent = list(heading_parent).index(heading_elem)
        except ValueError:
            continue

        paras_to_insert = []
        if not content_items:
            paras_to_insert.append(OxmlElement("w:p"))
        else:
            for item in content_items:
                text      = item["text"]
                orig_style = item["style"]
                new_p = OxmlElement("w:p")
                pPr   = OxmlElement("w:pPr")
                pStyle = OxmlElement("w:pStyle")
                if orig_style == "List Paragraph":
                    pStyle.set(qn("w:val"), "ListParagraph")
                elif orig_style == "List Number":
                    pStyle.set(qn("w:val"), "ListNumber")
                elif orig_style == "Body Text":
                    pStyle.set(qn("w:val"), "BodyText")
                else:
                    pStyle.set(qn("w:val"), "Normal")
                pPr.append(pStyle)
                new_p.append(pPr)
                r  = OxmlElement("w:r")
                t_el = OxmlElement("w:t")
                t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                t_el.text = text
                r.append(t_el)
                new_p.append(r)
                paras_to_insert.append(new_p)

        insert_pos = h_idx_in_parent + 1
        for para_elem in paras_to_insert:
            heading_parent.insert(insert_pos, para_elem)
            insert_pos += 1

    # Connectivity diagram images
    images = sections.get("connectivity_diagram", [])
    if images:
        h_p = OxmlElement("w:p")
        h_pPr = OxmlElement("w:pPr")
        h_pStyle = OxmlElement("w:pStyle")
        h_pStyle.set(qn("w:val"), "Heading1")
        h_pPr.append(h_pStyle)
        h_p.append(h_pPr)
        h_r = OxmlElement("w:r")
        h_t = OxmlElement("w:t")
        h_t.text = "Connectivity Diagram"
        h_r.append(h_t)
        h_p.append(h_r)
        body_list = list(body)
        last = body_list[-1]
        if last.tag == qn("w:sectPr"):
            body.insert(body_list.index(last), h_p)
        else:
            body.append(h_p)
        for img_bytes, ext in images:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.add_run().add_picture(io.BytesIO(img_bytes), width=Inches(5))

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
      <strong>Supported:</strong><br>
      · Normal-style solution docs<br>
      · Heading-style solution docs
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("**How to use:**")
    st.markdown("""
    <div style="font-size:.78rem;color:#90b8e0;line-height:1.8;">
    1️⃣ &nbsp;Templates via <strong>GitHub</strong> (<code>templates/</code>)<br>
    2️⃣ &nbsp;Upload your Solution Document<br>
    3️⃣ &nbsp;Click <strong>Generate MOP</strong><br>
    4️⃣ &nbsp;Download the output .docx
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
    <div style="font-size:.62rem;color:rgba(255,255,255,.2);text-align:center;letter-spacing:.5px;">
      Smart MOP Generator v7<br>Ericsson Internal Tool
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# MAIN AREA
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<div class="eri-topbar">
  <div>
    <div class="eri-logo-text">ERICSSON</div>
    <div class="eri-logo-sub">Telecom Automation Toolkit</div>
  </div>
  <div style="text-align:center;">
    <div class="eri-app-title">Smart MOP Generator</div>
    <div class="eri-app-sub">Solution Document → Template-Format MOP · Portrait · Audit-Ready</div>
  </div>
  <div><span class="eri-version">v7</span></div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="priv-bar">
  <strong>🔒 Zero Data Storage:</strong> Everything processed in-memory only.
  No files written to disk. No data logged. Session clears on browser close.
</div>
""", unsafe_allow_html=True)

col_left, col_right = st.columns([1.1, 1], gap="large")

with col_left:

    # ── Step 1: Template ────────────────────────────────────────
    st.markdown('<div class="eri-card"><div class="eri-card-title"><span class="step-badge">STEP 01</span> MOP Template</div>', unsafe_allow_html=True)
    templates = discover_templates()
    selected_template = None

    if not templates:
        st.markdown(
            f'<div class="pill-warn">⚠ Template not found — place '
            f'<strong>{ALLOWED_TEMPLATE}</strong> in <code>templates/</code> folder.</div>',
            unsafe_allow_html=True,
        )
    else:
        selected_template = templates[0]
        st.markdown(
            f'<div class="pill-ok">✔ &nbsp;<strong>{templates[0].name}</strong> &nbsp;ready</div>',
            unsafe_allow_html=True,
        )
    st.markdown("""
    <div style="font-size:.72rem;color:#5a7a9a;background:rgba(0,82,130,.05);
         border:1px solid rgba(0,130,200,.15);border-left:3px solid #0082C8;
         border-radius:0 6px 6px 0;padding:.5rem .8rem;margin-top:.5rem;">
      📁 &nbsp;<strong>To update template</strong>, commit to
      <code>templates/</code> on GitHub. Auto-reflects on next deploy.
    </div>
    """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Step 2: Solution Document ────────────────────────────────
    st.markdown('<div class="eri-card"><div class="eri-card-title"><span class="step-badge">STEP 02</span> Upload Solution Document</div>', unsafe_allow_html=True)
    sol_file = st.file_uploader("Solution Document (.docx)", type=["docx"],
                                key="sol_up", label_visibility="visible")
    if sol_file:
        act_preview = get_activity_name(Path(sol_file.name).stem)
        size_kb = sol_file.size / 1024
        st.markdown(
            f'<div class="pill-ok">✔ &nbsp;<strong>{sol_file.name}</strong>'
            f' &nbsp;·&nbsp; {size_kb:.1f} KB</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            f'<div style="font-size:.73rem;color:#003366;margin-top:.4rem;">'
            f'📄 Activity Name: <strong>{act_preview}</strong></div>',
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
        if not templates: missing.append("MOP template")
        if not sol_file:  missing.append("solution document")
        if missing:
            st.markdown(f'<div class="pill-warn">⏳ Still needed: {" + ".join(missing)}</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

with col_right:

    if gen_btn and can_go:
        st.markdown('<div class="eri-card"><div class="eri-card-title">⚙ Processing</div>', unsafe_allow_html=True)
        steps = [
            "Loading template",
            "Reading solution document",
            "Resolving activity name from filename",
            "Parsing all 12 sections",
            "Detecting embedded images",
            "Setting portrait orientation",
            "Applying heading colour (Blue Accent 1 Dark 25%)",
            "Clearing template placeholder content",
            "Injecting content into template headings",
            "Updating date & revision table",
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
                time.sleep(0.10)

                if i == 0:
                    tmpl_b = load_template_bytes(selected_template)
                elif i == 1:
                    sol_bytes = sol_file.read()
                    sol_doc   = Document(io.BytesIO(sol_bytes))
                elif i == 2:
                    # FIX 5 & 7: Activity Name = solution doc filename
                    activity_name = get_activity_name(Path(sol_file.name).stem)
                    today_str     = datetime.today().strftime("%d-%m-%Y")
                elif i == 3:
                    sections = extract_sections(sol_doc)
                elif i == 10:
                    output_bytes = build_mop(tmpl_b, activity_name, sections, today_str)

                ph.markdown(f'<div class="ps done"><div class="pd done"></div>{step} ✓</div>', unsafe_allow_html=True)
                time.sleep(0.04)

            st.session_state["output_bytes"]  = output_bytes
            st.session_state["activity_name"] = activity_name
            st.session_state["today_str"]     = today_str
            st.session_state["sections"]      = sections
            st.session_state["filled"]        = sum(1 for k in SECTION_KEYS[:-1] if sections.get(k))
            st.session_state["images_n"]      = len(sections.get("connectivity_diagram", []))
            st.session_state["total_n"]       = sum(len(v) for k, v in sections.items()
                                                    if k != "connectivity_diagram")
            st.markdown('</div>', unsafe_allow_html=True)

        except Exception as e:
            st.markdown('</div>', unsafe_allow_html=True)
            st.error(f"❌ Error: {e}")
            import traceback
            st.code(traceback.format_exc())

    # ── Result panel ─────────────────────────────────────────────
    if st.session_state.get("output_bytes"):
        activity_name = st.session_state["activity_name"]
        today_str     = st.session_state["today_str"]
        sections      = st.session_state["sections"]
        output_bytes  = st.session_state["output_bytes"]
        filled        = st.session_state["filled"]
        images_n      = st.session_state["images_n"]
        total_n       = st.session_state["total_n"]

        st.markdown(f"""
        <div class="success-card">
          <div class="success-icon">✅</div>
          <div class="success-title">MOP Generated Successfully</div>
          <div class="success-sub">
            <strong class="success-name">{activity_name}</strong>
            &nbsp;·&nbsp; {today_str}
          </div>
        </div>""", unsafe_allow_html=True)

        # FIX 5: Output filename = Activity Name
        safe_name = re.sub(r'[^\w\s\-]', '', activity_name).strip().replace(' ', '_')[:80]
        _dl_key   = f"dl_{abs(hash(safe_name + today_str)) % 10_000_000}"
        st.download_button(
            label="📥  Download MOP Document",
            data=io.BytesIO(output_bytes),
            file_name=f"{safe_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=_dl_key,
            use_container_width=True,
        )

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

        with st.expander("📋 Preview extracted content"):
            for key in SECTION_KEYS[:-1]:
                content = sections.get(key, [])
                label   = SECTION_LABELS[key]
                if content:
                    st.markdown(f"**{label}**")
                    for item in content[:3]:
                        line = item["text"] if isinstance(item, dict) else item
                        st.markdown(
                            f"<span style='color:#003366;font-size:.75rem;'>→ {line[:130]}</span>",
                            unsafe_allow_html=True,
                        )
                    if len(content) > 3:
                        st.caption(f"… +{len(content) - 3} more lines")
                else:
                    st.markdown(
                        f"<span style='color:#9aaab8;font-size:.74rem;'>○ {label} — empty</span>",
                        unsafe_allow_html=True,
                    )

    elif gen_btn:
        st.warning("⚠ Upload a Solution Document and ensure a template is available.")

    if not st.session_state.get("output_bytes"):
        st.markdown("""
        <div class="eri-card" style="border:2px dashed #dde4ed;background:#fafbfc;min-height:340px;
          display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;">
          <div style="font-size:3rem;margin-bottom:1rem;">📄</div>
          <div style="font-size:.95rem;font-weight:700;color:#003366;margin-bottom:.4rem;">
            Output will appear here
          </div>
          <div style="font-size:.76rem;color:#9aaab8;max-width:240px;line-height:1.6;">
            Upload your solution document and click Generate MOP to get started.
          </div>
        </div>
        """, unsafe_allow_html=True)

# ── Footer ────────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
  🔒 No data stored &nbsp;·&nbsp; In-memory processing only &nbsp;·&nbsp;
  Session cleared on close &nbsp;·&nbsp; Smart MOP Generator v7 &nbsp;·&nbsp; Ericsson Internal
</div>
""", unsafe_allow_html=True)
