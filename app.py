"""
Smart MOP Generator — v6
=========================
CORE FIX: Template-preserving approach.
- Load template as-is (keeps all header, footer, styles, TOC, revision table)
- Find each Heading 1 in template, clear content below it
- Inject content from solution document into correct heading section
- Activity name from solution doc (heading just before first Objective heading)
- Download filename = activity name
- Only Automation_MOP_Template.docx shown in template selector
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
from docx.shared import Pt, RGBColor, Inches
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

# ── Session state init ──
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
# CSS — Ericsson corporate palette
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Lato:wght@300;400;700;900&family=Source+Code+Pro:wght@400;600&display=swap');

html, body, [class*="css"] {
  font-family: 'Lato', sans-serif;
  background-color: #F4F6F9;
  color: #1A1A2E;
}
.block-container { padding-top: 1.5rem !important; padding-bottom: 2rem; max-width: 100%; }

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
.eri-logo-text { font-family:'Lato',sans-serif; font-weight:900; font-size:1.5rem; letter-spacing:3px; color:#0082C8; text-transform:uppercase; }
.eri-logo-sub  { font-size:0.7rem; letter-spacing:1.5px; color:rgba(255,255,255,0.4); text-transform:uppercase; }
.eri-app-title { font-size:1.15rem; font-weight:700; color:#ffffff; letter-spacing:0.3px; }
.eri-app-sub   { font-size:0.72rem; color:rgba(255,255,255,0.45); letter-spacing:0.5px; margin-top:2px; }
.eri-version   { background:rgba(0,130,200,0.15); border:1px solid rgba(0,130,200,0.3); border-radius:20px; padding:3px 12px; font-size:0.65rem; color:#0082C8; font-weight:700; letter-spacing:1px; }

.priv-bar { background:rgba(0,64,128,0.06); border:1px solid rgba(0,130,200,0.18); border-left:4px solid #0082C8; border-radius:0 8px 8px 0; padding:0.6rem 1rem; font-size:0.76rem; color:#003366; margin-bottom:1.2rem; }
.priv-bar strong { color:#0082C8; }

.eri-card { background:#ffffff; border:1px solid #dde4ed; border-radius:12px; padding:1.4rem 1.6rem; margin-bottom:1rem; box-shadow:0 2px 8px rgba(0,51,102,0.06); transition:box-shadow 0.2s,border-color 0.2s; }
.eri-card:hover { border-color:#0082C8; box-shadow:0 4px 16px rgba(0,130,200,0.10); }
.eri-card-title { font-size:0.68rem; font-weight:900; letter-spacing:1.8px; text-transform:uppercase; color:#003366; margin-bottom:0.9rem; display:flex; align-items:center; gap:8px; }
.step-badge { background:#003366; color:#ffffff; font-size:0.58rem; font-weight:700; padding:2px 8px; border-radius:4px; letter-spacing:0.5px; }

.pill-ok   { display:inline-flex; align-items:center; gap:6px; background:rgba(0,150,80,0.08); border:1px solid rgba(0,150,80,0.22); border-radius:6px; padding:5px 12px; font-size:0.76rem; color:#006633; margin:3px 0; }
.pill-warn { display:inline-flex; align-items:center; gap:6px; background:rgba(200,80,0,0.07); border:1px solid rgba(200,80,0,0.2); border-radius:6px; padding:5px 12px; font-size:0.76rem; color:#7a3500; margin:3px 0; }

.stButton > button { background:linear-gradient(135deg,#003366,#0082C8)!important; color:#ffffff!important; border:none!important; border-radius:8px!important; font-family:'Lato',sans-serif!important; font-weight:700!important; font-size:0.9rem!important; padding:0.6rem 2rem!important; width:100%!important; letter-spacing:0.4px!important; transition:all 0.2s!important; }
.stButton > button:hover { background:linear-gradient(135deg,#004080,#0099E6)!important; transform:translateY(-1px)!important; box-shadow:0 6px 20px rgba(0,130,200,0.28)!important; }
.stButton > button:disabled { opacity:0.38!important; transform:none!important; }
[data-testid="stDownloadButton"] > button { background:linear-gradient(135deg,#00527A,#006FA8)!important; color:#ffffff!important; border:none!important; border-radius:8px!important; font-family:'Lato',sans-serif!important; font-weight:700!important; font-size:0.9rem!important; padding:0.6rem 2rem!important; width:100%!important; }
[data-testid="stDownloadButton"] > button:hover { background:linear-gradient(135deg,#006090,#0082C8)!important; transform:translateY(-1px)!important; }

.prog-wrap { background:#f8fafc; border:1px solid #dde4ed; border-radius:10px; padding:1rem 1.2rem; }
.ps { display:flex; align-items:center; gap:10px; padding:6px 0; font-size:0.78rem; border-bottom:1px solid rgba(0,0,0,0.04); }
.ps:last-child { border-bottom:none; }
.ps.done  { color:#006633; font-weight:600; }
.ps.doing { color:#0082C8; font-weight:600; }
.ps.wait  { color:#9aaab8; }
.pd { width:8px; height:8px; border-radius:50%; flex-shrink:0; }
.pd.done  { background:#00a050; }
.pd.doing { background:#0082C8; animation:pulse 1s infinite; }
.pd.wait  { background:#c8d4e0; }
@keyframes pulse { 0%,100%{opacity:1;transform:scale(1)} 50%{opacity:.5;transform:scale(0.85)} }

.success-card { background:linear-gradient(135deg,rgba(0,51,102,0.04),rgba(0,130,200,0.06)); border:1px solid rgba(0,130,200,0.25); border-top:4px solid #0082C8; border-radius:12px; padding:1.6rem; margin:0.8rem 0; text-align:center; }
.success-icon  { font-size:2.2rem; margin-bottom:0.4rem; }
.success-title { font-size:1.1rem; font-weight:900; color:#003366; margin-bottom:0.25rem; }
.success-sub   { font-size:0.78rem; color:#0082C8; }
.success-name  { color:#003366; font-weight:700; }

.metric-row { display:grid; grid-template-columns:repeat(3,1fr); gap:12px; margin-top:1rem; }
.metric-box { background:#f0f5fb; border:1px solid #dde4ed; border-radius:10px; padding:1rem; text-align:center; }
.metric-val { font-family:'Lato',sans-serif; font-size:1.7rem; font-weight:900; color:#003366; }
.metric-sub { font-size:0.58rem; color:#9aaab8; margin-top:1px; font-weight:700; letter-spacing:.6px; text-transform:uppercase; }

.sidebar-info { background:rgba(0,130,200,0.08); border:1px solid rgba(0,130,200,0.18); border-radius:8px; padding:0.8rem 1rem; font-size:0.74rem; color:#b8d4f0!important; margin:0.5rem 0; }
.sidebar-logo { text-align:center; padding:1.2rem 0 0.8rem; }
.eri-big-logo { font-family:'Lato',sans-serif; font-weight:900; font-size:2rem; letter-spacing:6px; color:#0082C8; display:block; }
.eri-tagline  { font-size:0.6rem; letter-spacing:2px; color:rgba(255,255,255,0.3); text-transform:uppercase; display:block; margin-top:2px; }

hr { border-color:rgba(0,51,102,0.1)!important; }
.stFileUploader label { color:#003366!important; font-weight:600!important; font-size:0.82rem!important; }
.stSelectbox label   { color:#003366!important; font-weight:600!important; font-size:0.82rem!important; }
.footer { text-align:center; font-size:0.65rem; color:#9aaab8; padding:1rem 0 0.4rem; border-top:1px solid #dde4ed; letter-spacing:0.5px; margin-top:1rem; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────

# Maps internal key → list of substrings that identify that heading in template/solution doc
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

# The one allowed template name
ALLOWED_TEMPLATE = "Automation_MOP_Template.docx"

# ─────────────────────────────────────────────────────────────────
# TEMPLATE DISCOVERY — only Automation_MOP_Template.docx
# ─────────────────────────────────────────────────────────────────
def discover_templates() -> list:
    """Find only Automation_MOP_Template.docx in templates/ or root."""
    candidates = []
    tmpl_dir = Path("templates")
    if tmpl_dir.exists():
        candidates += list(tmpl_dir.glob("*.docx"))
    candidates += [p for p in Path(".").glob("*.docx")
                   if p.name not in [c.name for c in candidates]]
    # Filter: only the allowed template name
    found = [p for p in candidates if p.name == ALLOWED_TEMPLATE]
    return found

def load_template_bytes(path: Path) -> bytes:
    with open(path, "rb") as f:
        return f.read()

# ─────────────────────────────────────────────────────────────────
# HEADING NORMALIZER
# ─────────────────────────────────────────────────────────────────
def normalize_heading(text: str):
    """Strip leading numbers/bullets, lowercase, match against HEADING_MAP."""
    t = re.sub(r'^\d+[\.\)]\s*', '', text).strip().lower()
    t = re.sub(r'\s+', ' ', t)
    for key, aliases in HEADING_MAP.items():
        for alias in aliases:
            if alias in t:
                return key
    return None

# ─────────────────────────────────────────────────────────────────
# ACTIVITY NAME EXTRACTOR
# ─────────────────────────────────────────────────────────────────
def extract_activity_name(doc: Document, sol_file_stem: str = "") -> str:
    """
    Activity name = paragraph just before the first Heading-1 'Objective'
    in the solution document.
    Falls back to: first Heading 1 → italic+underline → sol_file_stem → 'Activity Name'.
    """
    paragraphs = doc.paragraphs
    obj_idx = None
    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        if not text:
            continue
        if para.style.name.startswith("Heading") and normalize_heading(text) == "objective":
            obj_idx = i
            break

    if obj_idx is not None and obj_idx > 0:
        for i in range(obj_idx - 1, -1, -1):
            text = paragraphs[i].text.strip()
            if not text:
                continue
            if text.upper() in ("METHOD OF PROCEDURE", "METHOD OF PROCEDURE (MOP)",
                                "CONTENTS:", "CONTENTS"):
                continue
            if re.match(r"^\d+[\.\)]\s+\w.*Page\s+\d+", text):
                continue
            if re.match(r"^(Customer|Header|Footer|Document|Activity Title)\s*:",
                        text, re.IGNORECASE):
                continue
            if "revision history" in text.lower():
                continue
            # Skip bare section heading names (short text that IS a section key)
            if normalize_heading(text) is not None and len(text) < 60:
                continue
            # Clean known prefixes
            name = re.sub(r"^MOP\s*:\s*", "", text, flags=re.IGNORECASE)
            name = re.sub(r"^UC\s*:\s*", "", name, flags=re.IGNORECASE)
            name = re.sub(r"^Method of Procedure\s*[\(\[]?MOP[\)\]]?\s*:\s*",
                          "", name, flags=re.IGNORECASE)
            name = re.sub(r"^Activity\s+Title\s*:\s*", "", name, flags=re.IGNORECASE)
            if name and len(name) > 3:
                return name.strip()

    # Fallback 1: first non-objective Heading 1
    for para in paragraphs[:8]:
        if para.style.name.startswith("Heading 1"):
            name = para.text.strip()
            name = re.sub(r"^MOP\s*:\s*", "", name, flags=re.IGNORECASE)
            name = re.sub(r"^UC\s*:\s*", "", name, flags=re.IGNORECASE)
            key = normalize_heading(name)
            if name and key != "objective":
                return name

    # Fallback 2: italic + underline run
    for para in paragraphs[:10]:
        for run in para.runs:
            if run.italic and run.underline and para.text.strip():
                return para.text.strip()

    # Fallback 3: use solution document filename (without extension)
    if sol_file_stem and sol_file_stem not in ("", "Activity Name"):
        return sol_file_stem.replace("_", " ").replace("-", " ").strip()

    return "Activity Name"

# ─────────────────────────────────────────────────────────────────
# SOLUTION DOC SECTION EXTRACTOR
# ─────────────────────────────────────────────────────────────────
def extract_sections(doc: Document) -> dict:
    """
    Extract content from solution document, keyed by section.
    Supports both Heading-style and Normal-style headings.
    Preserves paragraph style names for faithful re-injection.
    Each entry: {"text": str, "style": str}
    Images collected separately.
    """
    sections = {k: [] for k in SECTION_KEYS}
    sections["connectivity_diagram"] = []
    current_key = None

    # Collect image blobs
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

        # Detect heading by Heading style OR by matching text (for Normal-style docs)
        is_heading = style.startswith("Heading")
        key_from_text = normalize_heading(text) if text else None

        if is_heading or (key_from_text and len(text) < 100):
            key = normalize_heading(text) if text else None
            if key:
                current_key = key
            # Always skip heading lines regardless (matched or not)
            continue

        if current_key is None:
            continue

        # Check for images
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
        if re.match(r'^(Customer|Header|Footer|Activity Title|Document)[\s]*:',
                    text, re.IGNORECASE):
            continue
        if text == "sample...":
            continue

        if current_key in sections:
            # Store both text and original style
            sections[current_key].append({"text": text, "style": style})

    return sections

# ─────────────────────────────────────────────────────────────────
# TEMPLATE-BASED MOP BUILDER  ← CORE FIX
# ─────────────────────────────────────────────────────────────────
def build_mop(template_bytes: bytes, activity_name: str,
              sections: dict, today_str: str) -> bytes:
    """
    Load template → clear placeholder content under each Heading 1
    → inject solution doc content → preserve all template formatting.
    """
    doc = Document(io.BytesIO(template_bytes))
    body = doc.element.body
    all_paras = body.findall(f".//{qn('w:p')}")

    # ── Step 1: Update Revision History table date ──────────────
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        t = run.text.strip()
                        if t in ("DD-MM-YYYY", "17-04-2026", "{{date}}", ""):
                            if t in ("DD-MM-YYYY", "17-04-2026", "{{date}}"):
                                run.text = today_str
                        if t in ("Activity Name", "Initial draft V1.0"):
                            run.text = f"Initial draft — {activity_name}"

    # ── Step 2: Find all Heading 1 paragraphs in template body ──
    # Map: heading_key → paragraph element index in body
    body_children = list(body)

    def get_body_index(elem):
        """Get direct-child index of elem in body."""
        try:
            return body_children.index(elem)
        except ValueError:
            return -1

    # Find heading positions
    heading_positions = {}  # key -> body child index
    for child in body_children:
        if child.tag == qn("w:p"):
            style_el = child.find(f".//{qn('w:pStyle')}")
            if style_el is not None:
                style_val = style_el.get(qn("w:val"), "")
                if "Heading1" in style_val or style_val == "1":
                    text = "".join(t.text or "" for t in child.findall(f".//{qn('w:t')}")).strip()
                    key = normalize_heading(text)
                    if key and key not in heading_positions:
                        idx = get_body_index(child)
                        heading_positions[key] = idx

    # ── Step 3: Clear content between headings ───────────────────
    # For each heading, remove all body children between this heading
    # and the next heading (or end of body), except tables (revision history)
    sorted_keys = sorted(heading_positions.items(), key=lambda x: x[1])

    for i, (key, start_idx) in enumerate(sorted_keys):
        # Next heading index or end
        if i + 1 < len(sorted_keys):
            end_idx = sorted_keys[i + 1][1]
        else:
            end_idx = len(body_children)

        # Skip "Revision History" — never clear it
        heading_elem = body_children[start_idx]
        heading_text = "".join(
            t.text or "" for t in heading_elem.findall(f".//{qn('w:t')}")
        ).strip().lower()
        if "revision" in heading_text:
            continue

        # Collect elements to remove (between start+1 and end, excluding sectPr)
        to_remove = []
        for j in range(start_idx + 1, end_idx):
            if j < len(body_children):
                elem = body_children[j]
                if elem.tag == qn("w:sectPr"):
                    continue  # never remove section properties
                if elem.tag == qn("w:tbl"):
                    continue  # keep any tables (like revision table)
                to_remove.append(elem)

        for elem in to_remove:
            try:
                body.remove(elem)
            except Exception:
                pass

    # ── Step 4: Re-discover heading positions after removals ─────
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

    # ── Step 5: Inject content after each heading ────────────────
    # We insert paragraphs after each heading element in reverse order
    # to maintain correct positions

    sorted_keys = sorted(heading_positions.items(), key=lambda x: x[1], reverse=True)

    for key, heading_idx in sorted_keys:
        content_items = sections.get(key, [])
        if key == "connectivity_diagram":
            continue

        heading_elem = body_children[heading_idx]

        # Build paragraphs to insert (in reverse so first item ends up first)
        paras_to_insert = []

        if not content_items:
            # Insert one empty paragraph as placeholder
            new_p = OxmlElement("w:p")
            paras_to_insert.append(new_p)
        else:
            for item in content_items:
                text = item["text"]
                orig_style = item["style"]
                new_p = OxmlElement("w:p")

                # Set paragraph style based on original
                pPr = OxmlElement("w:pPr")
                pStyle = OxmlElement("w:pStyle")

                # Map original style to template style name
                if orig_style == "List Paragraph":
                    pStyle.set(qn("w:val"), "ListParagraph")
                elif orig_style == "List Number":
                    pStyle.set(qn("w:val"), "ListNumber")
                elif orig_style == "List Bullet":
                    pStyle.set(qn("w:val"), "ListBullet")
                elif orig_style == "Body Text":
                    pStyle.set(qn("w:val"), "BodyText")
                else:
                    pStyle.set(qn("w:val"), "Normal")

                pPr.append(pStyle)
                new_p.append(pPr)

                # Add run with text
                r = OxmlElement("w:r")
                t = OxmlElement("w:t")
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                t.text = text
                r.append(t)
                new_p.append(r)
                paras_to_insert.append(new_p)

        # Insert paragraphs after heading element
        # Find heading's position in body and insert after it
        heading_parent = heading_elem.getparent()
        if heading_parent is None:
            continue

        children_list = list(heading_parent)
        try:
            h_idx_in_parent = children_list.index(heading_elem)
        except ValueError:
            continue

        # Insert in forward order after heading
        insert_pos = h_idx_in_parent + 1
        for para_elem in paras_to_insert:
            heading_parent.insert(insert_pos, para_elem)
            insert_pos += 1

    # ── Step 6: Insert connectivity diagram images if any ────────
    images = sections.get("connectivity_diagram", [])
    if images:
        # Find or create a place at end before sectPr
        last_elem = list(body)[-1]
        if last_elem.tag == qn("w:sectPr"):
            insert_before = last_elem
        else:
            insert_before = None

        # Add heading
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
        if insert_before is not None:
            body.insert(list(body).index(insert_before), h_p)
        else:
            body.append(h_p)

        # Add images via python-docx (need to use doc.add_picture approach)
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
      Smart MOP Generator v6<br>
      Ericsson Internal Tool
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
    <div class="eri-app-sub">Solution Document → Formatted MOP · Template-Preserving · Audit-Ready</div>
  </div>
  <div><span class="eri-version">v6</span></div>
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
    st.markdown('<div class="eri-card"><div class="eri-card-title"><span class="step-badge">STEP 01</span> Select MOP Template</div>', unsafe_allow_html=True)

    templates = discover_templates()
    selected_template = None
    template_bytes    = None

    if not templates:
        st.markdown(
            f'<div class="pill-warn">⚠ Template not found. Place '
            f'<strong>{ALLOWED_TEMPLATE}</strong> in <code>templates/</code> '
            f'folder and restart.</div>',
            unsafe_allow_html=True,
        )
    else:
        sel = templates[0].name  # Only one: Automation_MOP_Template.docx
        selected_template = templates[0]
        template_bytes    = load_template_bytes(selected_template)
        st.markdown(
            f'<div class="pill-ok">✔ &nbsp;<strong>{sel}</strong> &nbsp;ready</div>',
            unsafe_allow_html=True,
        )

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
        if not templates:   missing.append("MOP template")
        if not sol_file:    missing.append("solution document")
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
            "Clearing template placeholder content",
            "Injecting content into template headings",
            "Updating revision table",
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
                    activity_name = extract_activity_name(sol_doc, Path(sol_file.name).stem)
                    today_str     = datetime.today().strftime("%d-%m-%Y")
                elif i == 3:
                    sections = extract_sections(sol_doc)
                elif i == 8:
                    output_bytes = build_mop(tmpl_b, activity_name, sections, today_str)

                ph.markdown(f'<div class="ps done"><div class="pd done"></div>{step} ✓</div>', unsafe_allow_html=True)
                time.sleep(0.04)

            # Store in session_state so download survives re-run
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

    # ── Show result panel whenever session_state has output ──────
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

        # Download filename = activity name (clean)
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
  Session cleared on close &nbsp;·&nbsp; Smart MOP Generator v6 &nbsp;·&nbsp; Ericsson Internal
</div>
""", unsafe_allow_html=True)
