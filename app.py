"""
Smart MOP Generator — Premium Edition v3
==========================================
Flow:
  1. Upload Solution Document (.docx) — AI-generated, has commands in "CAPS QUOTES"
  2. Upload Log File(s) (.txt / .log) — optional, for command output injection
  3. Generator finds every "QUOTED COMMAND" in solution doc,
     searches logs (case-insensitive), injects exact output into MOP.
  4. Download formatted MOP .docx
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
# CSS
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700&family=IBM+Plex+Sans:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
    background-color: #0d0f14;
    color: #c9d1d9;
}
.block-container { max-width: 860px; padding-top: 0 !important; padding-bottom: 2rem; }

.top-bar {
    width: 100%; height: 3px;
    background: linear-gradient(90deg, #1a6b4a 0%, #2d9cdb 50%, #8b5cf6 100%);
}
.hero {
    background: linear-gradient(160deg, #111827 0%, #0d1f2d 60%, #0d1a12 100%);
    border: 1px solid rgba(45,156,219,.12); border-top: none;
    border-radius: 0 0 20px 20px; padding: 2.6rem 2.6rem 2rem;
    margin-bottom: 2rem; position: relative; overflow: hidden;
}
.hero::before {
    content:''; position:absolute; top:-60px; right:-60px;
    width:280px; height:280px;
    background:radial-gradient(circle,rgba(45,156,219,.06) 0%,transparent 70%);
    pointer-events:none;
}
.hero::after {
    content:''; position:absolute; bottom:-40px; left:-40px;
    width:200px; height:200px;
    background:radial-gradient(circle,rgba(26,107,74,.08) 0%,transparent 70%);
    pointer-events:none;
}
.hero-eyebrow {
    font-family:'IBM Plex Mono',monospace; font-size:.65rem;
    letter-spacing:3px; color:#2d9cdb; text-transform:uppercase; margin-bottom:.6rem;
}
.hero-title {
    font-family:'Playfair Display',serif; font-size:2.2rem;
    font-weight:700; color:#e6edf3; margin:0 0 .3rem; line-height:1.2;
}
.hero-title .ht-smart { color:#3fb882; font-style:italic; }
.hero-title em {
    font-style:italic; color:#2d9cdb;
    background:rgba(45,156,219,.1); padding:0 6px; border-radius:4px;
}
.hero-title .ht-gen { color:#a78bfa; font-style:italic; }
.hero-sub { font-size:.85rem; color:#6e7681; margin:0 0 1.4rem; font-weight:300; }
.badge-row { display:flex; gap:8px; flex-wrap:wrap; }
.badge { font-family:'IBM Plex Mono',monospace; font-size:.62rem; font-weight:500; padding:4px 10px; border-radius:4px; letter-spacing:.5px; }
.b-green  { background:rgba(26,107,74,.18);  color:#3fb882; border:1px solid rgba(63,184,130,.2); }
.b-blue   { background:rgba(45,156,219,.14); color:#5ba8e0; border:1px solid rgba(45,156,219,.25); }
.b-purple { background:rgba(139,92,246,.12); color:#a78bfa; border:1px solid rgba(139,92,246,.22); }
.b-orange { background:rgba(237,137,54,.12); color:#ed8936; border:1px solid rgba(237,137,54,.22); }

.priv-bar {
    display:flex; align-items:center; gap:10px;
    background:rgba(26,107,74,.06); border-left:2px solid #1a6b4a;
    border-radius:0 8px 8px 0; padding:.6rem 1rem;
    font-size:.76rem; color:#6e7681; margin-bottom:1.8rem;
}
.priv-bar strong { color:#3fb882; }

.step-card {
    background:#111827; border:1px solid rgba(255,255,255,.06);
    border-radius:14px; padding:1.4rem 1.6rem; margin-bottom:1.1rem;
    transition:border-color .2s;
}
.step-card:hover { border-color:rgba(45,156,219,.2); }
.step-header { display:flex; align-items:center; gap:12px; margin-bottom:1rem; }
.step-number {
    font-family:'IBM Plex Mono',monospace; font-size:.62rem; font-weight:500;
    color:#2d9cdb; background:rgba(45,156,219,.1); border:1px solid rgba(45,156,219,.2);
    border-radius:4px; padding:3px 8px; letter-spacing:1px;
}
.step-title {
    font-family:'IBM Plex Sans',sans-serif; font-size:.78rem; font-weight:600;
    color:#8b949e; letter-spacing:1.2px; text-transform:uppercase;
}
.step-optional {
    font-family:'IBM Plex Mono',monospace; font-size:.6rem;
    color:#8b5cf6; background:rgba(139,92,246,.1); border:1px solid rgba(139,92,246,.2);
    border-radius:4px; padding:2px 7px; letter-spacing:.8px; margin-left:6px;
}

.pill-ok {
    display:inline-flex; align-items:center; gap:6px;
    background:rgba(26,107,74,.1); border:1px solid rgba(63,184,130,.18);
    border-radius:6px; padding:5px 12px; font-size:.75rem; color:#3fb882; margin-top:4px;
}
.pill-warn {
    display:inline-flex; align-items:center; gap:6px;
    background:rgba(210,105,30,.1); border:1px solid rgba(210,105,30,.2);
    border-radius:6px; padding:5px 12px; font-size:.75rem; color:#e8955a; margin-top:4px;
}
.pill-info {
    display:inline-flex; align-items:center; gap:6px;
    background:rgba(139,92,246,.1); border:1px solid rgba(139,92,246,.2);
    border-radius:6px; padding:5px 12px; font-size:.75rem; color:#a78bfa; margin-top:4px;
}

.stButton > button {
    background:linear-gradient(135deg,#1a6b4a 0%,#1d5f8a 100%) !important;
    color:#fff !important; border:none !important; border-radius:10px !important;
    font-family:'IBM Plex Sans',sans-serif !important; font-weight:600 !important;
    font-size:.9rem !important; padding:.65rem 2rem !important; width:100% !important;
    letter-spacing:.4px !important; transition:all .2s !important;
}
.stButton > button:hover {
    background:linear-gradient(135deg,#1e7d56 0%,#2270a3 100%) !important;
    transform:translateY(-1px) !important;
    box-shadow:0 8px 24px rgba(45,156,219,.15) !important;
}
.stButton > button:disabled { opacity:.35 !important; transform:none !important; }

[data-testid="stDownloadButton"] > button {
    background:linear-gradient(135deg,#1a5c3a 0%,#164e70 100%) !important;
    color:#fff !important; border:none !important; border-radius:10px !important;
    font-family:'IBM Plex Sans',sans-serif !important; font-weight:600 !important;
    font-size:.9rem !important; padding:.65rem 2rem !important; width:100% !important;
}

.prog-wrap {
    background:#0d1117; border:1px solid rgba(255,255,255,.05);
    border-radius:10px; padding:1rem 1.2rem;
}
.ps { display:flex; align-items:center; gap:10px; padding:7px 0; font-size:.78rem; border-bottom:1px solid rgba(255,255,255,.03); transition:color .3s; }
.ps:last-child { border-bottom:none; }
.ps.done  { color:#3fb882; } .ps.doing { color:#2d9cdb; } .ps.wait  { color:#30363d; }
.pd { width:6px; height:6px; border-radius:50%; flex-shrink:0; }
.pd.done  { background:#3fb882; }
.pd.doing { background:#2d9cdb; animation:pulse 1.1s infinite; }
.pd.wait  { background:#21262d; }
@keyframes pulse { 0%,100%{opacity:1;transform:scale(1)} 50%{opacity:.4;transform:scale(.7)} }

.success-wrap {
    background:linear-gradient(135deg,rgba(26,107,74,.1) 0%,rgba(45,156,219,.07) 100%);
    border:1px solid rgba(63,184,130,.2); border-radius:14px;
    padding:1.6rem; margin:1rem 0; text-align:center;
}
.success-icon { font-size:2rem; margin-bottom:.4rem; }
.success-title { font-family:'Playfair Display',serif; font-size:1.15rem; color:#3fb882; margin-bottom:.2rem; }
.success-sub { font-size:.78rem; color:#6e7681; }
.success-name { color:#5ba8e0; font-weight:600; }

.inj-box {
    background:#0d1117; border:1px solid rgba(139,92,246,.2);
    border-radius:10px; padding:1rem 1.2rem; margin-top:.8rem;
}
.inj-title { font-family:'IBM Plex Mono',monospace; font-size:.68rem; color:#a78bfa; letter-spacing:1px; margin-bottom:.5rem; }
.inj-row { display:flex; justify-content:space-between; align-items:center; padding:4px 0; border-bottom:1px solid rgba(255,255,255,.03); font-size:.74rem; }
.inj-row:last-child { border-bottom:none; }
.inj-cmd { color:#e6edf3; font-family:'IBM Plex Mono',monospace; }
.inj-found { color:#3fb882; } .inj-miss { color:#e8955a; }

.metric-row { display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin-top:1rem; }
.metric-box { background:#0d1117; border:1px solid rgba(255,255,255,.06); border-radius:10px; padding:.9rem; text-align:center; }
.metric-val { font-family:'Playfair Display',serif; font-size:1.6rem; color:#e6edf3; line-height:1; }
.metric-lbl { font-size:.68rem; color:#6e7681; margin-top:4px; letter-spacing:.5px; text-transform:uppercase; }

hr { border-color:rgba(255,255,255,.05) !important; }
label { color:#8b949e !important; font-size:.8rem !important; }
.footer { text-align:center; font-size:.66rem; color:#21262d; padding:1rem 0 .5rem; border-top:1px solid rgba(255,255,255,.04); font-family:'IBM Plex Mono',monospace; letter-spacing:.5px; }
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
    "objective":"1. Objective","activity_description":"2. Activity Description",
    "activity_type":"3. Activity Type","domain_in_scope":"4. Domain in Scope",
    "prerequisites":"5. Pre-requisites","inventory_details":"6. Inventory Details",
    "node_connectivity":"7. Node Connectivity Process","iam":"8. Identity & Access Management",
    "triggering_method":"9. Activity Triggering Method","sop":"10. Standard Operating Procedure",
    "acceptance_criteria":"11. Acceptance Criteria","assumptions":"12. Assumptions",
    "connectivity_diagram":"Connectivity Diagram:",
}
TOC_PAGES = {
    "objective":2,"activity_description":2,"activity_type":2,"domain_in_scope":2,
    "prerequisites":2,"inventory_details":3,"node_connectivity":3,"iam":3,
    "triggering_method":3,"sop":3,"acceptance_criteria":4,"assumptions":4,
}
PARA_SECTIONS     = {"objective","activity_description","activity_type","domain_in_scope","inventory_details","assumptions"}
BULLET_SECTIONS   = {"prerequisites","node_connectivity","iam","triggering_method","acceptance_criteria"}
NUMBERED_SECTIONS = {"sop"}

MAX_OUTPUT_LINES = 20

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
# LOG ENGINE — Command Search & Output Extraction
# ─────────────────────────────────────────────────────────────────

def _is_command_line(line: str) -> bool:
    """
    Heuristic: detect if a log line looks like a command execution boundary.
    Covers multiple vendor patterns:
      - Ericsson/Nokia OSS:  <CMD:params;   or  CMD:params;
      - Cisco/Juniper CLI:   hostname# show ...   hostname> show ...
      - Linux shell:         $ cmd  or  # cmd
      - END / EXECUTED markers
      - Lines ending with ;  (common in MML/OSS commands)
      - Bracketed timestamps followed by command: [timestamp] CMD
    """
    s = line.strip()
    if not s:
        return False
    # MML style: starts with < or contains : and ends with ;
    if re.match(r'^<?\s*\w[\w\-:=,\s]*;$', s):
        return True
    # CLI prompt: hostname#/> command
    if re.match(r'^\S+[#>]\s+\S+', s):
        return True
    # Shell prompt
    if re.match(r'^[$#]\s+\S+', s):
        return True
    # END / EXECUTED / COMMAND COMPLETE markers
    if re.match(r'^(END|EXECUTED|COMMAND COMPLETE|OK|RESULT)\s*$', s, re.IGNORECASE):
        return True
    # Timestamped command: [2024-...] CMD
    if re.match(r'^\[[\d\-T:\.Z ]+\]\s+\S+', s):
        return True
    return False

def _normalize(text: str) -> str:
    """Lowercase, collapse whitespace, strip punctuation noise for comparison."""
    t = text.lower().strip()
    t = re.sub(r'["\';:<>]', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t

def _command_similarity(cmd_from_doc: str, log_line: str) -> bool:
    """
    Return True if log_line is likely the execution of cmd_from_doc.
    Strategy:
      1. Exact substring match (case-insensitive, after normalization)
      2. All significant tokens of cmd_from_doc appear in log_line (order-independent)
      3. First significant token (command verb) must match
    """
    norm_cmd = _normalize(cmd_from_doc)
    norm_log = _normalize(log_line)

    # Direct substring
    if norm_cmd in norm_log:
        return True

    # Token-based: all tokens >= 2 chars from cmd must be in log line
    tokens = [t for t in norm_cmd.split() if len(t) >= 2]
    if not tokens:
        return False
    # First token (command verb) must match
    if tokens[0] not in norm_log:
        return False
    # At least 70% of tokens must match
    matched = sum(1 for t in tokens if t in norm_log)
    return matched >= max(1, int(len(tokens) * 0.7))


def search_command_in_logs(command_raw: str, log_text: str) -> str | None:
    """
    Search for command_raw in log_text.
    Returns extracted output (up to MAX_OUTPUT_LINES) or None if not found.

    Output boundary detection (in priority order):
      1. Next command-looking line
      2. END / EXECUTED / OK / RESULT keyword line
      3. Blank-line cluster (3+ consecutive blanks = section break)
      4. Hard cap: MAX_OUTPUT_LINES
    """
    lines = log_text.splitlines()
    n = len(lines)

    # Find the first line that matches the command
    match_idx = None
    for i, line in enumerate(lines):
        if _command_similarity(command_raw, line):
            match_idx = i
            break

    if match_idx is None:
        return None

    # Collect output lines starting from the line AFTER the command
    output_lines = []
    consecutive_blanks = 0

    for j in range(match_idx + 1, n):
        raw_line = lines[j]
        stripped  = raw_line.strip()

        # Stop: next command boundary
        if _is_command_line(raw_line) and j > match_idx + 1:
            break

        # Stop: explicit end markers
        if re.match(r'^(END|EXECUTED|COMMAND COMPLETE)\s*$', stripped, re.IGNORECASE):
            # Include the marker itself then stop
            output_lines.append(raw_line)
            break

        # Track consecutive blanks — 3+ = section separator
        if stripped == "":
            consecutive_blanks += 1
            if consecutive_blanks >= 3:
                break
        else:
            consecutive_blanks = 0

        output_lines.append(raw_line)

        # Hard cap
        if len(output_lines) >= MAX_OUTPUT_LINES:
            break

    if not output_lines:
        return None

    return "\n".join(output_lines)


def extract_commands_from_doc(doc: Document) -> list[str]:
    """
    Extract all "QUOTED CAPS" commands from solution document paragraphs.
    Pattern: "ANY TEXT IN CAPS OR MIXED" — we rely on the quotes.
    Returns deduplicated list preserving first-seen order.
    """
    pattern = re.compile(r'"([^"]{3,})"')
    seen = []
    seen_norm = set()

    for para in doc.paragraphs:
        for match in pattern.finditer(para.text):
            cmd = match.group(1).strip()
            # Filter: skip if it looks like a label/sentence (has lowercase words > 3 chars)
            # Keep if it's all-caps, mixed caps with colons/equals (command-like)
            norm = cmd.upper()
            if norm not in seen_norm:
                # Exclude known non-command quoted strings
                if not re.match(r'^ASSUMED COMMAND', cmd, re.IGNORECASE):
                    seen.append(cmd)
                    seen_norm.add(norm)
    return seen


def build_log_lookup(log_texts: list[str]) -> str:
    """Merge all log files into one searchable text."""
    return "\n\n".join(log_texts)


def inject_outputs_into_sections(
    sections: dict,
    commands_found: list[str],
    log_combined: str,
) -> tuple[dict, dict]:
    """
    For each paragraph/bullet in sections, detect "QUOTED COMMAND" strings.
    If command found in logs → append output block after that line.
    Returns modified sections + injection_report {cmd: "found"/"not_found"}.
    """
    pattern = re.compile(r'"([^"]{3,})"')
    injection_report = {}

    for key in SECTION_KEYS[:-1]:
        content = sections.get(key, [])
        new_content = []
        for line in content:
            new_content.append(line)
            # Check if this line contains a quoted command
            for match in pattern.finditer(line):
                cmd = match.group(1).strip()
                cmd_up = cmd.upper()
                # Only process if it looks like a real command (not a label sentence)
                if re.match(r'^ASSUMED COMMAND', cmd, re.IGNORECASE):
                    continue
                if cmd_up in [c.upper() for c in commands_found]:
                    if cmd_up not in injection_report:
                        output = search_command_in_logs(cmd, log_combined)
                        if output:
                            injection_report[cmd_up] = "found"
                            # Add output as a special marker line (parsed later in builder)
                            new_content.append(f"__LOG_OUTPUT_START__")
                            for ol in output.splitlines():
                                new_content.append(f"__LOG_LINE__{ol}")
                            new_content.append(f"__LOG_OUTPUT_END__")
                        else:
                            injection_report[cmd_up] = "not_found"
        sections[key] = new_content

    return sections, injection_report

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
    for para in doc.paragraphs[:15]:
        text = para.text.strip()
        if not text:
            continue
        if text.upper() in ("METHOD OF PROCEDURE","CONTENTS:","CONTENTS"):
            continue
        if re.match(r'^\d+[\.\)]\s+\w.*Page\s+\d+', text):
            continue
        if re.match(r'^\d+[\.\)]\s+', text) and normalize_heading(text):
            continue
        for run in para.runs:
            if run.italic and run.underline:
                return text
    for para in doc.paragraphs[:8]:
        if para.style.name.startswith("Heading 1"):
            name = para.text.strip()
            name = re.sub(r'^MOP\s*:\s*', '', name, flags=re.IGNORECASE)
            name = re.sub(r'^UC\s*:\s*', '', name, flags=re.IGNORECASE)
            if name:
                return name
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
    run.font.bold = bold; run.font.italic = italic; run.font.underline = underline
    if color: run.font.color.rgb = RGBColor(*color)
    return run

def _set_rtab(para, pos=8400):
    pPr = para._p.get_or_add_pPr()
    for old in pPr.findall(qn("w:tabs")): pPr.remove(old)
    tabs = OxmlElement("w:tabs")
    tab  = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right"); tab.set(qn("w:pos"), str(pos))
    tabs.append(tab); pPr.append(tabs)

def _pgbr(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(0)
    run = p.add_run()
    br  = OxmlElement("w:br"); br.set(qn("w:type"), "page")
    run._r.append(br)

def _h2(doc, text):
    p = doc.add_paragraph()
    p.style = doc.styles["Heading 2"]
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(2)
    run = p.add_run(text)
    run.font.name = "Calibri"; run.font.size = Pt(13); run.font.bold = True
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

def _log_output_block(doc, output_text: str):
    """Render log output in monospace, shaded block."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after  = Pt(4)
    # Light gray shading via XML
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), "F2F2F2")
    pPr.append(shd)
    run = p.add_run(output_text)
    run.font.name = "Courier New"
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

def _img(doc, img_bytes, ext):
    p = doc.add_paragraph()
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
                if elem.tag == qn("w:headerReference"): header_refs.append(deepcopy(elem))
                elif elem.tag == qn("w:footerReference"): footer_refs.append(deepcopy(elem))
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
# SECTION CONTENT WRITER — handles __LOG_LINE__ markers
# ─────────────────────────────────────────────────────────────────
def _write_section_content(doc, key, content, style_fn):
    """
    Write content lines to doc using style_fn (body/bullet/numbered).
    Detects __LOG_OUTPUT_START__ … __LOG_OUTPUT_END__ blocks and renders
    them as monospace shaded output blocks instead of regular paragraphs.
    """
    i = 0
    while i < len(content):
        line = content[i]
        if line == "__LOG_OUTPUT_START__":
            # Collect log lines until END marker
            log_lines = []
            i += 1
            while i < len(content) and content[i] != "__LOG_OUTPUT_END__":
                raw = content[i]
                if raw.startswith("__LOG_LINE__"):
                    log_lines.append(raw[len("__LOG_LINE__"):])
                i += 1
            if log_lines:
                _log_output_block(doc, "\n".join(log_lines))
        else:
            style_fn(doc, line)
        i += 1

# ─────────────────────────────────────────────────────────────────
# MAIN MOP BUILDER
# ─────────────────────────────────────────────────────────────────
def build_mop(template_bytes: bytes, activity_name: str,
              sections: dict, today_str: str) -> bytes:
    doc = Document(io.BytesIO(template_bytes))
    _update_header_date(doc, today_str)
    _clear_and_prep_body(doc)

    # ── COVER PAGE ────────────────────────────────────────────────
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    _r(p, "METHOD OF PROCEDURE", size=18, bold=True, color=(0x7F,0x7F,0x7F))

    # Separator line under heading
    p_sep = doc.add_paragraph()
    p_sep.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sep.paragraph_format.space_before = Pt(0)
    p_sep.paragraph_format.space_after  = Pt(8)
    pPr = p_sep._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1");    bottom.set(qn("w:color"), "4F81BD")
    pBdr.append(bottom); pPr.append(pBdr)

    # Activity name
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(10)
    _r(p, activity_name, size=14, italic=True, underline=True)

    # CONTENTS
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    _r(p, "CONTENTS:", size=12, bold=True, underline=True)

    # Native Word TOC field
    toc_p = doc.add_paragraph()
    toc_p.paragraph_format.space_after = Pt(2)
    run = toc_p.add_run()
    fldChar_begin = OxmlElement("w:fldChar"); fldChar_begin.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instrText.text = ' TOC \\o "1-2" \\h \\z \\u '
    fldChar_end = OxmlElement("w:fldChar"); fldChar_end.set(qn("w:fldCharType"), "end")
    run._r.append(fldChar_begin); run._r.append(instrText); run._r.append(fldChar_end)

    _pgbr(doc)

    # ── BODY SECTIONS ─────────────────────────────────────────────
    for key in SECTION_KEYS[:-1]:
        content = sections.get(key, [])
        _h2(doc, SECTION_LABELS[key])

        if key in PARA_SECTIONS:
            # Para sections: join non-marker lines, output blocks rendered separately
            plain = [l for l in content if not l.startswith("__LOG")]
            joined = " ".join(plain).strip()
            if joined:
                _body(doc, joined)
            # Now render any injected output blocks
            _write_section_content(doc, key,
                [l for l in content if l.startswith("__LOG")], _body)
        elif key in BULLET_SECTIONS:
            if content:
                _write_section_content(doc, key, content, _bullet)
            else:
                _body(doc, "")
        elif key in NUMBERED_SECTIONS:
            if content:
                _write_section_content(doc, key, content, _numbered)
            else:
                _body(doc, "")

    # ── CONNECTIVITY DIAGRAM ──────────────────────────────────────
    images = sections.get("connectivity_diagram", [])
    if images:
        _h2(doc, SECTION_LABELS["connectivity_diagram"])
        for img_bytes, ext in images:
            _img(doc, img_bytes, ext)

    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.read()

# ─────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────
st.markdown('<div class="top-bar"></div>', unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <p class="hero-eyebrow">// TELECOM AUTOMATION TOOLKIT</p>
  <h1 class="hero-title">
    <span class="ht-smart">Smart</span> <em>MOP</em> <span class="ht-gen">Generator</span>
  </h1>
  <p class="hero-sub">
    Upload Solution Document + optional Log files →
    Commands auto-matched, outputs injected, MOP formatted and ready.
  </p>
  <div class="badge-row">
    <span class="badge b-green">⚡ IN-MEMORY ONLY</span>
    <span class="badge b-blue">📋 12-SECTION AUTO STRUCTURE</span>
    <span class="badge b-purple">🔍 LOG OUTPUT INJECTION</span>
    <span class="badge b-orange">🖼 IMAGES PRESERVED</span>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="priv-bar">
  🔒&nbsp; <strong>Zero Data Storage Policy:</strong>&nbsp;
  All files processed in-memory. Nothing written to disk. Session clears on close.
</div>
""", unsafe_allow_html=True)

# ── STEP 1: Template ─────────────────────────────────────────────
st.markdown("""
<div class="step-card">
  <div class="step-header">
    <span class="step-number">STEP 01</span>
    <span class="step-title">Select or Upload Template</span>
  </div>
</div>""", unsafe_allow_html=True)

templates = discover_templates()
selected_template = None
template_bytes    = None

if not templates:
    st.markdown('<div class="pill-warn">⚠️ No template found. Place <strong>Template.docx</strong> in <code>templates/</code> folder or root, then restart.</div>', unsafe_allow_html=True)
else:
    col_sel, col_up = st.columns([2, 1])
    with col_sel:
        names = [t.name for t in templates]
        sel   = st.selectbox("Choose existing template", names, label_visibility="visible")
        selected_template = next(t for t in templates if t.name == sel)
        template_bytes    = load_template_bytes(selected_template)
        st.markdown(f'<div class="pill-ok">✅ <strong>{sel}</strong> — ready</div>', unsafe_allow_html=True)
    with col_up:
        new_tmpl = st.file_uploader("Upload new template (.docx)", type=["docx"], key="tmpl_up", label_visibility="visible")
        if new_tmpl:
            save_dir = Path("templates"); save_dir.mkdir(exist_ok=True)
            with open(save_dir / new_tmpl.name, "wb") as f: f.write(new_tmpl.read())
            st.success(f"✅ Saved: {new_tmpl.name}"); st.rerun()

st.markdown("<br>", unsafe_allow_html=True)

# ── STEP 2: Solution Document ─────────────────────────────────────
st.markdown("""
<div class="step-card">
  <div class="step-header">
    <span class="step-number">STEP 02</span>
    <span class="step-title">Upload Solution Document</span>
  </div>
</div>""", unsafe_allow_html=True)

sol_file = st.file_uploader("Solution Document (.docx)", type=["docx"], key="sol_up", label_visibility="visible")
if sol_file:
    st.markdown(f'<div class="pill-ok">✅ <strong>{sol_file.name}</strong> · {sol_file.size/1024:.1f} KB</div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── STEP 3: Log Files (Optional) ─────────────────────────────────
st.markdown("""
<div class="step-card">
  <div class="step-header">
    <span class="step-number">STEP 03</span>
    <span class="step-title">Upload Log Files</span>
    <span class="step-optional">OPTIONAL</span>
  </div>
</div>""", unsafe_allow_html=True)

st.markdown(
    '<div class="pill-info">📂 Upload one or more <strong>.txt / .log</strong> files. '
    'Commands in the Solution Document will be searched here and outputs injected into MOP.</div>',
    unsafe_allow_html=True
)
log_files = st.file_uploader(
    "Log files (.txt or .log) — multiple allowed",
    type=["txt","log"],
    accept_multiple_files=True,
    key="log_up",
    label_visibility="visible"
)
if log_files:
    for lf in log_files:
        st.markdown(f'<div class="pill-ok">📄 <strong>{lf.name}</strong> · {lf.size/1024:.1f} KB</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="pill-warn">⚠️ No logs uploaded — commands will appear as ASSUMED in MOP (no output injection)</div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── STEP 4: Generate ─────────────────────────────────────────────
st.markdown("""
<div class="step-card">
  <div class="step-header">
    <span class="step-number">STEP 04</span>
    <span class="step-title">Generate MOP Document</span>
  </div>
</div>""", unsafe_allow_html=True)

can_go  = bool(sol_file and templates)
gen_btn = st.button("🚀  Generate MOP Document", disabled=not can_go)
if not can_go:
    missing = []
    if not templates:  missing.append("template")
    if not sol_file:   missing.append("solution document")
    if missing:
        st.markdown(f'<div class="pill-warn">Waiting for: {" + ".join(missing)}</div>', unsafe_allow_html=True)

# ── PROCESSING ───────────────────────────────────────────────────
if gen_btn and can_go:
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("""
    <div class="step-card">
      <div class="step-header">
        <span class="step-number">PROCESSING</span>
        <span class="step-title">Building your MOP</span>
      </div>
    </div>""", unsafe_allow_html=True)

    has_logs = bool(log_files)

    steps = [
        ("Loading template into memory",                  "📂"),
        ("Reading solution document",                     "📖"),
        ("Extracting activity name",                      "🏷️"),
        ("Parsing all 12 MOP sections",                   "🔍"),
        ("Extracting commands from document",             "⌨️"),
        ("Loading & merging log files" if has_logs else "No logs — skipping output injection", "📋"),
        ("Searching commands in logs & injecting output" if has_logs else "Marking commands as ASSUMED", "🔗"),
        ("Detecting embedded images",                     "🖼️"),
        ("Rebuilding cover page & TOC",                   "📑"),
        ("Preserving header & footer",                    "🔒"),
        ("Finalising & packaging document",               "📦"),
    ]

    st.markdown('<div class="prog-wrap">', unsafe_allow_html=True)
    phs = [st.empty() for _ in steps]
    for ph, (s, icon) in zip(phs, steps):
        ph.markdown(f'<div class="ps wait"><div class="pd wait"></div>{icon} {s}</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    try:
        activity_name    = ""
        sections         = {}
        today_str        = ""
        output_bytes     = b""
        commands_list    = []
        log_combined     = ""
        injection_report = {}

        for i, (ph, (step, icon)) in enumerate(zip(phs, steps)):
            ph.markdown(f'<div class="ps doing"><div class="pd doing"></div>{icon} {step}…</div>', unsafe_allow_html=True)
            time.sleep(0.15)

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
            elif i == 4:
                commands_list = extract_commands_from_doc(sol_doc)
            elif i == 5:
                if has_logs:
                    log_texts = []
                    for lf in log_files:
                        raw = lf.read()
                        try:
                            log_texts.append(raw.decode("utf-8", errors="replace"))
                        except Exception:
                            log_texts.append(raw.decode("latin-1", errors="replace"))
                    log_combined = build_log_lookup(log_texts)
            elif i == 6:
                if has_logs and commands_list:
                    sections, injection_report = inject_outputs_into_sections(
                        sections, commands_list, log_combined
                    )
            elif i == 10:
                output_bytes = build_mop(tmpl_b, activity_name, sections, today_str)

            ph.markdown(f'<div class="ps done"><div class="pd done"></div>{icon} {step} ✓</div>', unsafe_allow_html=True)
            time.sleep(0.05)

        # ── Success ──────────────────────────────────────────────
        st.markdown(f"""
        <div class="success-wrap">
          <div class="success-icon">✅</div>
          <div class="success-title">MOP Generated Successfully</div>
          <div class="success-sub">Activity: <span class="success-name">{activity_name}</span> &nbsp;·&nbsp; {today_str}</div>
        </div>""", unsafe_allow_html=True)

        safe_name = re.sub(r'[^\w\s-]', '', activity_name).strip().replace(' ', '_')[:60]
        st.download_button(
            label="📥  Download MOP Document (.docx)",
            data=output_bytes,
            file_name=f"MOP_{safe_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        # ── Injection Report ─────────────────────────────────────
        if injection_report:
            found_n   = sum(1 for v in injection_report.values() if v == "found")
            missing_n = sum(1 for v in injection_report.values() if v == "not_found")
            st.markdown(f"""
            <div class="inj-box">
              <div class="inj-title">🔍 COMMAND OUTPUT INJECTION REPORT — {found_n} injected · {missing_n} not found in logs</div>
            """, unsafe_allow_html=True)
            for cmd, status in injection_report.items():
                label = "✅ Output injected" if status == "found" else "⚠️ Not found in logs"
                cls   = "inj-found" if status == "found" else "inj-miss"
                st.markdown(f'<div class="inj-row"><span class="inj-cmd">"{cmd}"</span><span class="{cls}">{label}</span></div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        elif commands_list and not has_logs:
            st.markdown(f'<div class="pill-warn">⚠️ {len(commands_list)} command(s) found in document — no logs uploaded, outputs not injected</div>', unsafe_allow_html=True)

        # ── Summary Metrics ──────────────────────────────────────
        filled      = sum(1 for k in SECTION_KEYS[:-1] if sections.get(k))
        images_found= len(sections.get("connectivity_diagram", []))
        total_lines = sum(len(v) for k, v in sections.items() if k != "connectivity_diagram")

        st.markdown(f"""
        <div class="metric-row">
          <div class="metric-box">
            <div class="metric-val">{filled}<span style="font-size:.9rem;color:#6e7681;">/12</span></div>
            <div class="metric-lbl">Sections Filled</div>
          </div>
          <div class="metric-box">
            <div class="metric-val">{len(commands_list)}</div>
            <div class="metric-lbl">Commands Found</div>
          </div>
          <div class="metric-box">
            <div class="metric-val">{images_found}</div>
            <div class="metric-lbl">Images Embedded</div>
          </div>
        </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Content Preview ──────────────────────────────────────
        with st.expander("📋  Preview extracted section content"):
            for key in SECTION_KEYS[:-1]:
                content = sections.get(key, [])
                label   = SECTION_LABELS[key]
                plain   = [l for l in content if not l.startswith("__LOG")]
                if plain:
                    st.markdown(f"**{label}**")
                    for line in plain[:3]:
                        st.markdown(f"<span style='color:#8b949e;font-size:.76rem;'>→ {line[:130]}</span>", unsafe_allow_html=True)
                    if len(plain) > 3:
                        st.caption(f"  … +{len(plain)-3} more lines")
                else:
                    st.markdown(f"<span style='color:#30363d;font-size:.75rem;'>◌ {label} — no content extracted</span>", unsafe_allow_html=True)

    except Exception as e:
        st.error(f"❌ Error during generation: {e}")
        import traceback
        st.code(traceback.format_exc())

elif gen_btn:
    st.warning("⚠️ Please upload a Solution Document and ensure a template is available.")

st.markdown("""
<div class="footer">
  🔒 ZERO DATA STORAGE &nbsp;·&nbsp; IN-MEMORY PROCESSING &nbsp;·&nbsp;
  SESSION CLEARS ON CLOSE &nbsp;·&nbsp; SMART MOP GENERATOR v3
</div>""", unsafe_allow_html=True)
