"""Smart MOP Generator - Core Engine v3 (Clean)"""
import os, re, zipfile, io, shutil
from datetime import date
from copy import deepcopy
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

# ─── Heading Map ────────────────────────────────────────────
HEADING_MAP = {
    "objective": [
        "objective", "objectives", "goal", "goals", "purpose",
        "aim", "overview", "summary", "introduction",
    ],
    "activity": [
        "activity description", "activity", "work description",
        "task description", "scope of work", "activity steps",
        "activity window", "activity window/ network", "change description",
        "work summary",
    ],
    "type": [
        "activity type", "type of activity", "work type",
        "task type", "change type", "nature of activity",
    ],
    "domain": [
        "domain in scope", "domain", "scope", "domains",
        "network domain", "technology domain",
        "scope/description", "scope description",
    ],
    "prereq": [
        "pre-requisites", "prerequisites", "pre requisites",
        "preconditions", "prerequisite", "preparation",
        "pre-requisites check list", "prerequisites check list",
        "pre-checks", "pre checks", "prechecks",
    ],
    "inventory": [
        "inventory details", "inventory", "node details",
        "equipment details", "node inventory", "asset details",
        "input sheet", "node list",
    ],
    "connect": [
        "node connectivity process", "connectivity", "node connectivity",
        "connection process", "access method", "scripts used",
    ],
    "iam": [
        "identity and access management", "iam", "access management",
        "credentials", "login details", "authentication",
    ],
    "trigger": [
        "activity triggering method", "trigger", "triggering method",
        "activity trigger", "how to trigger",
    ],
    "sop": [
        "standard operating procedure", "sop",
        "operating procedure", "detailed procedure",
    ],
    "accept": [
        "acceptance criteria", "acceptance test criteria",
        "acceptance test", "acceptance",
        "uat scenarios", "uat criteria", "uat",
        "validation criteria", "success criteria",
        "sign off criteria", "sign-off criteria",
        "test acceptance", "post checks", "post-checks",
        "post requisites", "post-requisites",
    ],
    "assume": [
        "assumptions", "assumption", "assumed conditions",
        "dependencies", "notes", "caveats",
        "risk & impact", "risk and impact", "rollback", "fallback",
    ],
}

AUTO_FILL = {
    "objective": "The objective of this activity is to successfully execute the {activity} for {vendor}, ensuring minimal impact to network services and adherence to change management processes.",
    "activity":  "This document outlines the method of procedure for {activity} to be carried out by {vendor}.",
    "type":      "Planned Maintenance Activity – {activity} ({vendor})",
    "domain":    "This activity covers the network domain relevant to {activity} as scoped and agreed with {vendor}.",
    "prereq":    "Prior to initiating {activity} with {vendor}, ensure all approvals, maintenance windows, and access permissions are in place.",
    "inventory": "Node inventory for {activity} with {vendor} to be confirmed prior to execution.",
    "connect":   "Connectivity for {activity} shall be established by {vendor} per the approved access management process.",
    "iam":       "Access credentials for {activity} shall be provided by {vendor} per IAM policy.",
    "trigger":   "This activity will be triggered upon receipt of the approved change request for {activity} with {vendor}.",
    "accept":    "Acceptance criteria for {activity} with {vendor}: All services operational, no alarms, sign-off obtained.",
    "assume":    "All pre-requisites for {activity} with {vendor} are assumed met. Rollback procedure is available.",
}

JUNK_PATTERNS = [
    'this section is to provide', 'mention the',
    'provide activity detailed', 'provide fallback',
    'please include a detailed description',
]

# ─── Helpers ────────────────────────────────────────────────

def _strip_number(text):
    """Remove leading numbering like '1.', '10.', '1.2 ' from heading text."""
    return re.sub(r'^\d+(\.\d+)*[\.\)]\s*', '', text.strip())

def _normalize(text):
    text = _strip_number(text)
    return re.sub(r'\s+', ' ', text.lower()).strip()

def match_heading(raw_text):
    """Match heading text to a section key. Returns key or None."""
    norm = _normalize(raw_text)
    if not norm or len(norm) > 100:
        return None
    # 1. Exact match
    for key, syns in HEADING_MAP.items():
        for s in syns:
            if norm == s:
                return key
    # 2. Starts-with (longest wins)
    best_key, best_len = None, 0
    for key, syns in HEADING_MAP.items():
        for s in syns:
            if norm.startswith(s) and len(s) > best_len:
                best_key, best_len = key, len(s)
            if s.startswith(norm) and len(norm) > 3 and len(norm) > best_len:
                best_key, best_len = key, len(norm)
    if best_key:
        return best_key
    # 3. Substring (synonym must be >= 10 chars)
    for key, syns in HEADING_MAP.items():
        for s in syns:
            if len(s) >= 10 and s in norm and len(s) > best_len:
                best_key, best_len = key, len(s)
    return best_key

def _is_heading_style(para):
    return 'heading' in para.style.name.lower()

def _is_junk(text):
    t = text.strip().lower()
    if t.startswith('[') and t.endswith(']'):
        return True
    return any(j in t for j in JUNK_PATTERNS)

def _has_rich_content(xml):
    return 'blip' in xml or 'OLEObject' in xml or '<w:drawing>' in xml

# ─── Extraction ─────────────────────────────────────────────

def extract_docx(input_path):
    """
    Extract content from docx into section map.
    RULE: Only Heading-styled paragraphs trigger a new section.
    Normal/List content always goes into current section as body.
    """
    doc = Document(input_path)
    content_map = {}
    cur_key = None
    cur_paras = []   # plain text lines
    cur_elems = []   # XML elements

    def _save():
        if cur_key is None:
            return
        has_text = any(p.strip() and not _is_junk(p) for p in cur_paras)
        has_rich = any(_has_rich_content(etree.tostring(e).decode()) for e in cur_elems)
        if has_text or has_rich:
            content_map[cur_key] = {
                'text': '\n'.join(p for p in cur_paras if p.strip() and not _is_junk(p)),
                'elements': cur_elems,
            }

    for para in doc.paragraphs:
        text = para.text.strip()
        xml  = etree.tostring(para._element).decode()
        has_content = bool(text) or _has_rich_content(xml)

        if not has_content:
            if cur_key is not None:
                cur_elems.append(deepcopy(para._element))
            continue

        # Only heading-styled paragraphs can start a new section
        if _is_heading_style(para):
            key = match_heading(text)
            if key:
                _save()
                cur_key  = key
                cur_paras = []
                cur_elems = []
                continue  # heading itself not added to body

        # Everything else is body content
        if cur_key is not None:
            cur_paras.append(text)
            cur_elems.append(deepcopy(para._element))

    _save()
    return content_map, doc


def extract_txt(input_path):
    """Extract sections from plain text file using numbered/keyword headings."""
    with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()

    content_map = {}
    cur_key  = None
    cur_lines = []

    def _save():
        if cur_key and cur_lines:
            real = [l for l in cur_lines if l.strip() and not _is_junk(l)]
            if real:
                content_map[cur_key] = {'text': '\n'.join(real), 'elements': []}

    for line in lines:
        text = line.strip()
        if not text:
            continue
        # Only treat as heading if short (<=80 chars)
        if len(text) <= 80:
            key = match_heading(text.rstrip(':'))
            if key:
                _save()
                cur_key   = key
                cur_lines = []
                continue
        if cur_key is not None:
            cur_lines.append(text)

    _save()
    return content_map

# ─── Relationship Copying ────────────────────────────────────

def _copy_rels(elements, src_doc, dst_doc):
    """Copy all r:embed / r:id relationships from src to dst. Returns rId map."""
    xml_all = ' '.join(etree.tostring(e).decode() for e in elements)
    rids = set()
    for a, b in re.findall(r'r:embed="(rId\d+)"|r:id="(rId\d+)"', xml_all):
        rids.update(filter(None, [a, b]))

    rid_map = {}
    for rid in rids:
        if rid not in src_doc.part.rels:
            continue
        rel = src_doc.part.rels[rid]
        try:
            new_rid = dst_doc.part.relate_to(rel.target_part, rel.reltype)
            rid_map[rid] = new_rid
        except Exception:
            pass
    return rid_map

def _apply_rels(elements, rid_map):
    if not rid_map:
        return
    R_EMBED = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'
    R_ID    = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    for elem in elements:
        for node in elem.iter():
            for attr in [R_EMBED, R_ID]:
                old = node.get(attr)
                if old and old in rid_map:
                    node.set(attr, rid_map[old])

# ─── Template Filling ────────────────────────────────────────

def _fill_token(para, key, text):
    token = '{{' + key + '}}'
    if token not in para.text:
        return False
    for run in para.runs:
        if token in run.text:
            run.text = run.text.replace(token, text)
            return True
    # Fallback: rebuild runs
    full = para.text.replace(token, text)
    for run in para.runs:
        run.text = ''
    if para.runs:
        para.runs[0].text = full
    return True

def _find_token_para(doc, key):
    token = '{{' + key + '}}'
    for para in doc.paragraphs:
        if token in para.text:
            return para
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if token in para.text:
                        return para
    return None

def _insert_elements(output_doc, input_doc, key, content_data):
    """Replace {{key}} placeholder with actual content elements."""
    placeholder = _find_token_para(output_doc, key)
    if not placeholder:
        return

    elements = content_data.get('elements', [])
    text     = content_data.get('text', '').strip()

    if elements and input_doc:
        rid_map   = _copy_rels(elements, input_doc, output_doc)
        copies    = [deepcopy(e) for e in elements]
        _apply_rels(copies, rid_map)
        # Clear token from placeholder
        for run in placeholder.runs:
            run.text = run.text.replace('{{' + key + '}}', '')
        parent = placeholder._element.getparent()
        idx    = list(parent).index(placeholder._element)
        for j, elem in enumerate(copies):
            parent.insert(idx + 1 + j, elem)
    elif text:
        _fill_token(placeholder, key, text)
    else:
        _fill_token(placeholder, key, '')

def _fill_autofill(output_doc, key, text):
    placeholder = _find_token_para(output_doc, key)
    if placeholder:
        _fill_token(placeholder, key, text)

def _fill_simple(doc, replacements):
    for para in doc.paragraphs:
        for k, v in replacements.items():
            if '{{' + k + '}}' in para.text:
                _fill_token(para, k, v)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for k, v in replacements.items():
                        if '{{' + k + '}}' in para.text:
                            _fill_token(para, k, v)

# ─── SOP Section ─────────────────────────────────────────────

def _make_bold_para(text, color='C00000'):
    p   = OxmlElement('w:p')
    r   = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    b   = OxmlElement('w:b')
    col = OxmlElement('w:color')
    col.set(qn('w:val'), color)
    rPr.append(b); rPr.append(col)
    r.append(rPr)
    t = OxmlElement('w:t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    r.append(t); p.append(r)
    return p

def fill_sop_note(output_doc, ref_filename):
    """Put reference note in SOP section."""
    placeholder = _find_token_para(output_doc, 'sop')
    if not placeholder:
        return
    for run in placeholder.runs:
        run.text = run.text.replace('{{sop}}', '')
    parent = placeholder._element.getparent()
    idx    = list(parent).index(placeholder._element)
    parent.insert(idx + 1, _make_bold_para(
        f'[ Please open: {ref_filename} — delivered alongside this MOP in the ZIP ]',
        color='C00000'))
    parent.insert(idx + 2, _make_bold_para('─' * 55, color='808080'))

# ─── Full Input Copy ─────────────────────────────────────────

def _make_full_copy(input_doc):
    """Return a new Document that is a byte-perfect copy of input_doc."""
    new_doc = Document()
    new_doc.element.body.clear()

    body_xml = etree.tostring(input_doc.element.body).decode()
    rids = set()
    for a, b in re.findall(r'r:embed="(rId\d+)"|r:id="(rId\d+)"', body_xml):
        rids.update(filter(None, [a, b]))

    rid_map = {}
    for rid in rids:
        if rid not in input_doc.part.rels:
            continue
        rel = input_doc.part.rels[rid]
        try:
            rid_map[rid] = new_doc.part.relate_to(rel.target_part, rel.reltype)
        except Exception:
            pass

    R_EMBED = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'
    R_ID    = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

    for child in input_doc.element.body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag not in ('p', 'tbl', 'sdt', 'sectPr'):
            continue
        copy = deepcopy(child)
        for node in copy.iter():
            for attr in [R_EMBED, R_ID]:
                old = node.get(attr)
                if old and old in rid_map:
                    node.set(attr, rid_map[old])
        new_doc.element.body.append(copy)

    return new_doc

# ─── Header Date ─────────────────────────────────────────────

def _update_header_date(docx_path):
    today   = date.today().strftime('%Y-%m-%d')
    tmp     = docx_path + '.tmp'
    pattern = re.compile(r'\d{4}-\d{2}-\d{2}')
    try:
        with zipfile.ZipFile(docx_path, 'r') as zin, \
             zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            seen = set()
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in ('word/header1.xml', 'word/header2.xml'):
                    data = pattern.sub(today, data.decode()).encode()
                if item.filename not in seen:
                    zout.writestr(item, data)
                    seen.add(item.filename)
        os.replace(tmp, docx_path)
    except Exception:
        if os.path.exists(tmp):
            try: os.remove(tmp)
            except: pass

# ─── Main Entry Point ────────────────────────────────────────

def sanitize_filename(name):
    name = re.sub(r'[^\w\s-]', '', name)
    return re.sub(r'\s+', '_', name.strip())

def generate_mop(template_path, input_path, activity_name, vendor_name, output_path):
    ext = os.path.splitext(input_path)[1].lower()

    # Parse input
    input_doc = None
    if ext == '.docx':
        content_map, input_doc = extract_docx(input_path)
    elif ext in ('.txt', '.doc'):
        content_map = extract_txt(input_path)
    else:
        return {'success': False, 'message': f'Unsupported format: {ext}'}

    # Load template
    output_doc = Document(template_path)

    filled, autofilled = [], []

    # Fill each section (except sop)
    for key in HEADING_MAP:
        if key == 'sop':
            continue
        if key in content_map:
            _insert_elements(output_doc, input_doc, key, content_map[key])
            filled.append(key)
        else:
            auto = AUTO_FILL.get(key, f'{key} — {activity_name} ({vendor_name})')
            auto = auto.replace('{activity}', activity_name).replace('{vendor}', vendor_name)
            _fill_autofill(output_doc, key, auto)
            autofilled.append(key)

    # SOP — just note
    safe_act    = sanitize_filename(activity_name)
    ref_ext     = '.docx' if input_doc else ext
    ref_filename = f'{safe_act}_SOP_Reference{ref_ext}'
    fill_sop_note(output_doc, ref_filename)

    # Simple replacements (revision table)
    today = date.today().strftime('%d-%b-%Y')
    _fill_simple(output_doc, {
        'version': '1.0', 'revdate': today,
        'prepared': 'Automation SME', 'change': 'Initial Release',
    })

    # Save MOP
    output_doc.save(output_path)
    _update_header_date(output_path)

    # Create SOP Reference file
    sop_ref_path = os.path.join(os.path.dirname(output_path), ref_filename)
    if input_doc:
        _make_full_copy(input_doc).save(sop_ref_path)
    else:
        shutil.copy2(input_path, sop_ref_path)

    return {
        'success':           True,
        'message':           'MOP generated successfully.',
        'filled_sections':   filled,
        'autofilled_sections': autofilled,
        'total_sections':    len(HEADING_MAP),
        'sop_ref_path':      sop_ref_path,
    }
