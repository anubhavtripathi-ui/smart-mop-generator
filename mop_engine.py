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
        # from HEADING_ALIASES
        "scope", "objective",
    ],
    "activity": [
        "activity description", "activity", "work description",
        "task description", "scope of work", "activity steps",
        "activity window", "activity window/ network", "change description",
        "work summary",
        # from HEADING_ALIASES
        "task overview", "process description", "activity overview", "execution details",
    ],
    "type": [
        "activity type", "type of activity", "work type",
        "task type", "change type", "nature of activity",
        # from HEADING_ALIASES
        "task category", "process type", "operation type", "activity classification",
    ],
    "domain": [
        "domain in scope", "domain", "scope", "domains",
        "network domain", "technology domain",
        "scope/description", "scope description",
        # from HEADING_ALIASES
        "applicable domain", "functional scope",
    ],
    "prereq": [
        "pre-requisites", "prerequisites", "pre requisites",
        "preconditions", "prerequisite", "preparation",
        "pre-requisites check list", "prerequisites check list",
        "pre-checks", "pre checks", "prechecks",
        # from HEADING_ALIASES
        "pre-conditions", "initial requirements", "mandatory conditions",
    ],
    "inventory": [
        "inventory details", "inventory", "node details",
        "equipment details", "node inventory", "asset details",
        "input sheet", "node list",
        # from HEADING_ALIASES
        "infrastructure details", "system inventory",
    ],
    "connect": [
        "node connectivity process", "connectivity", "node connectivity",
        "connection process", "access method", "scripts used",
        # from HEADING_ALIASES
        "connectivity workflow", "integration process", "network configuration steps",
        "connection procedure",
    ],
    "iam": [
        "identity and access management", "iam", "access management",
        "credentials", "login details", "authentication",
        # from HEADING_ALIASES
        "access control details", "authentication process", "authorization matrix",
        "identity & access management",
    ],
    "trigger": [
        "activity triggering method", "trigger", "triggering method",
        "activity trigger", "how to trigger",
        # from HEADING_ALIASES
        "trigger mechanism", "initiation method", "activation process",
        "execution trigger", "event trigger",
    ],
    "sop": [
        "standard operating procedure", "sop",
        "operating procedure", "detailed procedure",
        # from HEADING_ALIASES
        "operational guidelines", "process manual", "execution procedure",
        "work instructions", "step-by-step guide",
        "standard operating procedure (attach the detailed sop)",
    ],
    "accept": [
        "acceptance criteria", "acceptance test criteria",
        "acceptance test", "acceptance",
        "uat scenarios", "uat criteria", "uat",
        "validation criteria", "success criteria",
        "sign off criteria", "sign-off criteria",
        "test acceptance", "post checks", "post-checks",
        "post requisites", "post-requisites",
        # from HEADING_ALIASES
        "test scenarios", "approval conditions", "success parameters",
        "uat checklist", "acceptance criteria (uat scenarios)",
    ],
    "assume": [
        "assumptions", "assumption", "assumed conditions",
        "dependencies", "notes", "caveats",
        "risk & impact", "risk and impact", "rollback", "fallback",
        # from HEADING_ALIASES
        "presumptions", "considerations", "operating assumptions",
    ],
}

AUTO_FILL = {
    "objective": "The objective of this activity is to perform {activity} as part of the operational process for {vendor}. This procedure ensures that all necessary steps are followed systematically and accurately. The goal is to achieve the desired outcome with minimal risk and downtime.",
    "activity":  "This activity involves the execution of {activity} by {vendor} team. The process covers all relevant steps from initiation to completion. All actions will be performed in accordance with standard operational guidelines.",
    "type":      "This is a planned operational activity of type: {activity}. It is categorized under standard change management procedures for {vendor}. The activity classification is based on its impact and execution scope.",
    "domain":    "The domain in scope for this activity includes the systems and components managed by {vendor}. All functional areas relevant to {activity} are included within this scope. Out-of-scope items will be documented separately if applicable.",
    "prereq":    "Prior to executing {activity}, all prerequisite conditions must be verified. Access credentials, system availability, and approvals from {vendor} must be confirmed. All stakeholders should be notified and change window should be secured.",
    "inventory": "The inventory for {activity} includes relevant nodes, systems, and components managed by {vendor}. A detailed inventory list should be prepared prior to execution. Node names, types, counts, and vendor details must be documented and verified.",
    "connect":   "The node connectivity process for {activity} involves verifying all network paths and connections. {vendor} team will ensure proper integration and connectivity between all nodes. Connectivity tests will be performed before and after the activity.",
    "iam":       "Access management for {activity} will follow the standard IAM process defined by {vendor}. All user accounts and roles must be verified prior to execution. Access logs will be maintained for audit and compliance purposes.",
    "trigger":   "The activity {activity} will be triggered based on the approved change request from {vendor}. Initiation will follow the standard trigger mechanism defined in the change management process. The execution will commence only after receiving explicit approval.",
    "accept":    "The acceptance criteria for {activity} will be validated by {vendor} team post-execution. All UAT scenarios must pass before the activity is marked as complete. Any deviations from expected results must be documented and escalated immediately.",
    "assume":    "It is assumed that all required systems are available and accessible during {activity}. {vendor} team will have necessary access and permissions throughout the execution window. Any changes in assumptions will be communicated to all stakeholders prior to execution.",
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
    """Rule 1: Word style = Heading 1/2/3 etc."""
    return 'heading' in para.style.name.lower()

def _is_bold_para(para):
    """Rule 2: ALL runs in para are bold (and para has text)."""
    runs = [r for r in para.runs if r.text.strip()]
    return bool(runs) and all(r.bold for r in runs)

def _ends_with_heading_marker(text):
    """Rule 3: Line ends with ':-', ':', or '-'."""
    t = text.rstrip()
    return t.endswith(':-') or t.endswith(':') or t.endswith('-')

def _is_underlined_para(para):
    """Rule 4: ALL runs in para are underlined (and para has text)."""
    runs = [r for r in para.runs if r.text.strip()]
    return bool(runs) and all(r.underline for r in runs)

def _get_para_font_size(para):
    """Return font size in half-points for first run that has one, else None."""
    for run in para.runs:
        if run.font.size:
            return run.font.size
    # fallback: check paragraph-level rPr sz element
    WNS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    rPr = para._element.find(f'.//{WNS}rPr')
    if rPr is not None:
        sz = rPr.find(f'{WNS}sz')
        if sz is not None:
            try:
                return int(sz.get(f'{WNS}val', 0))
            except ValueError:
                pass
    return None

def _is_heading_para(para, prev_size, next_size):
    """
    Returns True if paragraph qualifies as a heading by ANY of the 5 rules:
      1. Word Heading style
      2. All runs bold
      3. Ends with ':', '-', or ':-'
      4. All runs underlined
      5. Font size larger than BOTH surrounding paragraphs
    """
    text = para.text.strip()
    if not text:
        return False
    if _is_heading_style(para):              # Rule 1
        return True
    if _is_bold_para(para):                  # Rule 2
        return True
    if _ends_with_heading_marker(text):      # Rule 3
        return True
    if _is_underlined_para(para):            # Rule 4
        return True
    # Rule 5: font size bigger than both neighbours (if at least one available)
    my_size = _get_para_font_size(para)
    if my_size:
        neighbours = [s for s in [prev_size, next_size] if s is not None]
        if neighbours and all(my_size > s for s in neighbours):
            return True
    return False

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

    Heading detection uses 5 rules (ANY one is enough):
      1. Word Heading style (Heading 1/2/3 …)
      2. All runs are bold
      3. Line ends with ':', '-', or ':-'
      4. All runs are underlined
      5. Font size larger than both surrounding paragraphs

    Matching is always case-insensitive.
    """
    doc = Document(input_path)
    paragraphs = doc.paragraphs          # stable list, indexed for lookahead
    n = len(paragraphs)

    # Pre-compute font sizes for Rule 5 lookahead
    sizes = [_get_para_font_size(p) for p in paragraphs]

    content_map = {}
    cur_key   = None
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

    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        xml  = etree.tostring(para._element).decode()
        has_content = bool(text) or _has_rich_content(xml)

        if not has_content:
            if cur_key is not None:
                cur_elems.append(deepcopy(para._element))
            continue

        # Font-size neighbours for Rule 5
        prev_size = sizes[i - 1] if i > 0 else None
        next_size = sizes[i + 1] if i < n - 1 else None

        # Check all 5 heading rules
        if _is_heading_para(para, prev_size, next_size):
            # Strip trailing heading markers (':-', ':', '-') before text matching
            clean_text = text
            for marker in (':-', ':', '-'):
                if clean_text.rstrip().endswith(marker):
                    clean_text = clean_text.rstrip()[:-len(marker)].rstrip()
                    break
            key = match_heading(clean_text)
            if key:
                _save()
                cur_key   = key
                cur_paras = []
                cur_elems = []
                continue   # heading itself not added to body

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
