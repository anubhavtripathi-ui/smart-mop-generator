"""
Smart MOP Generator - Core Engine
"""
import os, re, zipfile, struct, io
from datetime import date
from copy import deepcopy
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

# ─────────────────────────────────────────────────────────────────────────────
# HEADING MAP  (expanded with telecom / network domain synonyms)
# ─────────────────────────────────────────────────────────────────────────────
HEADING_MAP = {
    "objective": [
        "objective", "objectives", "goal", "goals", "purpose", "aim",
        "overview", "summary", "scope summary", "executive summary",
        "background", "introduction", "project objective",
    ],
    "activity": [
        "activity description", "activity", "description", "work description",
        "task description", "scope of work", "work scope", "job description",
        "activity details", "task details", "project description",
        "change description", "activity overview",
    ],
    "type": [
        "activity type", "type of activity", "work type", "task type",
        "change type", "nature of activity", "change category",
        "activity classification", "type of change", "type of work",
    ],
    "domain": [
        "domain in scope", "domain", "scope", "domains", "network domain",
        "technology domain", "technology scope", "network scope",
        "impacted domain", "domain impacted", "affected domain",
        "systems in scope", "systems impacted",
    ],
    "prereq": [
        "pre-requisites", "prerequisites", "pre requisites", "preconditions",
        "prerequisite", "preparation", "pre-conditions", "readiness criteria",
        "pre-activity checklist", "checklist", "pre-checks", "pre checks",
        "pre activity", "activities before", "steps before",
    ],
    "inventory": [
        "inventory details", "inventory", "node details", "equipment details",
        "node inventory", "asset details", "device inventory",
        "network inventory", "site details", "site information",
        "hardware details", "equipment list", "node list",
        "impacted nodes", "affected nodes", "node information",
    ],
    "connect": [
        "node connectivity process", "connectivity", "node connectivity",
        "connection process", "access method", "how to connect",
        "connection procedure", "connection steps", "network access",
        "remote access", "terminal access", "ssh access", "telnet access",
        "vpn access", "jump server", "access procedure",
    ],
    "iam": [
        "identity and access management", "iam", "access management",
        "credentials", "login details", "authentication",
        "login process", "login procedure", "login steps",
        "access credentials", "username and password", "user access",
        "account details", "system access", "portal access",
        "access details", "oss access", "nms access", "tool access",
        "how to login", "login method", "sign in procedure",
    ],
    "trigger": [
        "activity triggering method", "trigger", "triggering method",
        "activity trigger", "how to trigger", "initiation process",
        "how to initiate", "kick off", "kickoff", "start procedure",
        "activity initiation", "triggering process", "trigger method",
        "how to start", "launch procedure", "initiation method",
        "activity launch", "escalation trigger",
    ],
    "sop": [
        "standard operating procedure", "sop", "procedure",
        "operating procedure", "detailed procedure", "step by step",
        "execution steps", "implementation steps", "detailed steps",
        "activity steps", "work procedure", "technical procedure",
        "task procedure", "process steps", "execution procedure",
    ],
    "accept": [
        "acceptance criteria", "acceptance", "uat scenarios", "uat",
        "validation criteria", "success criteria", "sign off criteria",
        "sign-off criteria", "test cases", "validation steps",
        "post activity checks", "post-activity checks", "verification steps",
        "testing criteria", "health check", "sanity check",
        "post implementation review", "post implementation checks",
        "kpi", "kpis", "performance criteria",
    ],
    "assume": [
        "assumptions", "assumption", "assumed conditions", "dependencies",
        "notes", "caveats", "constraints", "limitations",
        "risks", "risk assessment", "remarks", "additional notes",
        "general notes", "important notes", "exclusions",
        "out of scope", "exceptions", "conditions",
    ],
}

AUTO_FILL = {
    "objective":  "The objective of this activity is to successfully execute the {activity} for {vendor}, ensuring minimal impact to network services and adherence to change management processes.",
    "activity":   "This document outlines the method of procedure for {activity} to be carried out by {vendor}. The activity involves planned steps to ensure seamless execution.",
    "type":       "Planned Maintenance Activity - {activity} ({vendor})",
    "domain":     "This activity covers the network domain relevant to {activity} as scoped and agreed with {vendor}.",
    "prereq":     "Prior to initiating {activity} with {vendor}, ensure all necessary approvals, maintenance windows, and access permissions are in place.",
    "inventory":  "Node inventory details for {activity} with {vendor} to be confirmed prior to execution. Refer to the approved change request for specifics.",
    "connect":    "Connectivity to the relevant nodes for {activity} shall be established by {vendor} as per the approved access management process.",
    "iam":        "Access credentials for {activity} shall be provided by {vendor} in compliance with the Identity and Access Management policy.",
    "trigger":    "This activity will be triggered upon receipt of the approved change request and maintenance window confirmation for {activity} with {vendor}.",
    "accept":     "Acceptance criteria for {activity} with {vendor}: All configured services are operational, no alarms raised, sign-off obtained from stakeholders.",
    "assume":     "It is assumed all pre-requisites for {activity} with {vendor} are met, maintenance window is approved, and rollback procedure is available.",
}

JUNK_PATTERNS = [
    'this section is to provide', 'mention the', 'provide activity detailed',
    'provide fallback', 'please include a detailed description',
]

# ─────────────────────────────────────────────────────────────────────────────
# HEADING DETECTION
# ─────────────────────────────────────────────────────────────────────────────
def normalize(text):
    return re.sub(r'\s+', ' ', text.strip().lower())


def strip_numbering(text):
    """Remove leading numbering: '1.', '1)', '1.2.', '1.1 ', 'Step 3:', 'A.' etc."""
    t = text.strip()
    # Step/Phase/Section prefix
    t = re.sub(r'^(step|phase|section|part)\s*\d+\s*[:\-\.]?\s*', '', t, flags=re.IGNORECASE)
    # Numeric: 1. / 1) / 1.2 / 1.2. / 1.2.3 / 10. etc.
    t = re.sub(r'^[\d]+(?:\.[\d]+)*\.?\)?\s+', '', t)
    # Alpha: A. / a) / A)
    t = re.sub(r'^[A-Za-z][\.\)]\s*', '', t)
    return t.strip() if t.strip() else text


def match_heading(heading_text, strict=False):
    """
    Return HEADING_MAP key if heading_text matches any synonym, else None.

    strict=True  → used for non-styled paragraphs (no heading style in Word).
                   Only exact / prefix / suffix matches are accepted.
                   Substring matching is disabled to avoid false positives in
                   content lines like "To perform NDS health check on nodes."

    strict=False → used for proper Word Heading styles; allows looser substring
                   matching for short synonyms still >= 10 chars.
    """
    # Sentence-starter words that signal body content, not headings
    CONTENT_STARTERS = (
        'to ', 'the ', 'this ', 'these ', 'in ', 'all ', 'as ', 'it ', 'by ',
        'for ', 'please ', 'note ', 'ensure ', 'verify ', 'connect ', 'after ',
        'before ', 'once ', 'during ', 'following ', 'based ', 'due ',
    )

    candidates = [heading_text]
    stripped = strip_numbering(heading_text)
    if stripped != heading_text:
        candidates.append(stripped)

    for candidate in candidates:
        norm = normalize(candidate)

        # Skip obvious content sentences in strict mode
        if strict and any(norm.startswith(s) for s in CONTENT_STARTERS):
            continue

        # ── 1. Exact match ──
        for key, synonyms in HEADING_MAP.items():
            if norm in synonyms:
                return key

        # ── 2. Prefix / suffix match ──
        best_key, best_len = None, 0
        for key, synonyms in HEADING_MAP.items():
            for syn in synonyms:
                if norm.startswith(syn) and len(syn) > best_len:
                    best_key, best_len = key, len(syn)
                if syn.startswith(norm) and len(norm) > 2 and len(norm) > best_len:
                    best_key, best_len = key, len(norm)

        # ── 3. Substring match — only for Word Heading styles (not strict) ──
        if not best_key and not strict:
            for key, synonyms in HEADING_MAP.items():
                for syn in synonyms:
                    if len(syn) >= 10 and syn in norm and len(syn) > best_len:
                        best_key, best_len = key, len(syn)

        if best_key:
            return best_key
    return None


def is_junk(text):
    t = text.strip().lower()
    if t.startswith('[') and t.endswith(']'):
        return True
    return any(j in t for j in JUNK_PATTERNS)


def has_real_content(elements, paragraphs):
    if any(p.strip() and not is_junk(p) for p in paragraphs):
        return True
    for elem in elements:
        xml = etree.tostring(elem).decode()
        if 'blip' in xml or 'OLEObject' in xml or 'oleObject' in xml or '<w:drawing>' in xml:
            return True
    return False


# ─────────────────────────────────────────────────────────────────────────────
# CONTENT EXTRACTION
# ─────────────────────────────────────────────────────────────────────────────
def extract_content_from_docx(input_path):
    """
    Extract sections from input docx.
    Rules:
    - Heading styles -> matched immediately
    - Non-styled short lines -> also checked (bold paragraph headings, numbered headings)
    - Sub-headings (1.1, 1.2) -> treated as CONTENT of parent heading
    - Same heading key seen twice -> first occurrence wins
    - Floating images/tables -> attached to nearest preceding section
    """
    doc = Document(input_path)
    content_map   = {}
    seen_keys     = set()
    current_key   = None
    current_paras = []
    current_elems = []

    def flush():
        nonlocal current_key, current_paras, current_elems
        if current_key and has_real_content(current_elems, current_paras):
            if current_key not in content_map:
                content_map[current_key] = {
                    'text':     '\n'.join(p for p in current_paras if not is_junk(p)),
                    'elements': current_elems
                }
        current_key   = None
        current_paras = []
        current_elems = []

    for para in doc.paragraphs:
        style_name = para.style.name.lower()
        text       = para.text.strip()
        xml_str    = etree.tostring(para._element).decode()
        has_rich   = 'blip' in xml_str or 'OLEObject' in xml_str or '<w:drawing>' in xml_str

        if not text and not has_rich:
            if current_key:
                current_elems.append(deepcopy(para._element))
            continue

        is_heading_style = 'heading' in style_name
        matched_key = None

        if is_heading_style and text:
            matched_key = match_heading(text, strict=False)
        elif text and len(text) < 120:
            candidate = match_heading(text, strict=True)
            if candidate and candidate != current_key:
                matched_key = candidate

        if matched_key and matched_key not in seen_keys:
            # New top-level section
            flush()
            current_key = matched_key
            seen_keys.add(matched_key)

        elif matched_key and matched_key in seen_keys:
            # Sub-heading or duplicate -> fold as content
            if current_key:
                current_paras.append(text)
                current_elems.append(deepcopy(para._element))

        else:
            # Plain content
            if current_key:
                if text:
                    current_paras.append(text)
                current_elems.append(deepcopy(para._element))
            elif has_rich:
                # Floating rich content with no current section yet -> skip safely
                pass

    flush()
    return content_map, doc


def extract_content_from_txt(input_path):
    with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()
    content_map   = {}
    current_key   = None
    current_lines = []
    seen_keys     = set()

    for line in lines:
        text = line.strip()
        if not text:
            continue
        matched = match_heading(text.rstrip(':'))
        if matched and matched not in seen_keys:
            if current_key and current_lines:
                real = [l for l in current_lines if l.strip() and not is_junk(l)]
                if real:
                    content_map[current_key] = {'text': '\n'.join(real), 'elements': []}
            current_key   = matched
            current_lines = []
            seen_keys.add(matched)
        elif matched and matched in seen_keys:
            if current_key:
                current_lines.append(text)
        else:
            if current_key:
                current_lines.append(text)

    if current_key and current_lines:
        real = [l for l in current_lines if l.strip() and not is_junk(l)]
        if real:
            content_map[current_key] = {'text': '\n'.join(real), 'elements': []}
    return content_map


# ─────────────────────────────────────────────────────────────────────────────
# RELATIONSHIP / CONTENT HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def copy_all_relationships(input_doc, output_doc, elements):
    rId_map = {}
    try:
        input_part  = input_doc.part
        output_part = output_doc.part
        element_xml = ' '.join(etree.tostring(e).decode() for e in elements)
        flat_rids   = set()
        for pair in re.findall(r'r:embed="(rId\d+)"|r:id="(rId\d+)"', element_xml):
            for rid in pair:
                if rid: flat_rids.add(rid)
        for rId in flat_rids:
            if rId not in input_part.rels:
                continue
            rel = input_part.rels[rId]
            try:
                new_rId = output_part.relate_to(rel.target_part, rel.reltype)
                rId_map[rId] = new_rId
            except Exception:
                pass
    except Exception:
        pass
    return rId_map


def remap_rids_in_elements(elements, rId_map):
    if not rId_map:
        return
    r_embed = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed'
    r_id    = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    for elem in elements:
        for node in elem.iter():
            for attr in [r_embed, r_id]:
                old = node.get(attr)
                if old and old in rId_map:
                    node.set(attr, rId_map[old])


def replace_placeholder_in_paragraph(para, placeholder, text):
    token = '{{' + placeholder + '}}'
    if token not in para.text:
        return False
    for run in para.runs:
        if token in run.text:
            run.text = run.text.replace(token, text)
            return True
    full_text = para.text.replace(token, text)
    for run in para.runs:
        run.text = ''
    if para.runs:
        para.runs[0].text = full_text
    return True


def replace_placeholders_in_doc(doc, replacements):
    for para in doc.paragraphs:
        for key, value in replacements.items():
            if '{{' + key + '}}' in para.text:
                replace_placeholder_in_paragraph(para, key, value)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for key, value in replacements.items():
                        if '{{' + key + '}}' in para.text:
                            replace_placeholder_in_paragraph(para, key, value)


def replace_content_for_key(output_doc, input_doc, key, content_data, activity_name, vendor_name):
    token = '{{' + key + '}}'
    placeholder_para = None
    for para in output_doc.paragraphs:
        if token in para.text:
            placeholder_para = para
            break
    if not placeholder_para:
        for table in output_doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if token in para.text:
                            placeholder_para = para
    if not placeholder_para:
        return

    elements = content_data.get('elements', [])
    text     = content_data.get('text', '').strip()

    if elements and input_doc:
        rId_map = copy_all_relationships(input_doc, output_doc, elements)
        if rId_map:
            remap_rids_in_elements(elements, rId_map)
        for run in placeholder_para.runs:
            run.text = run.text.replace(token, '')
        parent = placeholder_para._element.getparent()
        idx    = list(parent).index(placeholder_para._element)
        for j, elem in enumerate(elements):
            parent.insert(idx + 1 + j, deepcopy(elem))
    elif text:
        replace_placeholder_in_paragraph(placeholder_para, key, text)
    else:
        replace_placeholder_in_paragraph(placeholder_para, key, '')


# ─────────────────────────────────────────────────────────────────────────────
# SOP SECTION -> OLE MARKER
# ─────────────────────────────────────────────────────────────────────────────
def insert_sop_ole_marker(output_doc):
    """Replace {{sop}} with an XML comment marker for post-process OLE injection."""
    token = '{{sop}}'
    for para in output_doc.paragraphs:
        if token in para.text:
            for run in para.runs:
                run.text = run.text.replace(token, '')
            comment = etree.Comment('OLE_SOP_HERE')
            para._element.append(comment)
            return
    for table in output_doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if token in para.text:
                        for run in para.runs:
                            run.text = run.text.replace(token, '')
                        comment = etree.Comment('OLE_SOP_HERE')
                        para._element.append(comment)
                        return


# ─────────────────────────────────────────────────────────────────────────────
# WMF ICON GENERATOR
# ─────────────────────────────────────────────────────────────────────────────
def _create_word_icon_wmf():
    """Build a minimal WMF binary (Word-blue rectangle) used as OLE icon."""
    def rec(func, params):
        size_words = 3 + len(params)
        data = struct.pack('<IH', size_words, func)
        for p in params:
            data += struct.pack('<h', int(p))
        return data

    recs = b''
    recs += rec(0x0103, [8])
    recs += rec(0x020B, [0, 0])
    recs += rec(0x020C, [120, 96])
    recs += rec(0x02FC, [0, 0x7244, 0x00C4, 0])
    recs += rec(0x012D, [0])
    recs += rec(0x02FA, [0, 0, 0, 0, 5])
    recs += rec(0x012D, [1])
    recs += rec(0x041B, [120, 96, 0, 0])
    recs += struct.pack('<IH', 3, 0)

    total_words = (18 + len(recs)) // 2
    pos = max_rec = 0
    while pos < len(recs):
        if pos + 4 > len(recs): break
        rs = struct.unpack_from('<I', recs, pos)[0]
        if rs == 0: break
        max_rec = max(max_rec, rs)
        pos += rs * 2

    header = struct.pack('<HHHHIHH', 2, 9, 0x0300, total_words, 2, max_rec, 0)
    return header + recs


# ─────────────────────────────────────────────────────────────────────────────
# OLE INJECTION  (post-process after docx.save())
# ─────────────────────────────────────────────────────────────────────────────
def inject_ole_attachment(docx_path, input_file_path, display_name):
    """Embed input_file_path as OLE Word icon in place of <!--OLE_SOP_HERE-->."""
    wmf_bytes   = _create_word_icon_wmf()
    input_bytes = open(input_file_path, 'rb').read()
    safe_name   = re.sub(r'[^\w\s\-.]', '', display_name)[:50]

    buf = io.BytesIO(open(docx_path, 'rb').read())
    out_buf = io.BytesIO()

    with zipfile.ZipFile(buf, 'r') as zin:
        doc_xml  = zin.read('word/document.xml').decode('utf-8')
        rels_xml = zin.read('word/_rels/document.xml.rels').decode('utf-8')
        ct_xml   = zin.read('[Content_Types].xml').decode('utf-8')

        existing  = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels_xml)]
        next_rid  = max(existing, default=0) + 1
        rId_icon  = f"rId{next_rid}"
        rId_embed = f"rId{next_rid + 1}"

        media_nums = [int(m) for m in re.findall(r'media/\w+?(\d+)\.', rels_xml + doc_xml)]
        emb_nums   = [int(m) for m in re.findall(r'embeddings/\w+?(\d+)', rels_xml + doc_xml)]
        icon_num   = max(media_nums, default=0) + 1
        embed_num  = max(emb_nums,   default=0) + 1

        icon_zip_path  = f'word/media/ole_icon{icon_num}.wmf'
        embed_zip_path = f'word/embeddings/Microsoft_Word_Document{embed_num}.docx'

        w_dxa, h_dxa = 1440, 1872
        w_pt  = w_dxa // 20
        h_pt  = h_dxa // 20

        ole_xml = (
            '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:v="urn:schemas-microsoft-com:vml" '
            'xmlns:o="urn:schemas-microsoft-com:office:office">'
            '<w:r>'
            f'<w:object w:dxaOrig="{w_dxa}" w:dyaOrig="{h_dxa}">'
            f'<v:shape id="_x0000_i1025" type="#_x0000_t75" '
            f'style="width:{w_pt}pt;height:{h_pt}pt" o:ole="">'
            f'<v:imagedata r:id="{rId_icon}" o:title="{safe_name}"/>'
            '</v:shape>'
            f'<o:OLEObject Type="Embed" ProgID="Word.Document.12" '
            f'ShapeID="_x0000_i1025" DrawAspect="Icon" '
            f'ObjectID="_1000000001" r:id="{rId_embed}">'
            r'<o:FieldCodes>\s</o:FieldCodes>'
            '</o:OLEObject>'
            '</w:object>'
            '</w:r>'
            '</w:p>'
        )

        label_xml = (
            '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:pPr><w:jc w:val="left"/></w:pPr>'
            '<w:r>'
            f'<w:t xml:space="preserve">{safe_name}</w:t>'
            '</w:r>'
            '</w:p>'
        )

        new_doc_xml = doc_xml.replace('<!--OLE_SOP_HERE-->', ole_xml + label_xml)

        new_rels = rels_xml.replace('</Relationships>',
            f'<Relationship Id="{rId_icon}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
            f'Target="media/ole_icon{icon_num}.wmf"/>\n'
            f'<Relationship Id="{rId_embed}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" '
            f'Target="embeddings/Microsoft_Word_Document{embed_num}.docx"/>\n'
            '</Relationships>'
        )

        new_ct = ct_xml
        if 'image/x-wmf' not in new_ct:
            new_ct = new_ct.replace('</Types>',
                '<Default Extension="wmf" ContentType="image/x-wmf"/>\n</Types>')
        new_ct = new_ct.replace('</Types>',
            f'<Override PartName="/word/embeddings/Microsoft_Word_Document{embed_num}.docx" '
            f'ContentType="application/vnd.openxmlformats-officedocument.'
            f'wordprocessingml.document.main+xml"/>\n</Types>'
        )

        with zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            seen = set()
            for item in zin.infolist():
                if item.filename in seen:
                    continue
                seen.add(item.filename)
                if item.filename == 'word/document.xml':
                    zout.writestr(item, new_doc_xml.encode('utf-8'))
                elif item.filename == 'word/_rels/document.xml.rels':
                    zout.writestr(item, new_rels.encode('utf-8'))
                elif item.filename == '[Content_Types].xml':
                    zout.writestr(item, new_ct.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item.filename))
            zout.writestr(icon_zip_path,  wmf_bytes)
            zout.writestr(embed_zip_path, input_bytes)

    with open(docx_path, 'wb') as f:
        f.write(out_buf.getvalue())


# ─────────────────────────────────────────────────────────────────────────────
# HEADER DATE UPDATE
# ─────────────────────────────────────────────────────────────────────────────
def _update_header_date(docx_path):
    today    = date.today().strftime('%Y-%m-%d')
    tmp_path = docx_path + '.hdrtmp'
    date_pat = re.compile(r'\d{4}-\d{2}-\d{2}')
    try:
        with zipfile.ZipFile(docx_path, 'r') as zin:
            with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                seen = set()
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename in ('word/header1.xml', 'word/header2.xml'):
                        data = date_pat.sub(today, data.decode('utf-8')).encode('utf-8')
                    if item.filename not in seen:
                        zout.writestr(item, data)
                        seen.add(item.filename)
        os.replace(tmp_path, docx_path)
    except Exception:
        if os.path.exists(tmp_path):
            try: os.remove(tmp_path)
            except: pass


# ─────────────────────────────────────────────────────────────────────────────
# MAIN GENERATOR
# ─────────────────────────────────────────────────────────────────────────────
def generate_mop(template_path, input_path, activity_name, vendor_name, output_path):
    ext = os.path.splitext(input_path)[1].lower()
    input_doc = None

    if ext == '.docx':
        content_map, input_doc = extract_content_from_docx(input_path)
    elif ext == '.txt':
        content_map = extract_content_from_txt(input_path)
    else:
        return {'success': False,
                'message': f'Unsupported format: {ext}. Upload .docx or .txt'}

    output_doc = Document(template_path)
    today      = date.today().strftime("%d-%b-%Y")

    simple_replacements = {
        'version': '1.0', 'revdate': today,
        'prepared': 'Automation SME', 'change': 'Initial Release',
    }

    filled_sections     = []
    autofilled_sections = []
    all_keys = list(HEADING_MAP.keys())

    for key in all_keys:
        if key == 'sop':
            continue
        if key in content_map:
            replace_content_for_key(
                output_doc, input_doc, key, content_map[key],
                activity_name, vendor_name
            )
            filled_sections.append(key)
        else:
            auto_text = AUTO_FILL.get(key, f'{key} details for {activity_name} ({vendor_name})')
            auto_text = auto_text.replace('{activity}', activity_name).replace('{vendor}', vendor_name)
            token = '{{' + key + '}}'
            for para in output_doc.paragraphs:
                if token in para.text:
                    replace_placeholder_in_paragraph(para, key, auto_text)
                    break
            for table in output_doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            if token in para.text:
                                replace_placeholder_in_paragraph(para, key, auto_text)
            autofilled_sections.append(key)

    insert_sop_ole_marker(output_doc)
    replace_placeholders_in_doc(output_doc, simple_replacements)
    output_doc.save(output_path)

    # Post-process: inject OLE attachment
    input_filename = os.path.splitext(os.path.basename(input_path))[0]
    try:
        inject_ole_attachment(output_path, input_path, input_filename)
    except Exception:
        pass  # Non-fatal: MOP still usable without OLE

    _update_header_date(output_path)

    return {
        'success':             True,
        'message':             'MOP generated successfully.',
        'filled_sections':     filled_sections,
        'autofilled_sections': autofilled_sections,
        'total_sections':      len(all_keys)
    }


def sanitize_filename(name):
    name = re.sub(r'[^\w\s-]', '', name)
    name = re.sub(r'\s+', '_', name.strip())
    return name
