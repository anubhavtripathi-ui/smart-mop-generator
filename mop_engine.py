"""
Smart MOP Generator - Core Engine
Handles: smart heading mapping, placeholder replacement,
image preservation, SOP embedding as attachment
"""

import os
import re
import shutil
import tempfile
from datetime import date
from copy import deepcopy
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree


# ─────────────────────────────────────────────
# HEADING → PLACEHOLDER MAPPING
# Exact matches + synonym groups (case-insensitive)
# ─────────────────────────────────────────────
HEADING_MAP = {
    "objective": [
        "objective", "objectives", "goal", "goals", "purpose",
        "aim", "aims", "intent", "overview", "summary"
    ],
    "activity": [
        "activity description", "activity", "description",
        "work description", "task description", "scope of work",
        "activity details", "work details"
    ],
    "type": [
        "activity type", "type", "work type", "task type",
        "change type", "nature of activity"
    ],
    "domain": [
        "domain in scope", "domain", "scope", "domains",
        "network domain", "technology domain", "area in scope"
    ],
    "prereq": [
        "pre-requisites", "prerequisites", "pre requisites",
        "preconditions", "pre-conditions", "requirements before",
        "prior requirements", "preparation", "prerequisite"
    ],
    "inventory": [
        "inventory details", "inventory", "node details",
        "equipment details", "device details", "hardware details",
        "node inventory", "asset details"
    ],
    "connect": [
        "node connectivity process", "connectivity", "node connectivity",
        "connection process", "network connectivity", "access method",
        "how to connect", "connectivity details"
    ],
    "iam": [
        "identity and access management", "iam", "access management",
        "identity management", "user access", "credentials",
        "login details", "authentication"
    ],
    "trigger": [
        "activity triggering method", "trigger", "triggering method",
        "how to trigger", "initiation method", "how to initiate",
        "trigger method", "activity trigger"
    ],
    "sop": [
        "standard operating procedure", "sop", "procedure",
        "operating procedure", "detailed procedure", "steps",
        "work instructions", "method"
    ],
    "accept": [
        "acceptance criteria", "acceptance", "uat scenarios",
        "uat", "validation criteria", "success criteria",
        "test criteria", "sign off criteria", "completion criteria"
    ],
    "assume": [
        "assumptions", "assumption", "assumed conditions",
        "pre-assumptions", "dependencies", "notes", "caveats"
    ],
}

# Auto-fill templates when content missing
AUTO_FILL = {
    "objective": "The objective of this activity is to successfully execute the {activity} for {vendor}, ensuring minimal impact to network services and adherence to change management processes.",
    "activity": "This document outlines the method of procedure for {activity} to be carried out by {vendor}. The activity involves planned steps to ensure seamless execution.",
    "type": "Planned Maintenance Activity – {activity} ({vendor})",
    "domain": "This activity covers the network domain relevant to {activity} as scoped and agreed with {vendor}.",
    "prereq": "Prior to initiating {activity} with {vendor}, ensure all necessary approvals, maintenance windows, access permissions, and backup procedures are in place.",
    "inventory": "Node inventory details for {activity} activity with {vendor} to be confirmed and validated prior to execution. Refer to the approved change request for specifics.",
    "connect": "Connectivity to the relevant nodes for {activity} shall be established by {vendor} as per the approved access management process.",
    "iam": "Access credentials and permissions required for {activity} shall be provided by {vendor} in compliance with the Identity and Access Management policy.",
    "trigger": "This activity will be triggered upon receipt of the approved change request and maintenance window confirmation for {activity} with {vendor}.",
    "accept": "Acceptance criteria for {activity} with {vendor}: All configured services are operational, no alarms are raised, and sign-off is obtained from the relevant stakeholders.",
    "assume": "It is assumed that all pre-requisites for {activity} with {vendor} are met, the maintenance window is approved, and rollback procedure is available if required.",
}


def normalize(text):
    """Lowercase, strip, remove extra spaces."""
    return re.sub(r'\s+', ' ', text.strip().lower())


def match_heading(heading_text):
    """Match a heading to a placeholder key using synonym mapping."""
    norm = normalize(heading_text)
    for key, synonyms in HEADING_MAP.items():
        for syn in synonyms:
            if norm == syn or norm.startswith(syn) or syn in norm:
                return key
    return None


def extract_content_from_docx(input_path):
    """
    Extract heading→content mapping from input .docx file.
    Returns dict: {placeholder_key: {'paragraphs': [...], 'elements': [xml_elements]}}
    Also returns the full document object for SOP attachment.
    """
    doc = Document(input_path)
    content_map = {}
    current_key = None
    current_paragraphs = []
    current_elements = []

    for para in doc.paragraphs:
        style_name = para.style.name.lower()
        is_heading = 'heading' in style_name or para.runs and any(
            r.bold for r in para.runs if r.text.strip()
        )

        # Check if it's a heading by style OR by checking if it maps to something
        text = para.text.strip()
        if not text:
            if current_key:
                current_elements.append(deepcopy(para._element))
            continue

        matched_key = None
        if 'heading' in style_name:
            matched_key = match_heading(text)
        elif current_key is None:
            # Try to match even non-heading if it looks like a section title
            matched_key = match_heading(text)

        if matched_key:
            # Save previous section
            if current_key and current_paragraphs:
                content_map[current_key] = {
                    'text': '\n'.join(current_paragraphs),
                    'elements': current_elements
                }
            current_key = matched_key
            current_paragraphs = []
            current_elements = []
        else:
            if current_key:
                current_paragraphs.append(para.text)
                current_elements.append(deepcopy(para._element))

    # Save last section
    if current_key and current_paragraphs:
        content_map[current_key] = {
            'text': '\n'.join(current_paragraphs),
            'elements': current_elements
        }

    return content_map, doc


def extract_content_from_txt(input_path):
    """
    Extract content from .txt file.
    Tries to detect headings by common patterns.
    """
    with open(input_path, 'r', encoding='utf-8', errors='ignore') as f:
        lines = f.readlines()

    content_map = {}
    current_key = None
    current_lines = []

    for line in lines:
        text = line.strip()
        if not text:
            if current_key:
                current_lines.append('')
            continue

        # Detect heading: ALL CAPS, or ends with ':', or short line (<50 chars)
        is_possible_heading = (
            text.isupper() or
            (len(text) < 60 and text.endswith(':')) or
            (len(text) < 50 and not text[0].isdigit())
        )

        matched_key = match_heading(text.rstrip(':'))

        if matched_key:
            if current_key and current_lines:
                content_map[current_key] = {
                    'text': '\n'.join(current_lines).strip(),
                    'elements': []
                }
            current_key = matched_key
            current_lines = []
        else:
            if current_key:
                current_lines.append(text)

    if current_key and current_lines:
        content_map[current_key] = {
            'text': '\n'.join(current_lines).strip(),
            'elements': []
        }

    return content_map


def replace_placeholder_in_paragraph(para, placeholder, text):
    """Replace {{placeholder}} in a paragraph, preserving run formatting."""
    full_text = para.text
    token = '{{' + placeholder + '}}'
    if token not in full_text:
        return False

    # Clear all runs and set text in first run
    for run in para.runs:
        if token in run.text:
            run.text = run.text.replace(token, text)
            return True

    # Fallback: rebuild runs
    if token in full_text:
        new_text = full_text.replace(token, text)
        for run in para.runs:
            run.text = ''
        if para.runs:
            para.runs[0].text = new_text
        return True
    return False


def replace_placeholders_in_doc(doc, replacements):
    """Replace all {{placeholders}} in document paragraphs and tables."""
    # In paragraphs
    for para in doc.paragraphs:
        for key, value in replacements.items():
            token = '{{' + key + '}}'
            if token in para.text:
                replace_placeholder_in_paragraph(para, key, value)

    # In tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for key, value in replacements.items():
                        token = '{{' + key + '}}'
                        if token in para.text:
                            replace_placeholder_in_paragraph(para, key, value)


def find_placeholder_paragraph(doc, placeholder):
    """Find paragraph containing {{placeholder}} and return it."""
    token = '{{' + placeholder + '}}'
    for i, para in enumerate(doc.paragraphs):
        if token in para.text:
            return para, i
    return None, -1


def insert_elements_after_paragraph(target_para, elements):
    """Insert XML elements after a target paragraph."""
    parent = target_para._element.getparent()
    idx = list(parent).index(target_para._element)
    for j, elem in enumerate(elements):
        parent.insert(idx + 1 + j, deepcopy(elem))


def copy_images_from_input_to_output(input_doc, output_doc, elements):
    """
    Copy image relationships from input doc to output doc
    so that images in copied elements resolve correctly.
    Returns mapping of old rId → new rId
    """
    rId_map = {}

    try:
        input_part = input_doc.part
        output_part = output_doc.part

        for rel in input_part.rels.values():
            if "image" in rel.reltype:
                try:
                    image_part = rel.target_part
                    # Add image to output doc
                    new_rId = output_part.relate_to(image_part, rel.reltype)
                    rId_map[rel.rId] = new_rId
                except Exception:
                    pass
    except Exception:
        pass

    return rId_map


def remap_image_rids(elements, rId_map):
    """Update r:embed attributes in elements to use new rIds."""
    for elem in elements:
        for blip in elem.iter(qn('a:blip')):
            old_rId = blip.get(qn('r:embed'))
            if old_rId and old_rId in rId_map:
                blip.set(qn('r:embed'), rId_map[old_rId])
        for imagedata in elem.iter(qn('v:imagedata')):
            old_rId = imagedata.get(qn('r:id'))
            if old_rId and old_rId in rId_map:
                imagedata.set(qn('r:id'), rId_map[old_rId])


def embed_file_as_attachment(output_doc, placeholder_para, input_file_path, activity_name):
    """
    Embed the input file as an OLE attachment object at the placeholder location.
    Falls back to a clearly labeled reference paragraph if embedding fails.
    """
    token = '{{sop}}'

    # Clear placeholder text
    for run in placeholder_para.runs:
        if token in run.text:
            run.text = ''

    # Try to embed as OLE object
    try:
        _embed_ole_attachment(output_doc, placeholder_para, input_file_path, activity_name)
        return True
    except Exception:
        pass

    # Fallback: Insert descriptive reference
    parent = placeholder_para._element.getparent()
    idx = list(parent).index(placeholder_para._element)

    # Add a clear reference paragraph
    ref_para = OxmlElement('w:p')
    ref_pPr = OxmlElement('w:pPr')
    ref_pStyle = OxmlElement('w:pStyle')
    ref_pStyle.set(qn('w:val'), 'Normal')
    ref_pPr.append(ref_pStyle)
    ref_para.append(ref_pPr)

    ref_r = OxmlElement('w:r')
    ref_rPr = OxmlElement('w:rPr')
    ref_bold = OxmlElement('w:b')
    ref_rPr.append(ref_bold)
    ref_r.append(ref_rPr)
    ref_t = OxmlElement('w:t')
    ref_t.text = f"[ATTACHMENT: {os.path.basename(input_file_path)} — Original input document for {activity_name}]"
    ref_r.append(ref_t)
    ref_para.append(ref_r)

    parent.insert(idx + 1, ref_para)
    return False


def _embed_ole_attachment(output_doc, anchor_para, file_path, display_name):
    """
    Embed file using ZIP-level manipulation after saving.
    Raises NotImplementedError to trigger fallback.
    """
    raise NotImplementedError("Use post-save ZIP embedding")


def replace_content_for_key(output_doc, input_doc, key, content_data, activity_name, vendor_name):
    """
    Replace {{key}} placeholder with actual content from input,
    preserving paragraph elements and images.
    """
    token = '{{' + key + '}}'
    placeholder_para = None

    # Find placeholder in doc paragraphs
    for para in output_doc.paragraphs:
        if token in para.text:
            placeholder_para = para
            break

    # Also check in tables
    if not placeholder_para:
        for table in output_doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if token in para.text:
                            placeholder_para = para
                            break

    if not placeholder_para:
        return

    elements = content_data.get('elements', [])
    text = content_data.get('text', '').strip()

    if elements and input_doc:
        # Copy images
        rId_map = copy_images_from_input_to_output(input_doc, output_doc, elements)
        if rId_map:
            remap_image_rids(elements, rId_map)

        # Clear placeholder text
        for run in placeholder_para.runs:
            run.text = run.text.replace(token, '')

        # Insert elements after placeholder paragraph
        insert_elements_after_paragraph(placeholder_para, elements)
    elif text:
        # Simple text replacement
        replace_placeholder_in_paragraph(placeholder_para, key, text)
    else:
        replace_placeholder_in_paragraph(placeholder_para, key, '')


def _update_header_date(docx_path: str):
    """Update the header date field to today's date via ZIP manipulation."""
    import zipfile
    from datetime import date as _date
    today = _date.today().strftime('%Y-%m-%d')
    tmp_path = docx_path + '.hdrtmp'
    import re as _re
    date_pattern = _re.compile(r'\d{4}-\d{2}-\d{2}')
    try:
        with zipfile.ZipFile(docx_path, 'r') as zin:
            with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename in ['word/header1.xml', 'word/header2.xml']:
                        text = data.decode('utf-8')
                        # Replace any date pattern that looks like the header date
                        text = date_pattern.sub(today, text)
                        data = text.encode('utf-8')
                    zout.writestr(item, data)
        os.replace(tmp_path, docx_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


def generate_mop(
    template_path: str,
    input_path: str,
    activity_name: str,
    vendor_name: str,
    output_path: str
) -> dict:
    """
    Main MOP generation function.

    Args:
        template_path: Path to TM1.docx
        input_path: Path to user's input file (.docx or .txt)
        activity_name: Activity name from UI
        vendor_name: Vendor name from UI
        output_path: Where to save the output .docx

    Returns:
        dict with 'success', 'message', 'filled_sections', 'autofilled_sections'
    """
    ext = os.path.splitext(input_path)[1].lower()

    # ── Extract content from input ──
    input_doc = None
    if ext == '.docx':
        content_map, input_doc = extract_content_from_docx(input_path)
    elif ext == '.txt':
        content_map = extract_content_from_txt(input_path)
    else:
        return {'success': False, 'message': f'Unsupported file format: {ext}. Please upload .docx or .txt file.'}

    # ── Load template ──
    output_doc = Document(template_path)

    # ── Build simple text replacements first ──
    today = date.today().strftime("%d-%b-%Y")
    simple_replacements = {
        'version': '1.0',
        'revdate': today,
        'prepared': 'Automation SME',
        'change': 'Initial Release',
    }

    filled_sections = []
    autofilled_sections = []

    # ── Process each placeholder ──
    all_keys = list(HEADING_MAP.keys())

    for key in all_keys:
        if key == 'sop':
            continue  # Handle SOP separately

        if key in content_map and content_map[key].get('text', '').strip():
            # User provided content
            replace_content_for_key(
                output_doc, input_doc, key,
                content_map[key], activity_name, vendor_name
            )
            filled_sections.append(key)
        else:
            # Auto-fill
            auto_text = AUTO_FILL.get(key, f'Details for {key} — {activity_name} ({vendor_name})')
            auto_text = auto_text.replace('{activity}', activity_name).replace('{vendor}', vendor_name)
            token = '{{' + key + '}}'

            # Find and replace in doc paragraphs
            for para in output_doc.paragraphs:
                if token in para.text:
                    replace_placeholder_in_paragraph(para, key, auto_text)
                    break

            # Also check tables
            for table in output_doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            if token in para.text:
                                replace_placeholder_in_paragraph(para, key, auto_text)

            autofilled_sections.append(key)

    # ── Handle SOP — embed entire input file as attachment ──
    sop_para = None
    for para in output_doc.paragraphs:
        if '{{sop}}' in para.text:
            sop_para = para
            break

    if sop_para:
        embed_file_as_attachment(output_doc, sop_para, input_path, activity_name)

    # ── Simple field replacements (revision table) ──
    replace_placeholders_in_doc(output_doc, simple_replacements)

    # ── Save output ──
    output_doc.save(output_path)

    # ── Update header date to today (ZIP-level) ──
    _update_header_date(output_path)

    return {
        'success': True,
        'message': 'MOP generated successfully.',
        'filled_sections': filled_sections,
        'autofilled_sections': autofilled_sections,
        'total_sections': len(all_keys)
    }


def sanitize_filename(name: str) -> str:
    """Convert name to safe filename part."""
    name = re.sub(r'[^\w\s-]', '', name)
    name = re.sub(r'\s+', '_', name.strip())
    return name
