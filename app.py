"""
Smart MOP Generator
Professional Streamlit UI
"""

import os
import shutil
import tempfile
import streamlit as st
from mop_engine import generate_mop, sanitize_filename

# ─── Page Config ───
st.set_page_config(
    page_title="Smart MOP Generator",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ─── Custom CSS ───
st.markdown("""
<style>
    /* Main background */
    .stApp {
        background: linear-gradient(135deg, #0f1724 0%, #1a2744 50%, #0f1724 100%);
        min-height: 100vh;
    }

    /* Header banner */
    .mop-header {
        background: linear-gradient(90deg, #1e3a5f 0%, #2563a8 50%, #1e3a5f 100%);
        border: 1px solid #3a7bd5;
        border-radius: 12px;
        padding: 28px 32px;
        margin-bottom: 28px;
        text-align: center;
        box-shadow: 0 4px 24px rgba(37,99,168,0.3);
    }
    .mop-header h1 {
        color: #ffffff;
        font-size: 2rem;
        font-weight: 700;
        margin: 0 0 6px 0;
        letter-spacing: 0.5px;
    }
    .mop-header p {
        color: #a8c4e0;
        font-size: 0.95rem;
        margin: 0;
    }

    /* Card containers */
    .mop-card {
        background: rgba(255,255,255,0.04);
        border: 1px solid rgba(255,255,255,0.10);
        border-radius: 10px;
        padding: 24px;
        margin-bottom: 20px;
    }

    /* Section labels */
    .section-label {
        color: #a8c4e0;
        font-size: 0.78rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-bottom: 8px;
    }

    /* Privacy badge */
    .privacy-badge {
        background: rgba(34,197,94,0.10);
        border: 1px solid rgba(34,197,94,0.30);
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 20px;
        display: flex;
        align-items: center;
        gap: 10px;
        color: #86efac;
        font-size: 0.85rem;
    }

    /* Success box */
    .success-box {
        background: rgba(34,197,94,0.08);
        border: 1px solid rgba(34,197,94,0.35);
        border-radius: 10px;
        padding: 20px 24px;
        margin-top: 16px;
    }
    .success-box h3 {
        color: #4ade80;
        margin: 0 0 10px 0;
        font-size: 1.1rem;
    }
    .success-box p {
        color: #86efac;
        margin: 4px 0;
        font-size: 0.88rem;
    }

    /* Processing box */
    .process-box {
        background: rgba(59,130,246,0.08);
        border: 1px solid rgba(59,130,246,0.30);
        border-radius: 10px;
        padding: 16px 20px;
        color: #93c5fd;
        font-size: 0.88rem;
        margin-top: 12px;
    }

    /* Error box */
    .error-box {
        background: rgba(239,68,68,0.08);
        border: 1px solid rgba(239,68,68,0.30);
        border-radius: 10px;
        padding: 16px 20px;
        color: #fca5a5;
        font-size: 0.88rem;
        margin-top: 12px;
    }

    /* Section chips */
    .chip-row {
        display: flex;
        flex-wrap: wrap;
        gap: 6px;
        margin-top: 8px;
    }
    .chip-user {
        background: rgba(37,99,168,0.25);
        border: 1px solid rgba(37,99,168,0.5);
        color: #93c5fd;
        border-radius: 20px;
        padding: 3px 12px;
        font-size: 0.78rem;
    }
    .chip-auto {
        background: rgba(234,179,8,0.15);
        border: 1px solid rgba(234,179,8,0.4);
        color: #fde68a;
        border-radius: 20px;
        padding: 3px 12px;
        font-size: 0.78rem;
    }

    /* Input labels */
    label { color: #cbd5e1 !important; }

    /* Streamlit overrides */
    .stTextInput > div > div > input,
    .stTextInput input,
    input[type="text"],
    div[data-baseweb="input"] input,
    div[data-baseweb="base-input"] input {
        background: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        color: #1e293b !important;
        border-radius: 8px !important;
        -webkit-text-fill-color: #1e293b !important;
    }
    .stFileUploader {
        background: rgba(255,255,255,0.04) !important;
        border: 1px dashed rgba(255,255,255,0.20) !important;
        border-radius: 8px !important;
    }
    .stButton > button {
        background: linear-gradient(90deg, #1d4ed8, #2563eb) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        padding: 0.6rem 2rem !important;
        width: 100% !important;
        transition: all 0.2s !important;
    }
    .stButton > button:hover {
        background: linear-gradient(90deg, #1e40af, #1d4ed8) !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 16px rgba(37,99,168,0.4) !important;
    }
    .stDownloadButton > button {
        background: linear-gradient(90deg, #15803d, #16a34a) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        padding: 0.6rem 2rem !important;
        width: 100% !important;
    }

    /* Hide streamlit branding */
    #MainMenu, footer, header { visibility: hidden; }
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
</style>
""", unsafe_allow_html=True)

# ─── Template Path ───
TEMPLATES_DIR = os.path.join(os.path.dirname(__file__), "templates")
TEMPLATE_OPTIONS = {
    "TM-1 (Default)": "TM1.docx",
    "TM-2 (Custom)":  "TM2.docx",
}

def check_template(path):
    if not os.path.exists(path):
        st.markdown(f"""
        <div class="error-box">
        ❌ <strong>Template not found:</strong> {os.path.basename(path)}<br>
        Upload it to GitHub → <code>templates/</code> folder to enable.
        </div>""", unsafe_allow_html=True)
        return False
    return True


# ─── Header ───
st.markdown("""
<div class="mop-header">
    <h1>📋 Smart MOP Generator</h1>
    <p>Enterprise Telecom — Method of Procedure Automation</p>
</div>
""", unsafe_allow_html=True)

# ─── Privacy Notice ───
st.markdown("""
<div class="privacy-badge">
    🔒 &nbsp;<strong>Privacy First:</strong> Your uploaded files are processed in memory only.
    No data is stored, logged, or retained on any server after processing.
</div>
""", unsafe_allow_html=True)

# ─── Template Selector ───
st.markdown('<div class="mop-card">', unsafe_allow_html=True)
st.markdown('<div class="section-label">📄 Select Template</div>', unsafe_allow_html=True)

selected_label = st.radio(
    "Choose output MOP template:",
    options=list(TEMPLATE_OPTIONS.keys()),
    index=0,
    horizontal=True,
    help="TM-1 is default. To enable TM-2, upload TM2.docx to GitHub → templates/ folder."
)
selected_template_path = os.path.join(TEMPLATES_DIR, TEMPLATE_OPTIONS[selected_label])

if selected_label == "TM-2 (Custom)" and not os.path.exists(selected_template_path):
    st.markdown(
        '<div style="color:#fca5a5;font-size:0.85rem;margin-top:4px;">'
        '⚠️ TM-2 not available yet. Upload <code>TM2.docx</code> to GitHub → templates/ folder.</div>',
        unsafe_allow_html=True
    )

st.markdown('</div>', unsafe_allow_html=True)

if not check_template(selected_template_path):
    st.stop()

# ─── Input Section ───
st.markdown('<div class="mop-card">', unsafe_allow_html=True)
st.markdown('<div class="section-label">📝 MOP Details</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    activity_name = st.text_input(
        "Activity Name",
        placeholder="e.g. Fiber Cut Restoration",
        help="Name of the network activity"
    )
with col2:
    vendor_name = st.text_input(
        "Vendor Name",
        placeholder="e.g. Nokia",
        help="Vendor responsible for this activity"
    )

st.markdown('<div class="section-label" style="margin-top:16px;">📎 Upload Input MOP File</div>',
            unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Upload your MOP document",
    type=['docx', 'txt'],
    help="Upload your input MOP as .docx or .txt file. (.doc files: please convert to .docx first)",
    label_visibility="collapsed"
)

# .doc warning
if uploaded_file and uploaded_file.name.endswith('.doc') and not uploaded_file.name.endswith('.docx'):
    st.markdown("""
    <div class="error-box">
    ⚠️ <strong>.doc format detected.</strong> Please convert your file to <strong>.docx</strong> format first.<br>
    In Microsoft Word: File → Save As → Word Document (.docx)
    </div>
    """, unsafe_allow_html=True)
    uploaded_file = None

st.markdown('</div>', unsafe_allow_html=True)

# ─── Generate Button ───
generate_clicked = st.button("🚀 Generate MOP", use_container_width=True)

# ─── Processing ───
if generate_clicked:
    # Validation
    errors = []
    if not activity_name.strip():
        errors.append("Activity Name is required.")
    if not vendor_name.strip():
        errors.append("Vendor Name is required.")
    if not uploaded_file:
        errors.append("Please upload an input MOP file.")

    if errors:
        for err in errors:
            st.markdown(f'<div class="error-box">⚠️ {err}</div>', unsafe_allow_html=True)
    else:
        # Processing
        status_placeholder = st.empty()
        status_placeholder.markdown("""
        <div class="process-box">
            ⏳ <strong>Processing your MOP...</strong><br><br>
            ▸ Reading input document<br>
            ▸ Mapping headings to template sections<br>
            ▸ Inserting content into template<br>
            ▸ Embedding input as SOP attachment<br>
            ▸ Generating output file
        </div>
        """, unsafe_allow_html=True)

        tmp_dir = None
        try:
            tmp_dir = tempfile.mkdtemp()
            input_filename = uploaded_file.name
            input_path = os.path.join(tmp_dir, input_filename)

            # Save uploaded file to temp
            with open(input_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())

            # Build output filename
            act_safe = sanitize_filename(activity_name.strip())
            ven_safe = sanitize_filename(vendor_name.strip())
            output_filename = f"{act_safe}_{ven_safe}_MOP.docx"
            output_path = os.path.join(tmp_dir, output_filename)

            # Generate MOP
            result = generate_mop(
                template_path=selected_template_path,
                input_path=input_path,
                activity_name=activity_name.strip(),
                vendor_name=vendor_name.strip(),
                output_path=output_path
            )

            if result['success']:
                status_placeholder.empty()
                filled = result.get('filled_sections', [])
                auto   = result.get('autofilled_sections', [])
                sop_ref = result.get('sop_ref_path', '')

                filled_chips = ''.join([f'<span class="chip-user">✅ {s}</span>' for s in filled])
                auto_chips   = ''.join([f'<span class="chip-auto">🤖 {s}</span>' for s in auto])

                # Build ZIP in memory BEFORE tmp_dir cleanup
                import zipfile as _zf, io as _io
                safe_act     = sanitize_filename(activity_name.strip())
                zip_filename = f"{safe_act}_MOP_Package.zip"
                buf = _io.BytesIO()
                with _zf.ZipFile(buf, 'w', _zf.ZIP_DEFLATED) as zf:
                    zf.write(output_path, output_filename)
                    if sop_ref and os.path.exists(sop_ref):
                        zf.write(sop_ref, os.path.basename(sop_ref))
                
                # Store in session_state so download button survives re-render
                st.session_state['zip_bytes']    = buf.getvalue()
                st.session_state['zip_filename'] = zip_filename
                st.session_state['zip_sop_ref']  = os.path.basename(sop_ref) if sop_ref else ''
                st.session_state['zip_mop_name'] = output_filename
                st.session_state['filled']       = filled
                st.session_state['auto']         = auto

            # Show results from session_state (survives re-render)
            if 'zip_bytes' in st.session_state:
                filled       = st.session_state['filled']
                auto         = st.session_state['auto']
                zip_bytes    = st.session_state['zip_bytes']
                zip_filename = st.session_state['zip_filename']
                sop_ref_name = st.session_state['zip_sop_ref']
                output_filename = st.session_state['zip_mop_name']

                filled_chips2 = ''.join([f'<span class="chip-user">✅ {s}</span>' for s in filled])
                auto_chips2   = ''.join([f'<span class="chip-auto">🤖 {s}</span>' for s in auto])

                st.markdown(f"""
                <div class="success-box">
                    <h3>✅ MOP Generated Successfully!</h3>
                    <p>📦 ZIP contains:</p>
                    <p>&nbsp;&nbsp;📄 <strong>{output_filename}</strong> — Output MOP</p>
                    <p>&nbsp;&nbsp;📎 <strong>{sop_ref_name}</strong> — Full input copy (SOP Reference)</p>
                    <p>🔒 No data stored or retained on server</p>
                    <br>
                    <p><strong style="color:#93c5fd;">Sections from your input ({len(filled)}):</strong></p>
                    <div class="chip-row">{filled_chips2 if filled_chips2 else '<span style="color:#64748b;font-size:0.8rem;">None detected</span>'}</div>
                    <br>
                    <p><strong style="color:#fde68a;">Auto-filled sections ({len(auto)}):</strong></p>
                    <div class="chip-row">{auto_chips2 if auto_chips2 else '<span style="color:#64748b;font-size:0.8rem;">None</span>'}</div>
                </div>
                """, unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)
                st.download_button(
                    label=f"⬇️ Download {zip_filename}",
                    data=zip_bytes,
                    file_name=zip_filename,
                    mime="application/zip",
                    use_container_width=True
                )
            
            if False:  # dummy to close old if block
                pass

            if not result['success']:
                status_placeholder.empty()
                st.markdown(f"""
                <div class="error-box">
                    ❌ <strong>Generation Failed</strong><br>{result['message']}
                </div>
                """, unsafe_allow_html=True)

        except Exception as e:
            status_placeholder.empty()
            st.markdown(f"""
            <div class="error-box">
                ❌ <strong>Unexpected Error:</strong> {str(e)}<br>
                Please check your input file and try again.
            </div>
            """, unsafe_allow_html=True)

        finally:
            # Always cleanup temp files
            if tmp_dir and os.path.exists(tmp_dir):
                shutil.rmtree(tmp_dir, ignore_errors=True)

# ─── Footer ───
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown("""
<div style="text-align:center; color:#475569; font-size:0.78rem; padding: 12px;">
    Smart MOP Generator &nbsp;|&nbsp; Enterprise Telecom Tool &nbsp;|&nbsp;
    🔒 Zero Data Retention Policy
</div>
""", unsafe_allow_html=True)
