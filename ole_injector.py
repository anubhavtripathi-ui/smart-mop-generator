"""OLE object injector for SOP section."""
import os, re, zipfile


def embed_ole_in_sop(docx_path, attach_path, activity_name):
    """
    Embed input file as a ZIP attachment inside the output docx.
    Reliable cross-platform approach — no OLE rendering issues.
    """
    file_name = os.path.basename(attach_path)
    tmp_path = docx_path + '.oletmp'

    with open(attach_path, 'rb') as f:
        file_data = f.read()

    try:
        with zipfile.ZipFile(docx_path, 'r') as zin:
            with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                seen = set()
                for item in zin.infolist():
                    if item.filename not in seen:
                        zout.writestr(item, zin.read(item.filename))
                        seen.add(item.filename)
                # Embed input file
                embed_name = f'word/attachments/{file_name}'
                if embed_name not in seen:
                    zout.writestr(embed_name, file_data)
        os.replace(tmp_path, docx_path)
    except Exception:
        if os.path.exists(tmp_path):
            try: os.remove(tmp_path)
            except: pass


def _simple_embed(docx_path, attach_path, file_name, file_data):
    tmp = docx_path + '.fb'
    try:
        with zipfile.ZipFile(docx_path, 'r') as zin:
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
                seen = set()
                for item in zin.infolist():
                    if item.filename not in seen:
                        zout.writestr(item, zin.read(item.filename))
                        seen.add(item.filename)
                zout.writestr(f'word/attachments/{file_name}', file_data)
        os.replace(tmp, docx_path)
    except Exception:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except Exception:
                pass
