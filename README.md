# Smart MOP Generator

Enterprise Telecom — Method of Procedure Automation Tool

## Project Structure

```
smart-mop-generator/
├── app.py                 ← Streamlit UI
├── mop_engine.py          ← Core processing engine
├── requirements.txt       ← Python dependencies
└── templates/
    └── TM1.docx           ← Default MOP template (DO NOT MODIFY)
```

## Local Setup (Run on your laptop)

```bash
# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
```

App will open at: http://localhost:8501

## Deploy to Streamlit Cloud (Free, Shareable Link)

### Step 1 — GitHub Setup
1. Go to https://github.com and create account (free)
2. Create new **Private** repository named `smart-mop-generator`
3. Upload all files:
   - `app.py`
   - `mop_engine.py`
   - `requirements.txt`
   - `templates/TM1.docx`

### Step 2 — Streamlit Cloud
1. Go to https://share.streamlit.io
2. Sign in with GitHub
3. Click "New app"
4. Select your repository → `app.py`
5. Click "Deploy"
6. Get your shareable link!

## How It Works

1. User provides Activity Name, Vendor Name, uploads input MOP (.docx or .txt)
2. System extracts content from each section of input file
3. Smart heading mapping matches user headings to template placeholders
4. User content is inserted into correct template sections
5. Missing sections are auto-filled with generic, professional text
6. Original input file is embedded as SOP attachment (proof of input)
7. Output: `ActivityName_VendorName_MOP.docx`

## Features

- ✅ Template structure never modified
- ✅ Smart heading synonym matching
- ✅ Images & tables preserved from input
- ✅ Missing sections auto-filled (activity + vendor name included)
- ✅ Input file embedded as SOP proof attachment
- ✅ Zero data retention — all files deleted after processing
- ✅ Professional UI

## Supported Input Formats

- `.docx` — Microsoft Word (recommended, supports images)
- `.txt` — Plain text
- `.doc` — Not supported, please convert to .docx first

## Privacy

- No database
- No logging
- No file storage
- Temp files auto-deleted after each generation
- Zero data retained on server
