# Resume Screening MVP

FastAPI app to upload one or many resumes (PDF/TXT), extract key profile data, auto-tag designation by experience, and export batch results to Excel and PDF.

## Features
- Multi-file upload (PDF/TXT) with 10MB per-file limit
- Extracts core fields:
  - Full name
  - Email
  - Phone
  - Technical skills
  - Total years of experience
  - Highest degree
- Extra HSE fields from your format:
  - Designation (Inspector/Officer/Engineer/Manager)
  - NEBOSH (YES/NO)
  - ADOSH/OSHAD (YES/NO)
  - LEVEL 6 (YES/NO)
  - Nature of experience (rail, infrastructure, bridges, buildings, offshore, onshore, facility management)
- Requirement matching and score
- Export processed batch to:
  - Excel (`.xlsx`)
  - PDF (`.pdf`)
- SQLite persistence in `processed_cvs` table

## Run
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Open: `http://127.0.0.1:8000`

## Notes
- For best extraction accuracy, use text-based PDFs.
- DOCX parsing is not included in this MVP.
