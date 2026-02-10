from __future__ import annotations

import json
import re
import sqlite3
import uuid
from dataclasses import dataclass
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import Workbook
from PyPDF2 import PdfReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = BASE_DIR / "data" / "cv_results.db"
MAX_FILE_BYTES = 10 * 1024 * 1024
ALLOWED_TYPES = {"application/pdf", "text/plain"}
ALLOWED_SUFFIXES = {".pdf", ".txt"}

DEFAULT_SKILLS = {
    "python",
    "sql",
    "excel",
    "power bi",
    "autocad",
    "nebosh",
    "adosh",
    "oshad",
    "safety auditing",
    "risk assessment",
    "incident investigation",
    "hse",
    "ohs",
    "project management",
}

DEGREE_KEYWORDS = [
    "phd",
    "doctorate",
    "master",
    "bachelor",
    "diploma",
    "nvq level 6",
    "level 6",
    "high school",
]

ROLE_RULES = [
    {"role": "HSE/Safety Inspector", "experience_min": 0, "experience_max": 5},
    {"role": "HSE/Safety Officer", "experience_min": 5, "experience_max": 10},
    {"role": "HSE/Safety Engineer", "experience_min": 10, "experience_max": 15},
    {"role": "HSE/Safety Manager", "experience_min": 15, "experience_max": 100},
]

NATURE_KEYWORDS = {
    "rail": ["rail"],
    "infrastructure": ["infrastructure"],
    "bridges": ["bridge", "bridges"],
    "buildings": ["building", "buildings", "villa"],
    "offshore": ["offshore"],
    "onshore": ["onshore"],
    "facility management": ["facility management", "facilities management"],
}


@dataclass
class ResumeData:
    full_name: str
    email: str
    phone: str
    skills: list[str]
    total_experience_years: float
    highest_degree: str
    certifications: dict[str, bool]
    nature_of_experience: list[str]
    suggested_role: str


app = FastAPI(title="Resume Screening MVP")
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


def init_db() -> None:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS processed_cvs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                batch_id TEXT,
                filename TEXT NOT NULL,
                extracted_data TEXT NOT NULL,
                match_score REAL NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )
        columns = [row[1] for row in conn.execute("PRAGMA table_info(processed_cvs)").fetchall()]
        if "batch_id" not in columns:
            conn.execute("ALTER TABLE processed_cvs ADD COLUMN batch_id TEXT")
        conn.commit()


@app.on_event("startup")
def startup() -> None:
    init_db()


@app.get("/", response_class=HTMLResponse)
def index(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(
        "index.html",
        {"request": request, "result": None, "errors": []},
    )


@app.post("/screen", response_class=HTMLResponse)
async def screen_resume(
    request: Request,
    resume_files: list[UploadFile] = File(...),
    job_description: str = Form(""),
    required_skills: str = Form(""),
    min_experience: float = Form(0),
    required_education: str = Form(""),
) -> HTMLResponse:
    errors: list[str] = []
    if not resume_files:
        errors.append("Please upload at least one resume.")

    requirements = parse_requirements(job_description, required_skills, min_experience, required_education)
    batch_id = str(uuid.uuid4())
    rows: list[dict[str, Any]] = []

    for resume_file in resume_files:
        file_ext = Path(resume_file.filename or "").suffix.lower()
        if file_ext not in ALLOWED_SUFFIXES and resume_file.content_type not in ALLOWED_TYPES:
            errors.append(f"{resume_file.filename}: only PDF and TXT files are supported.")
            continue

        content = await resume_file.read()
        if len(content) > MAX_FILE_BYTES:
            errors.append(f"{resume_file.filename}: exceeds 10MB limit.")
            continue

        text = extract_text(content, file_ext)
        if not text.strip():
            errors.append(f"{resume_file.filename}: unable to extract text.")
            continue

        candidate = extract_candidate_data(text)
        match_result = calculate_match(candidate, requirements)
        save_result(batch_id, resume_file.filename or "unknown", candidate, match_result["score"])
        rows.append(build_display_row(resume_file.filename or "unknown", candidate, match_result))

    result = None
    if rows:
        average = round(sum(row["match_score"] for row in rows) / len(rows), 2)
        result = {
            "batch_id": batch_id,
            "rows": rows,
            "count": len(rows),
            "average_score": average,
            "requirements": requirements,
        }

    return templates.TemplateResponse(
        "index.html",
        {"request": request, "result": result, "errors": errors},
    )


@app.get("/export/excel/{batch_id}")
def export_excel(batch_id: str) -> StreamingResponse:
    rows = fetch_batch_rows(batch_id)
    if not rows:
        raise HTTPException(status_code=404, detail="No records found for this batch.")

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Resume Screening"
    headers = [
        "Name",
        "Email",
        "Designation",
        "Exp",
        "NEBOSH",
        "ADOSH",
        "L6",
        "Nature of Exp",
    ]
    sheet.append(headers)

    for row in rows:
        sheet.append(
            [
                row["full_name"],
                row["email"],
                row["designation"],
                row["experience"],
                row["nebosh"],
                row["adosh"],
                row["level_6"],
                row["nature_of_experience"],
            ]
        )

    stream = BytesIO()
    workbook.save(stream)
    stream.seek(0)
    return StreamingResponse(
        stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=resume-screening-{batch_id}.xlsx"},
    )


@app.get("/export/pdf/{batch_id}")
def export_pdf(batch_id: str) -> StreamingResponse:
    rows = fetch_batch_rows(batch_id)
    if not rows:
        raise HTTPException(status_code=404, detail="No records found for this batch.")

    stream = BytesIO()
    doc = SimpleDocTemplate(stream, pagesize=landscape(A4), leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20)

    data = [["Name", "Email", "Designation", "Exp", "NEBOSH", "ADOSH", "L6", "Nature of Exp"]]
    for row in rows:
        data.append(
            [
                row["full_name"],
                row["email"],
                row["designation"],
                f"{row['experience']} Y",
                row["nebosh"],
                row["adosh"],
                row["level_6"],
                row["nature_of_experience"],
            ]
        )

    table = Table(data, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e5e7eb")),
                ("GRID", (0, 0), (-1, -1), 0.7, colors.black),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f9fafb")]),
            ]
        )
    )
    doc.build([table])
    stream.seek(0)

    return StreamingResponse(
        stream,
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename=resume-screening-{batch_id}.pdf"},
    )


def extract_text(content: bytes, file_ext: str) -> str:
    if file_ext == ".txt":
        return content.decode("utf-8", errors="ignore")
    if file_ext == ".pdf":
        reader = PdfReader(BytesIO(content))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    return ""


def extract_candidate_data(text: str) -> ResumeData:
    lowered = text.lower()
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    email = first_match(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text)
    phone = first_match(r"(?:\+?\d{1,3}[ -]?)?(?:\(?\d{2,4}\)?[ -]?)?\d{3,4}[ -]?\d{3,4}", text)
    full_name = extract_name(lines, email)

    skills = sorted([skill for skill in DEFAULT_SKILLS if skill in lowered])
    experience = extract_experience_years(lowered)
    degree = extract_highest_degree(lowered)

    certifications = {
        "nebosh": "nebosh" in lowered,
        "level_6": any(keyword in lowered for keyword in ["level 6", "nvq level 6", "othm", "nebosh diploma"]),
        "adosh_oshad": any(keyword in lowered for keyword in ["adosh", "oshad"]),
    }

    nature = []
    for label, keywords in NATURE_KEYWORDS.items():
        if any(keyword in lowered for keyword in keywords):
            nature.append(label)

    suggested_role = suggest_role(experience)

    return ResumeData(
        full_name=full_name,
        email=email,
        phone=phone,
        skills=skills,
        total_experience_years=experience,
        highest_degree=degree,
        certifications=certifications,
        nature_of_experience=nature,
        suggested_role=suggested_role,
    )


def build_display_row(filename: str, candidate: ResumeData, match_result: dict[str, Any]) -> dict[str, Any]:
    inspector_highlight = candidate.suggested_role == "HSE/Safety Inspector"
    preferred_band = preferred_band_label(candidate.total_experience_years, candidate.suggested_role)
    return {
        "filename": filename,
        "full_name": candidate.full_name,
        "email": candidate.email,
        "phone": candidate.phone,
        "designation": candidate.suggested_role,
        "experience": candidate.total_experience_years,
        "nebosh": "YES" if candidate.certifications["nebosh"] else "NO",
        "adosh": "YES" if candidate.certifications["adosh_oshad"] else "NO",
        "level_6": "YES" if candidate.certifications["level_6"] else "NO",
        "nature_of_experience": ", ".join(candidate.nature_of_experience) if candidate.nature_of_experience else "Not detected",
        "skills": ", ".join(candidate.skills) if candidate.skills else "Not detected",
        "degree": candidate.highest_degree,
        "match_score": match_result["score"],
        "missing_requirements": ", ".join(match_result["missing_requirements"]) if match_result["missing_requirements"] else "None",
        "matching_skills": ", ".join(match_result["matching_skills"]) if match_result["matching_skills"] else "None",
        "inspector_highlight": inspector_highlight,
        "preferred_band": preferred_band,
    }


def extract_name(lines: list[str], email: str) -> str:
    for line in lines[:8]:
        if email and email in line:
            continue
        if re.search(r"\d", line):
            continue
        tokens = line.split()
        if 2 <= len(tokens) <= 4 and all(token[0].isupper() for token in tokens if token and token[0].isalpha()):
            return line
    return lines[0] if lines else "Unknown"


def first_match(pattern: str, text: str) -> str:
    found = re.search(pattern, text)
    return found.group(0).strip() if found else "Not found"


def extract_experience_years(text: str) -> float:
    ranges = re.findall(r"(\d+(?:\.\d+)?)\s*(?:-|to)\s*(\d+(?:\.\d+)?)\s*(?:years|yrs)?", text)
    range_max = [float(end) for _, end in ranges]
    patterns = [
        r"(\d+(?:\.\d+)?)\+?\s*(?:years|yrs)\s*(?:of)?\s*experience",
        r"experience\s*[:\-]?\s*(\d+(?:\.\d+)?)",
        r"(\d+(?:\.\d+)?)\s*(?:years|yrs)",
    ]
    singles = []
    for pattern in patterns:
        singles.extend(float(val) for val in re.findall(pattern, text))
    candidates = range_max + singles
    return max(candidates) if candidates else 0.0


def extract_highest_degree(text: str) -> str:
    rank = {
        "phd": 8,
        "doctorate": 8,
        "master": 7,
        "bachelor": 6,
        "diploma": 5,
        "nvq level 6": 5,
        "level 6": 5,
        "high school": 1,
    }
    found = [keyword for keyword in DEGREE_KEYWORDS if keyword in text]
    if not found:
        return "Not found"
    return max(found, key=lambda k: rank.get(k, 0)).title()


def suggest_role(experience: float) -> str:
    for rule in ROLE_RULES:
        if rule["experience_min"] <= experience < rule["experience_max"]:
            return rule["role"]
    if experience >= 15:
        return "HSE/Safety Manager"
    return "Unclassified"


def preferred_band_label(experience: float, role: str) -> str:
    if role == "HSE/Safety Inspector" and experience < 5:
        return "Inspector (highlight)"
    if role == "HSE/Safety Officer" and 5 <= experience <= 7:
        return "Officer (5-7 preferred)"
    if role == "HSE/Safety Engineer" and 10 <= experience <= 15:
        return "Engineer (10-15 preferred)"
    return "Standard band"


def parse_requirements(
    job_description: str,
    required_skills: str,
    min_experience: float,
    required_education: str,
) -> dict[str, Any]:
    jd_lower = job_description.lower()
    explicit_skills = [skill.strip().lower() for skill in required_skills.split(",") if skill.strip()]
    jd_skills = [skill for skill in DEFAULT_SKILLS if skill in jd_lower]

    education = required_education.strip() or infer_education_from_jd(jd_lower)
    custom_flags = {
        "nebosh_required": "nebosh" in jd_lower,
        "level_6_required": "level 6" in jd_lower or "nvq level 6" in jd_lower,
        "adosh_oshad_required": "adosh" in jd_lower or "oshad" in jd_lower,
    }
    required_nature = [label for label, keys in NATURE_KEYWORDS.items() if any(k in jd_lower for k in keys)]

    return {
        "skills": sorted(set(explicit_skills + jd_skills)),
        "min_experience": min_experience,
        "education": education,
        "custom_flags": custom_flags,
        "required_nature": required_nature,
    }


def infer_education_from_jd(jd_lower: str) -> str:
    for degree in ["phd", "master", "bachelor", "diploma", "level 6"]:
        if degree in jd_lower:
            return degree.title()
    return ""


def calculate_match(candidate: ResumeData, requirements: dict[str, Any]) -> dict[str, Any]:
    required_skills = requirements["skills"]
    matching_skills = [skill for skill in required_skills if skill in candidate.skills]
    missing_skills = [skill for skill in required_skills if skill not in candidate.skills]

    skill_score = (len(matching_skills) / len(required_skills) * 100) if required_skills else 100
    exp_met = candidate.total_experience_years >= requirements["min_experience"]

    education_required = requirements["education"].lower()
    education_met = (not education_required) or (education_required in candidate.highest_degree.lower())

    certs_met = {
        "nebosh": not requirements["custom_flags"]["nebosh_required"] or candidate.certifications["nebosh"],
        "level_6": not requirements["custom_flags"]["level_6_required"] or candidate.certifications["level_6"],
        "adosh_oshad": not requirements["custom_flags"]["adosh_oshad_required"] or candidate.certifications["adosh_oshad"],
    }

    nature_required = requirements["required_nature"]
    matching_nature = [n for n in nature_required if n in candidate.nature_of_experience]
    nature_score = (len(matching_nature) / len(nature_required) * 100) if nature_required else 100

    checks = [
        skill_score,
        100 if exp_met else 0,
        100 if education_met else 0,
        (sum(certs_met.values()) / 3) * 100,
        nature_score,
    ]
    score = round(sum(checks) / len(checks), 2)

    missing_requirements = []
    if not exp_met:
        missing_requirements.append(f"Minimum experience: {requirements['min_experience']} years")
    if not education_met and requirements["education"]:
        missing_requirements.append(f"Education: {requirements['education']}")
    if not certs_met["nebosh"]:
        missing_requirements.append("NEBOSH")
    if not certs_met["level_6"]:
        missing_requirements.append("Level 6")
    if not certs_met["adosh_oshad"]:
        missing_requirements.append("ADOSH/OSHAD")
    if missing_skills:
        missing_requirements.append("Skills: " + ", ".join(missing_skills))

    return {
        "score": score,
        "matching_skills": matching_skills,
        "missing_skills": missing_skills,
        "experience_met": exp_met,
        "education_met": education_met,
        "certifications_met": certs_met,
        "matching_nature": matching_nature,
        "missing_requirements": missing_requirements,
    }


def save_result(batch_id: str, filename: str, candidate: ResumeData, match_score: float) -> None:
    payload = {
        "full_name": candidate.full_name,
        "email": candidate.email,
        "phone": candidate.phone,
        "skills": candidate.skills,
        "total_experience_years": candidate.total_experience_years,
        "highest_degree": candidate.highest_degree,
        "certifications": candidate.certifications,
        "nature_of_experience": candidate.nature_of_experience,
        "suggested_role": candidate.suggested_role,
    }
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            INSERT INTO processed_cvs (batch_id, filename, extracted_data, match_score, created_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            (batch_id, filename, json.dumps(payload), match_score, datetime.now(timezone.utc).isoformat()),
        )
        conn.commit()


def fetch_batch_rows(batch_id: str) -> list[dict[str, Any]]:
    with sqlite3.connect(DB_PATH) as conn:
        records = conn.execute(
            "SELECT filename, extracted_data, match_score FROM processed_cvs WHERE batch_id = ? ORDER BY id",
            (batch_id,),
        ).fetchall()

    rows = []
    for filename, extracted_data, match_score in records:
        data = json.loads(extracted_data)
        rows.append(
            {
                "filename": filename,
                "full_name": data.get("full_name", "Unknown"),
                "email": data.get("email", "Not found"),
                "designation": data.get("suggested_role", "Unclassified"),
                "experience": data.get("total_experience_years", 0),
                "nebosh": "YES" if data.get("certifications", {}).get("nebosh") else "NO",
                "adosh": "YES" if data.get("certifications", {}).get("adosh_oshad") else "NO",
                "level_6": "YES" if data.get("certifications", {}).get("level_6") else "NO",
                "nature_of_experience": ", ".join(data.get("nature_of_experience", [])) or "Not detected",
                "match_score": match_score,
                "missing_requirements": "See UI details",
                "inspector_highlight": "YES" if data.get("suggested_role") == "HSE/Safety Inspector" else "NO",
                "preferred_band": preferred_band_label(
                    float(data.get("total_experience_years", 0)),
                    data.get("suggested_role", "Unclassified"),
                ),
            }
        )
    return rows
