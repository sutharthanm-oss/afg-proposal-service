"""
AFG Proposal Generation Service
POST /generate -> runs overlay_engine, converts to PDF via LibreOffice, returns both files.

Deploy: Railway (Dockerfile below handles LibreOffice install).
Called by: Make.com, after the Claude API extraction step.
"""
import os
import subprocess
import uuid

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from typing import List, Optional

from overlay_engine import generate_proposal, list_templates

app = FastAPI(title="AFG Proposal Generation Service")

GENERATED_DIR = os.path.join(os.path.dirname(__file__), "generated")
os.makedirs(GENERATED_DIR, exist_ok=True)


class Prospect(BaseModel):
    name: str
    dob: str = "-"
    age: str = "-"
    smoking_status: str = "-"


class Tier(BaseModel):
    tier_label: str  # "A" | "B" | "C"
    monthly_premium: str = "-"
    death: str = "-"
    tpd: str = "-"
    terminal_illness: str = "-"
    ci_minus: str = "-"
    ci_plus: str = "-"
    simplified_ci: str = "-"
    waiver_policy_owner: str = "-"
    waiver_life_assured: str = "-"
    medical_card: str = "-"
    room_and_board: str = "-"
    annual_limit: str = "-"
    lifetime_limit: str = "-"
    co_insurance_deductible: str = "-"
    personal_accident: str = "-"
    coverage_up_to_age: str = "-"
    remarks: str = "-"


class GenerateRequest(BaseModel):
    template_id: str = Field(..., description="e.g. 'future_first' -- see GET /templates")
    agent_name: str
    prospect: Prospect
    tiers: List[Tier]


@app.get("/templates")
def get_templates():
    """Lets the bot/Make check which products can currently be generated."""
    return {"templates": list_templates()}


@app.post("/generate")
def generate(req: GenerateRequest):
    if req.template_id not in list_templates():
        raise HTTPException(400, f"Unknown template_id '{req.template_id}'. "
                                  f"Available: {list_templates()}")

    job_id = uuid.uuid4().hex[:10]
    safe_name = "".join(c for c in req.prospect.name if c.isalnum() or c in " -_").strip() or "Proposal"
    base_name = f"{safe_name} - {job_id}"
    pptx_path = os.path.join(GENERATED_DIR, f"{base_name}.pptx")

    data = req.dict()
    generate_proposal(req.template_id, data, pptx_path)

    # Convert to PDF via LibreOffice headless (installed in the Docker image)
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", GENERATED_DIR, pptx_path],
        check=True, timeout=60,
    )
    pdf_path = pptx_path.rsplit(".", 1)[0] + ".pdf"

    return {
        "job_id": job_id,
        "pptx_url": f"/files/{os.path.basename(pptx_path)}",
        "pdf_url": f"/files/{os.path.basename(pdf_path)}",
    }


@app.get("/files/{filename}")
def get_file(filename: str):
    path = os.path.join(GENERATED_DIR, filename)
    if not os.path.isfile(path):
        raise HTTPException(404, "File not found")
    return FileResponse(path)


@app.get("/health")
def health():
    return {"status": "ok", "templates": list_templates()}
