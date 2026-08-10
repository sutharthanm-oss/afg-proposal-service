"""
AFG Proposal Generation Service
POST /generate -> runs overlay_engine, converts to PDF via LibreOffice, returns both files.
POST /generate-from-telegram -> downloads screenshots from Telegram, extracts data via
    Claude, then generates the proposal. This is the endpoint Make.com actually calls --
    it does all the heavy lifting (image download, vision API, JSON parsing) that Make's
    module set isn't well suited for, so Make just sends file_ids and gets back file URLs.

Deploy: Railway (Dockerfile below handles LibreOffice install).
Env vars required: ANTHROPIC_API_KEY, TELEGRAM_BOT_TOKEN
"""
import base64
import json
import os
import subprocess
import threading
import time
import uuid

import httpx
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field
from typing import List, Optional

from overlay_engine import generate_proposal, list_templates

app = FastAPI(title="AFG Proposal Generation Service")

GENERATED_DIR = os.path.join(os.path.dirname(__file__), "generated")
os.makedirs(GENERATED_DIR, exist_ok=True)

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN", "")
AIRTABLE_API_KEY = os.environ.get("AIRTABLE_API_KEY", "")
AIRTABLE_BASE_ID = "appCF1x7Gsju6Z3xy"
AIRTABLE_TABLE_ID = "tblzJe74G8yCdxJAJ"
# Airtable's REST API (default) keys/filters records by FIELD NAME, not field ID.
FLD_TELEGRAM_ID = "Telegram ID"
FLD_STATE = "State"
FLD_AGENT_NAME = "Agent Name"
FLD_SESSION_JSON = "Session Data JSON"

EXTRACTION_SYSTEM_PROMPT = """You are an extraction engine for AFG's insurance proposal automation. You will be shown
screenshots from the FWD Malaysia agent quoting app. Extract the data precisely and
return ONLY valid JSON matching the schema below -- no preamble, no markdown fences,
no commentary.

INPUT TYPES YOU MAY RECEIVE:
1. A "Who will be covered" profile screen -- shows name, DOB, age, gender, smoking
   habit, religion, nationality, occupation, mobile number, and which product(s)
   are covered.
2. A "Benefit illustration details" screen, "Base plan" tab -- shows the base
   product name, sum covered, and total/basic/initial contribution figures.
3. A "Benefit illustration details" screen, "Additional protection" tab -- shows
   one or more riders, each with its own sum covered and modal contribution.

RULES:
- If a field is not visible in any screenshot provided, output "-" for that field.
  Never guess or infer a number that isn't shown.
- "Class" is never shown in this app flow. Always output "-" for class.
- Monthly Premium = the base plan's "Total Contribution" figure, exactly as shown.
  Do NOT add rider contributions on top -- "Total Contribution" already includes
  every rider attached to that quote.
- Coverage up to age = person's current age + contribution/certificate term.
- Map product sum covered into these benefit rows by matching the product name:
  - "FWD Future First" (base) -> death, tpd, terminal_illness (same sum covered
    value for all three)
  - "FWD Critical Illness Rider" or "FWD CI First" (full/regular version) -> ci_plus
  - "FWD Critical Illness Lite Rider" (reduced condition list) -> ci_minus
  - "FWD Critical Illness Waiver of Contribution Rider" -> waiver_life_assured,
    value "Included" (NOT its sum covered figure)
  - Any product/rider you don't recognize -> put its name and sum covered in
    "unmapped_items" rather than forcing it into a row.

OUTPUT SCHEMA (return exactly this shape, as a single tier labeled "A"):
{
  "prospect": {"name": string, "dob": string, "age": string, "smoking_status": string},
  "tier": {
    "monthly_premium": string, "death": string, "tpd": string, "terminal_illness": string,
    "ci_minus": string, "ci_plus": string, "simplified_ci": string,
    "waiver_policy_owner": string, "waiver_life_assured": string, "medical_card": string,
    "room_and_board": string, "annual_limit": string, "lifetime_limit": string,
    "co_insurance_deductible": string, "personal_accident": string,
    "coverage_up_to_age": string, "remarks": string
  }
}"""


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
    base_name = f"Proposal - {safe_name} - {job_id}"
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


class TelegramGenerateRequest(BaseModel):
    template_id: str
    agent_name: str
    profile_file_id: str
    base_file_id: str
    rider_file_ids: List[str] = []


def _telegram_download_base64(file_id: str) -> tuple[str, str]:
    """Returns (base64_data, media_type) for a Telegram file_id."""
    if not TELEGRAM_BOT_TOKEN:
        raise HTTPException(500, "TELEGRAM_BOT_TOKEN not configured on server")

    with httpx.Client(timeout=30) as client:
        r = client.get(
            f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getFile",
            params={"file_id": file_id},
        )
        r.raise_for_status()
        result = r.json().get("result", {})
        file_path = result.get("file_path")
        if not file_path:
            raise HTTPException(502, f"Telegram getFile failed for {file_id}: {r.text}")

        img = client.get(
            f"https://api.telegram.org/file/bot{TELEGRAM_BOT_TOKEN}/{file_path}"
        )
        img.raise_for_status()

    media_type = "image/png" if file_path.lower().endswith(".png") else "image/jpeg"
    return base64.b64encode(img.content).decode("utf-8"), media_type


def _call_claude_extraction(images_b64: List[tuple[str, str]]) -> dict:
    if not ANTHROPIC_API_KEY:
        raise HTTPException(500, "ANTHROPIC_API_KEY not configured on server")

    content = []
    for b64, media_type in images_b64:
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": media_type, "data": b64},
        })
    content.append({"type": "text", "text": "Extract the data from these screenshots per your instructions."})

    with httpx.Client(timeout=60) as client:
        r = client.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": ANTHROPIC_API_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json",
            },
            json={
                "model": "claude-sonnet-4-6",
                "max_tokens": 2000,
                "system": EXTRACTION_SYSTEM_PROMPT,
                "messages": [{"role": "user", "content": content}],
            },
        )
        r.raise_for_status()
        data = r.json()

    text = "".join(block.get("text", "") for block in data.get("content", []) if block.get("type") == "text")
    text = text.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    try:
        return json.loads(text.strip())
    except json.JSONDecodeError as e:
        raise HTTPException(502, f"Claude returned unparseable JSON: {e}. Raw: {text[:500]}")


@app.post("/generate-from-telegram")
def generate_from_telegram(req: TelegramGenerateRequest):
    if req.template_id not in list_templates():
        raise HTTPException(400, f"Unknown template_id '{req.template_id}'. "
                                  f"Available: {list_templates()}")

    # Download all screenshots from Telegram
    images = [_telegram_download_base64(req.profile_file_id)]
    images.append(_telegram_download_base64(req.base_file_id))
    for rid in req.rider_file_ids:
        images.append(_telegram_download_base64(rid))

    # Extract structured data via Claude
    extracted = _call_claude_extraction(images)
    prospect = extracted["prospect"]
    tier = extracted["tier"]
    tier["tier_label"] = "A"

    # Generate the proposal using the same engine as /generate
    job_id = uuid.uuid4().hex[:10]
    safe_name = "".join(c for c in prospect["name"] if c.isalnum() or c in " -_").strip() or "Proposal"
    base_name = f"Proposal - {safe_name} - {job_id}"
    pptx_path = os.path.join(GENERATED_DIR, f"{base_name}.pptx")

    data = {"prospect": prospect, "agent_name": req.agent_name, "tiers": [tier]}
    generate_proposal(req.template_id, data, pptx_path)

    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", GENERATED_DIR, pptx_path],
        check=True, timeout=60,
    )
    pdf_path = pptx_path.rsplit(".", 1)[0] + ".pdf"

    return {
        "job_id": job_id,
        "extracted": extracted,
        "pptx_url": f"/files/{os.path.basename(pptx_path)}",
        "pdf_url": f"/files/{os.path.basename(pdf_path)}",
    }


@app.get("/files/{filename}")
def get_file(filename: str):
    path = os.path.join(GENERATED_DIR, filename)
    if not os.path.isfile(path):
        raise HTTPException(404, "File not found")
    return FileResponse(path)


def _parse_session_data(raw: str) -> dict:
    """Parses the semicolon-delimited 'key:value;key:value' session string
    the Make bot writes into Session Data JSON. Screenshots are grouped into
    tiers IN THE ORDER THEY WERE UPLOADED: each 'base_photo' starts a new
    tier, and any 'rider_photo' entries before the next 'base_photo' belong
    to that tier. This lets an agent upload any number of quotes (1-3) with
    any number of riders each, with no counting logic needed on the Make side.
    """
    result = {"tiers": []}
    current_tier = None
    for segment in raw.split(";"):
        if ":" not in segment:
            continue
        key, _, value = segment.partition(":")
        if key == "base_photo":
            current_tier = {"base_photo": value, "rider_photos": []}
            result["tiers"].append(current_tier)
        elif key == "rider_photo" and current_tier is not None:
            current_tier["rider_photos"].append(value)
        else:
            result[key] = value
    return result


def _telegram_send_document(chat_id: str, file_url: str):
    with httpx.Client(timeout=30) as client:
        r = client.post(
            f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument",
            data={"chat_id": chat_id, "document": file_url},
        )
        r.raise_for_status()


def _telegram_send_message(chat_id: str, text: str):
    with httpx.Client(timeout=30) as client:
        r = client.post(
            f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage",
            data={"chat_id": chat_id, "text": text},
        )
        r.raise_for_status()


def _airtable_headers():
    return {"Authorization": f"Bearer {AIRTABLE_API_KEY}", "Content-Type": "application/json"}


def _airtable_update_state(record_id: str, state: str):
    with httpx.Client(timeout=15) as client:
        r = client.patch(
            f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_TABLE_ID}/{record_id}",
            headers=_airtable_headers(),
            json={"fields": {FLD_STATE: state}, "typecast": True},
        )
        if r.status_code >= 400:
            print(f"[airtable] update FAILED for {record_id} -> {state}: {r.status_code} {r.text}", flush=True)
        else:
            print(f"[airtable] update OK for {record_id} -> {state}", flush=True)


def _process_one_record(record: dict):
    fields = record["fields"]
    record_id = record["id"]
    telegram_id = fields.get(FLD_TELEGRAM_ID)
    agent_name = fields.get(FLD_AGENT_NAME, "Agent")
    session_raw = fields.get(FLD_SESSION_JSON, "")

    # Claim the record immediately so no other poll cycle double-processes it.
    _airtable_update_state(record_id, "GENERATING")

    try:
        parsed = _parse_session_data(session_raw)
        profile_id = parsed.get("profile_photo")
        tiers_raw = parsed.get("tiers", [])

        if not profile_id or not tiers_raw:
            _telegram_send_message(telegram_id, "Something went wrong — missing screenshots. Please type /start to try again.")
            _airtable_update_state(record_id, "ERROR")
            return

        profile_image = _telegram_download_base64(profile_id)
        tier_labels = ["A", "B", "C"]
        tiers_data = []
        prospect = None

        for i, tier_info in enumerate(tiers_raw[:3]):  # template supports up to 3 columns
            base_image = _telegram_download_base64(tier_info["base_photo"])
            images = [profile_image, base_image]
            for rid in tier_info["rider_photos"]:
                images.append(_telegram_download_base64(rid))

            extracted = _call_claude_extraction(images)
            if prospect is None:
                prospect = extracted["prospect"]
            tier = extracted["tier"]
            tier["tier_label"] = tier_labels[i]
            tiers_data.append(tier)

        job_id = uuid.uuid4().hex[:10]
        safe_name = "".join(c for c in prospect["name"] if c.isalnum() or c in " -_").strip() or "Proposal"
        base_name = f"Proposal - {safe_name} - {job_id}"
        pptx_path = os.path.join(GENERATED_DIR, f"{base_name}.pptx")

        data = {"prospect": prospect, "agent_name": agent_name, "tiers": tiers_data}
        generate_proposal("future_first", data, pptx_path)

        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", GENERATED_DIR, pptx_path],
            check=True, timeout=60,
        )
        pdf_path = pptx_path.rsplit(".", 1)[0] + ".pdf"

        public_base = os.environ.get("PUBLIC_BASE_URL", "https://afg-proposal-service-production.up.railway.app")
        # PDF only for now — PPTX delivery deferred to a later phase to keep this simpler.
        _telegram_send_document(telegram_id, f"{public_base}/files/{os.path.basename(pdf_path)}")

        _airtable_update_state(record_id, "DONE")
    except Exception as e:
        print(f"[process] error for record {record_id}: {e}", flush=True)
        try:
            _telegram_send_message(telegram_id, f"Something went wrong generating your proposal. Please type /start to try again.")
        except Exception:
            pass
        _airtable_update_state(record_id, "ERROR")


def _poll_loop():
    print("[poller] background thread started", flush=True)
    cycle = 0
    while True:
        cycle += 1
        try:
            if AIRTABLE_API_KEY:
                with httpx.Client(timeout=15) as client:
                    r = client.get(
                        f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_TABLE_ID}",
                        headers=_airtable_headers(),
                        params={"filterByFormula": f"{{{FLD_STATE}}}='READY_TO_GENERATE'"},
                    )
                    r.raise_for_status()
                    records = r.json().get("records", [])
                    if cycle % 15 == 1:
                        print(f"[poller] cycle {cycle}: {len(records)} record(s) matched", flush=True)
                    for record in records:
                        print(f"[poller] processing record {record['id']}", flush=True)
                        _process_one_record(record)
        except Exception as e:
            print(f"[poller] error: {e}", flush=True)
        time.sleep(4)


@app.on_event("startup")
def _start_poller():
    if AIRTABLE_API_KEY:
        threading.Thread(target=_poll_loop, daemon=True).start()


@app.get("/health")
def health():
    return {"status": "ok", "templates": list_templates(), "poller_enabled": bool(AIRTABLE_API_KEY)}
