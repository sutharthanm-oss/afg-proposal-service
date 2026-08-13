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
from datetime import datetime
from urllib.parse import quote

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

# Permanent agent roster (separate base) -- registered once at onboarding by Make,
# updated here on every completed proposal to track usage.
AGENTS_BASE_ID = "appoXG0DH4k5onBmd"
AGENTS_TABLE_ID = "tbl4pcT77XPif1RpV"
AGENTS_FLD_TELEGRAM_ID = "Telegram ID"
AGENTS_FLD_TOTAL_PROPOSALS = "Total Proposals Generated"
AGENTS_FLD_LAST_ACTIVE = "Last Active"

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
  - "FWD Medical Rider" -> this rider has no single sum-covered figure; instead
    populate FIVE separate fields:
    medical_card = "Included", room_and_board = its Room & Board figure (as
    shown on screen), annual_limit = its Overall Annual Limit figure (as shown
    on screen), co_insurance_deductible = its Deductible figure (as shown on
    screen), lifetime_limit = "Unlimited". Also, whenever this rider is present,
    append this exact sentence to remarks (add it after any existing remarks
    text, don't replace them): "Medical Card rider premium will increase every
    5 years (following your age band)."
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


def _build_filename(name: str) -> str:
    """Proposal - Name - DDMMYY v1.0 (v2.0, v3.0... if the same person already
    has a proposal generated today)."""
    safe_name = "".join(c for c in name if c.isalnum() or c in " -_").strip() or "Proposal"
    date_str = datetime.now().strftime("%d%m%y")
    version = 1
    while True:
        candidate = f"Proposal - {safe_name} - {date_str} v{version}.0"
        if not os.path.exists(os.path.join(GENERATED_DIR, f"{candidate}.pptx")):
            return candidate
        version += 1


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

    base_name = _build_filename(req.prospect.name)
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
        "filename": base_name,
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
    base_name = _build_filename(prospect["name"])
    pptx_path = os.path.join(GENERATED_DIR, f"{base_name}.pptx")

    data = {"prospect": prospect, "agent_name": req.agent_name, "tiers": [tier]}
    generate_proposal(req.template_id, data, pptx_path)

    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", GENERATED_DIR, pptx_path],
        check=True, timeout=60,
    )
    pdf_path = pptx_path.rsplit(".", 1)[0] + ".pdf"

    return {
        "filename": base_name,
        "extracted": extracted,
        "pptx_url": f"/files/{os.path.basename(pptx_path)}",
        "pdf_url": f"/files/{os.path.basename(pdf_path)}",
    }


@app.get("/files/{filename}")
def get_file(filename: str):
    path = os.path.join(GENERATED_DIR, filename)
    if not os.path.isfile(path):
        raise HTTPException(404, "File not found")
    # FastAPI/Starlette can't always guess the right MIME type for .pptx on
    # every host, and Telegram's URL-fetcher rejects documents with a wrong
    # or missing Content-Type -- set it explicitly for the types we serve.
    media_types = {
        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".pdf": "application/pdf",
    }
    ext = os.path.splitext(filename)[1].lower()
    media_type = media_types.get(ext)
    return FileResponse(path, media_type=media_type, filename=filename)


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


def _telegram_send_document_by_path(chat_id: str, file_path: str):
    """Uploads the file directly (multipart), rather than giving Telegram a
    URL to fetch -- Telegram's URL-fetch path is picky about Content-Type
    for less-common formats (e.g. .pptx) and can reject valid files with
    'wrong type of the web page content'. Direct upload sidesteps that
    entirely and is the standard, most reliable way to send documents."""
    filename = os.path.basename(file_path)
    mime_map = {
        ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".pdf": "application/pdf",
    }
    ext = os.path.splitext(filename)[1].lower()
    mime_type = mime_map.get(ext, "application/octet-stream")

    with httpx.Client(timeout=60) as client:
        with open(file_path, "rb") as f:
            r = client.post(
                f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument",
                data={"chat_id": chat_id},
                files={"document": (filename, f, mime_type)},
            )
        if r.status_code >= 400:
            print(f"[telegram] sendDocument FAILED for {filename}: {r.status_code} {r.text}", flush=True)
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


def _track_agent_usage(telegram_id: str):
    """Increments Total Proposals Generated and stamps Last Active on the
    agent's permanent roster record. Best-effort -- a failure here must never
    block delivery of an already-generated proposal."""
    try:
        with httpx.Client(timeout=15) as client:
            r = client.get(
                f"https://api.airtable.com/v0/{AGENTS_BASE_ID}/{AGENTS_TABLE_ID}",
                headers=_airtable_headers(),
                params={"filterByFormula": f"{{{AGENTS_FLD_TELEGRAM_ID}}}='{telegram_id}'", "maxRecords": 1},
            )
            r.raise_for_status()
            records = r.json().get("records", [])
            if not records:
                print(f"[usage] no Agents record found for {telegram_id}, skipping", flush=True)
                return
            agent_record = records[0]
            current_total = agent_record["fields"].get(AGENTS_FLD_TOTAL_PROPOSALS, 0) or 0
            client.patch(
                f"https://api.airtable.com/v0/{AGENTS_BASE_ID}/{AGENTS_TABLE_ID}/{agent_record['id']}",
                headers=_airtable_headers(),
                json={"fields": {
                    AGENTS_FLD_TOTAL_PROPOSALS: current_total + 1,
                    AGENTS_FLD_LAST_ACTIVE: datetime.utcnow().isoformat(),
                }, "typecast": True},
            )
            print(f"[usage] {telegram_id} -> {current_total + 1} total proposals", flush=True)
    except Exception as e:
        print(f"[usage] tracking failed for {telegram_id}: {e}", flush=True)


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

        base_name = _build_filename(prospect["name"])
        pptx_path = os.path.join(GENERATED_DIR, f"{base_name}.pptx")

        data = {"prospect": prospect, "agent_name": agent_name, "tiers": tiers_data}
        generate_proposal("future_first", data, pptx_path)

        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", GENERATED_DIR, pptx_path],
            check=True, timeout=60,
        )
        pdf_path = pptx_path.rsplit(".", 1)[0] + ".pdf"

        public_base = os.environ.get("PUBLIC_BASE_URL", "https://afg-proposal-service-production.up.railway.app")
        # Upload both files directly (multipart) rather than by URL -- more
        # reliable, and PDF going first means a PPTX hiccup can't block it.
        try:
            _telegram_send_document_by_path(telegram_id, pdf_path)
        except Exception as e:
            print(f"[process] PDF delivery failed for {record_id}: {e}", flush=True)

        try:
            _telegram_send_document_by_path(telegram_id, pptx_path)
        except Exception as e:
            print(f"[process] PPTX delivery failed for {record_id}: {e}", flush=True)

        _track_agent_usage(telegram_id)
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
