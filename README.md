# AFG Proposal Generation Service

Turns extracted quote data into a filled PPTX + PDF, using the exact pixel-mapped
template logic validated manually for FWD Future First. Called by Make.com after
the Claude API extraction step.

## Project structure
```
service/
├── main.py              # FastAPI app: POST /generate, GET /templates, GET /files/{name}
├── overlay_engine.py     # Generic engine -- reads coords.json, fills any master.pptx
├── requirements.txt
├── Dockerfile            # Python + LibreOffice (needed for PDF conversion)
├── templates/
│   └── future_first/
│       ├── master.pptx   # Your real AFG master template, untouched
│       └── coords.json   # Every pixel-mapped coordinate, hand-tuned + validated
└── generated/            # Output files land here (pptx + pdf per request)
```

## Deploy to Railway
1. Push this `service/` folder to a GitHub repo (or connect Railway directly to a repo).
2. In Railway: New Project → Deploy from GitHub repo → select it.
3. Railway auto-detects the `Dockerfile` and builds it (LibreOffice install takes
   a few minutes on first deploy, then it's cached).
4. Railway assigns a public URL, e.g. `https://afg-proposals.up.railway.app`.
5. Test: `GET https://<your-url>/health` should return `{"status":"ok","templates":["future_first"]}`.

No environment variables are required for this service itself (it doesn't call
Claude or Telegram directly -- Make.com orchestrates that and just calls this
service's `/generate` endpoint with already-extracted JSON).

## API

### `POST /generate`
Request body matches the extraction-prompt.md schema, plus `template_id` and `agent_name`:
```json
{
  "template_id": "future_first",
  "agent_name": "Sutharthan Marimuthu",
  "prospect": { "name": "John Doe", "dob": "01/01/1997", "age": "30", "smoking_status": "Non-Smoker" },
  "tiers": [
    {
      "tier_label": "A",
      "monthly_premium": "290.41",
      "death": "250,000", "tpd": "250,000", "terminal_illness": "250,000",
      "ci_minus": "250,000", "ci_plus": "-", "simplified_ci": "-",
      "waiver_policy_owner": "-", "waiver_life_assured": "-",
      "medical_card": "-", "room_and_board": "-", "annual_limit": "-",
      "lifetime_limit": "-", "co_insurance_deductible": "-",
      "personal_accident": "-", "coverage_up_to_age": "80",
      "remarks": "Future First + CI Lite Rider"
    }
  ]
}
```
Response:
```json
{ "job_id": "a1b2c3d4e5", "pptx_url": "/files/John Doe - a1b2c3d4e5.pptx", "pdf_url": "/files/John Doe - a1b2c3d4e5.pdf" }
```
Make.com then does an HTTP GET on those `/files/...` URLs (prefixed with your Railway
domain) to grab the actual bytes and forward them to Telegram.

### `GET /templates`
Returns which product templates are currently available -- so the bot only offers
products it can actually generate for.

## How white-labeling works

Nothing about a new product or a new organization's branding requires touching
`main.py` or `overlay_engine.py`. It's purely a data addition:

1. Create `templates/<new_id>/master.pptx` -- the new brand/product's real template file.
2. Create `templates/<new_id>/coords.json` -- pixel-map its title slide (name/prepared-by
   position) and quote slide (info boxes, column x-ranges, row y-ranges) the same way we
   did for Future First. This is manual, one-time work per template (the process: crop
   the slide background images, scan for grid-line pixel boundaries, convert px → inches
   at the master's canvas scale -- same steps used to build coords.json for Future First).
3. Deploy (or just push -- Railway auto-redeploys on git push). The new `template_id`
   immediately shows up in `GET /templates` and is generatable via `POST /generate`.

This is exactly the same pattern for Life First and CI First (still pending from you) --
each is a `templates/life_first/` and `templates/ci_first/` folder, nothing more.

## Font note
Calibri is a Microsoft font, not available on Linux by default. The Dockerfile installs
`fonts-crosextra-carlito`, a metrically-identical open-source substitute -- LibreOffice
auto-substitutes it during PDF conversion, so the layout stays pixel-accurate even though
the container doesn't have real Calibri installed.

## Local testing (optional, before deploying)
```bash
pip install -r requirements.txt
uvicorn main:app --reload
# then POST to http://localhost:8000/generate with the JSON body above
```
Requires LibreOffice installed locally for the PDF step to work (`soffice --headless`).
