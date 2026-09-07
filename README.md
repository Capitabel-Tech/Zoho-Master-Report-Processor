# Zoho Master Report Processor

Turns an unfiltered Zoho CRM export into a finished, formula-complete MIS workbook —
no manual Excel editing, no manual filtering in Zoho.

## What it does

Upload two files exported straight from Zoho CRM, with **no filters applied**:

1. **Master Report** — every deal, every stage, unfiltered.
2. **Leads Module report** — every lead, unfiltered.

The app figures out the current fiscal year (April–March) and today's quarter, then
builds a single combined workbook:

- `DEALS For Q<n>` — one sheet per quarter of the current fiscal year that has already
  fully elapsed. Filtered to Stage = `Closed Won` or `Resolved-Completed`, Closing Date
  within that quarter.
- `DEALS For Q<current> - Till Date` — same filter, for the quarter still in progress.
- `Pipeline Deals for Q<current>` — Stage in `Login` / `Sanction` / `Disbursement`,
  across the whole fiscal year.
- `Pipeline Deals for Q<current> - <40%` — Stage in `Document Collection` /
  `Identify Bank` / `Customer Visit / Telecommunication`, across the whole fiscal year.
- `New Leads FY - <year> - <year+1>` — Leads with Status = `New`, within the fiscal year.

Every sheet's formula columns (loan amount differences, percentages, etc.) are added
automatically using the shared templates below.

## How templates work

Each report type (Deals / Pipeline / Leads) has its own template — an `.xlsx` with a
title row, a header row, and one example row. A column in the example row that starts
with `=` is a **formula column**: the same formula is reused for every row, with only
the row number shifted. Any other column is a **passthrough column**: its value is
copied from the raw upload by matching header text.

To add a new column: download the template for that report type from the app, add a
header in row 2 and either a formula or an example value in row 3, save, and upload it
back. No code changes needed — every future report picks it up automatically.

## Project structure

```
backend/    FastAPI app - the processing engine and API
  app/
    config.py               fiscal-year/quarter date logic, stage/status filters
    services/excel_engine.py   template parsing, filtering, workbook generation
    routers/                 API endpoints
  data/templates/            the 3 live templates (deals, pipeline, leads)
  scripts/seed_templates.py  regenerates a template back to its default structure

frontend/   React (Vite) app - upload UI + template management
  src/
    components/MasterReportPanel.jsx   the main upload flow
    components/TemplatePanel.jsx        per-report-type template download/upload
```

## Running locally

**Backend:**
```bash
cd backend
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```

**Frontend:**
```bash
cd frontend
npm install
npm run dev
```

Then open http://localhost:5173.

## Deployment

- **Backend**: deployed on Render as a manually-configured Web Service (Root Directory
  `backend`, Build Command `pip install -r requirements.txt`, Start Command
  `uvicorn app.main:app --host 0.0.0.0 --port $PORT`). No environment variables required.
- **Frontend**: deployed on Netlify (Base directory `frontend`, Build command
  `npm run build`, Publish directory `dist`), with one environment variable:
  `VITE_API_BASE` set to the deployed backend's URL plus `/api`
  (e.g. `https://your-backend.onrender.com/api`).

Both auto-deploy on push to `main`.

## Known limitations

- Render's free tier has an ephemeral filesystem — any template updated through the
  app's "Upload the edited template" resets back to what's committed in this repo on
  the next redeploy/restart. Edit the template files under `backend/data/templates/`
  and commit them if you want a change to persist permanently.
- No authentication — anyone with the deployed URL can use the app.
