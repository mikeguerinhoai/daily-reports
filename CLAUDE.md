# Daily Reports

Cross-channel daily reporting (Voice + SMS + Webchat). Outputs Excel workbook (4 tabs) + branded HTML/PDF (6 landscape pages) + interactive dashboard.

## Pipeline

`fetch-daily-data.py --date YYYY-MM-DD` queries Supabase → `data/daily-report-YYYY-MM-DD.json` → `generate-daily-report.py` writes Excel + HTML/PDF → `generate-daily-dashboard.js` writes interactive dashboard HTML.

## Commands

```bash
# Full refresh (fetch yesterday + generate all outputs):
npm run daily
# Specific date:
npm run daily:date 2026-04-04
# Dashboard only (from existing JSON):
npm run daily-dashboard
# Or step by step:
python daily-reports/fetch-daily-data.py --date 2026-04-04
python daily-reports/generate-daily-report.py
node daily-reports/generate-daily-dashboard.js
# Send email:
python daily-reports/send-daily-email.py --date 2026-04-04
# Open outputs:
start daily-reports/output/HOAi_Daily_Report_2026-04-04.xlsx
start daily-reports/output/HOAi_Daily_Report_2026-04-04.html
start daily-reports/output/HOAi_Daily_Dashboard_2026-04-04.html
```

## Key Files

| File | Role |
|------|------|
| `fetch-daily-data.py` | CLI with `--date`, `--company` flags. Fetches 3 channels, computes 7-day trailing, DOW comparison, MTD pace, revenue intelligence, repeat callers, alerts. Output: ~40 KB JSON |
| `generate-daily-report.py` | Reads JSON, generates Excel + HTML/PDF + per-company outreach pages. Falls back to HTML if weasyprint unavailable |
| `generate-daily-dashboard.js` | Reads daily JSON, injects into `dashboard-template.html`, outputs interactive React dashboard HTML |
| `dashboard-template.html` | React 18 + Chart.js 4 template for interactive daily dashboard (includes filter bar, sort, expand rows) |
| `templates/daily-report-template.html` | Jinja2 template for email/PDF report (6 landscape pages) |
| `templates/outreach-template.html` | Jinja2 template for per-company outreach briefs (usage-adaptive recommendations) |
| `send-daily-email.py` | Sends report via Office 365 SMTP. Requires `SMTP_USERNAME`, `SMTP_PASSWORD` env vars |
| `daily-report-config.json` | Consolidated benchmarks (voice/sms/webchat), revenue intelligence thresholds, COGS rates, per-customer packages, brand tokens |
| `setup-scheduler.ps1` | Windows Task Scheduler: daily at 7 AM ET |

## Config

`daily-report-config.json` — Consolidated benchmarks (voice/sms/webchat), revenue intelligence thresholds (75%/100%/125%), alert thresholds, COGS rates for all 3 channels, per-customer business hours and package details, brand tokens.

## Output Folder Structure

```
output/
  YYYY-MM/                          # Month grouping
    YYYY-MM-DD/                     # Date folder
      Dashboard/                    # Main deliverables
        HOAi_Daily_Dashboard_*.html # Interactive React dashboard
        HOAi_Daily_Report_*.html    # Static branded report
        HOAi_Daily_Report_*.xlsx    # Excel workbook
      Outreach/                     # Per-company outreach pages
        Company_Name/
          outreach.html             # Branded outreach brief
  HOAi_Daily_Dashboard_*.html       # Flat backward-compat copies
  HOAi_Daily_Report_*.html
  HOAi_Daily_Report_*.xlsx
```

Outreach pages are generated for companies with alerts, bad usage status, repeat callers, declining volume, or under-use. Recommendations vary by usage tag (Overage/Proj. Overage/On Track/Under-Use/No Contract) and historic volume trend (Accelerating/Stable/Declining).

## Data Paths

| Path | Contents |
|------|----------|
| `data/daily-report-YYYY-MM-DD.json` | Generated daily JSON payloads (~40 KB each, gitignored) |
| `output/` | Generated reports in nested folder structure (gitignored) |

## Dependencies

- `supabase/queries/call_logs.py` — Voice + SMS call data
- `supabase/queries/sms.py` — Broadcast data
- `supabase/queries/webchat.py` — Homeowner chat data (`homeowner_chat_view`)
