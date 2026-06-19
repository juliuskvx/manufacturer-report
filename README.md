# EuroAir Intel — Monthly Manufacturer Intelligence Report

Automated monthly PowerPoint report comparing Boeing and Airbus:
- Deliveries YTD by model
- Orders YTD and total backlog
- Stock prices (BA + AIR.PA) with 30-day trend
- Head-to-head scorecard

**Triggers on the 15th of every month** — after both manufacturers have published their monthly data (they release by the 10th–12th).

## Repo
`github.com/juliuskvx/manufacturer-report` (public)

## Files
```
scripts/
  fetch_data.py    ← Boeing + Airbus XLS + yfinance stock prices
  gen_charts.py    ← 4 matplotlib chart PNGs
  gen_pptx.js      ← 8-slide PPTX via pptxgenjs
  send_email.py    ← Gmail SMTP delivery
.github/workflows/
  main.yml         ← GitHub Actions (backup cron + workflow_dispatch)
package.json
```

## GitHub Secrets needed
| Secret | Value |
|--------|-------|
| `GMAIL_USER` | aviotiongeek@gmail.com |
| `GMAIL_APP_PASSWORD` | Gmail App Password |

*(No aviation API keys needed — data comes from free official manufacturer XLS files and Yahoo Finance)*

## cron-job.org setup
- **URL:** `https://api.github.com/repos/juliuskvx/manufacturer-report/actions/workflows/main.yml/dispatches`
- **Method:** POST
- **Schedule:** 15th of every month, 11:00 Vilnius time
- **Headers:**
  - `Authorization: Bearer <GitHub PAT>`
  - `Content-Type: application/json`
- **Body:** `{"ref": "main"}`

## Data sources
| Data | Source | Reliability |
|------|--------|-------------|
| Boeing deliveries/orders/backlog | boeing.com XLS (free, monthly) | ✅ Official |
| Airbus deliveries/orders/backlog | airbus.com XLS (free, monthly) | ✅ Official |
| BA stock price | Yahoo Finance via yfinance | ✅ Free, daily |
| AIR.PA stock price | Yahoo Finance via yfinance | ✅ Free, daily |

**Note:** If the manufacturer XLS URLs change (they update the date in the filename annually),
update `BOEING_XLS_URL` and `AIRBUS_XLS_URLS` in `scripts/fetch_data.py`.
The script falls back to verified hardcoded YTD data if download fails.

## Manual test run
```bash
pip install yfinance openpyxl xlrd matplotlib numpy
npm install
python scripts/fetch_data.py
python scripts/gen_charts.py
node scripts/gen_pptx.js
```
