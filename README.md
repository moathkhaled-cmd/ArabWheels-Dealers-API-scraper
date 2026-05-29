# ArabWheels Dealer Scraper

Automated scraper for UAE used car dealer APIs. Fetches listings from **Alba Cars**, **GTA Cars**, and **AutoMax**, reconciles against previous data, writes to Google Sheets, and emails a summary report.

Runs automatically via GitHub Actions every **Monday** and on the **1st of each month** at 07:00 GST.

---

## Repo structure

```
arabwheels-scraper/
├── .github/
│   └── workflows/
│       └── dealer_scraper.yml   ← GitHub Actions schedule
├── dealer_scraper_combined.py   ← Main script
├── email_sender_v2.gs           ← Google Apps Script (deploy separately)
├── requirements.txt
└── .gitignore
```

---

## First-time setup

### 1. Add GitHub Secrets

Go to your repo → **Settings → Secrets and variables → Actions → New repository secret**

Add these three secrets:

| Secret name | Value |
|-------------|-------|
| `GOOGLE_SERVICE_ACCOUNT_JSON` | The full contents of your `service_account.json` file (paste the entire JSON) |
| `SPREADSHEET_ID` | Your Google Sheet ID from the URL |
| `APPS_SCRIPT_EMAIL_URL` | Your deployed Apps Script Web App URL |

### 2. Push the code

```bash
git init
git remote add origin https://github.com/YOUR_USERNAME/arabwheels-scraper.git
git add .
git commit -m "Initial commit"
git push -u origin main
```

### 3. Verify the workflow

Go to your repo → **Actions** tab. You should see **Dealer Scraper** listed.
Click **Run workflow** → **Run workflow** to trigger a manual test run.

---

## Schedule

| Trigger | When |
|---------|------|
| Weekly | Every Monday at 07:00 GST (03:00 UTC) |
| Monthly | 1st of every month at 07:00 GST (03:00 UTC) |
| Manual | Anytime via Actions → Run workflow |

> Note: if Monday falls on the 1st, GitHub Actions deduplicates and runs once.

---

## Running locally

```bash
pip install -r requirements.txt

# Place service_account.json next to the script, then:
export SPREADSHEET_ID="your_sheet_id"
export APPS_SCRIPT_EMAIL_URL="your_apps_script_url"
python dealer_scraper_combined.py
```

Or just fill in the fallback strings directly in the `CONFIG` block at the top of `dealer_scraper_combined.py`.
