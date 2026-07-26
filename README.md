# Sales Reporting Pipeline — GitHub Actions + Google Drive Deployment

Runs entirely for free on GitHub's own servers, on a daily schedule. **No card,
no server, no PC needs to stay on.** GitHub Actions gives every account free
scheduled-workflow minutes — plenty for a job that runs once a day for a
minute or two.

---

## Part 1: Google Cloud setup (one-time, ~10 minutes)

1. **Create a Google Cloud project**
   https://console.cloud.google.com/ → top-left project dropdown → **New
   Project** → name it (e.g. `sales-pipeline`) → Create.

2. **Enable the Google Drive API**
   Same project → search bar → "Google Drive API" → click it → **Enable**.

3. **Create a Service Account**
   Left menu (☰) → **IAM & Admin** → **Service Accounts** → **Create Service
   Account**. Name it anything (e.g. `sales-pipeline-bot`). Skip the optional
   role/access steps → **Done**.

4. **Generate a JSON key**
   Click the service account's name/email (the blue link, not the checkbox)
   → **Keys** tab at the top → **Add Key** → **Create new key** → **JSON** →
   **Create**. A `.json` file downloads — keep it safe.

   *If the Keys tab is missing or "Add Key" is greyed out:* your Google
   account may have an Organization Policy blocking key creation (common on
   work/company Google accounts). Create the Cloud project under a personal
   Gmail account instead — this almost always fixes it.

5. **Note the service account's email**
   Open the downloaded JSON — the `client_email` field looks like
   `sales-pipeline-bot@sales-pipeline.iam.gserviceaccount.com`. Copy it.

---

## Part 2: Google Drive folder setup

1. In your normal Google Drive, create two folders:
   - `Sales_Regional_Raw` — where managers upload their Excel files
   - `Sales_Reports_Output` — where the generated report will appear

2. **Share both folders** with the service account email from Part 1,
   giving it **Editor** access (right-click folder → Share → paste the
   email → Editor → Send).

3. **Get each folder's ID** from its URL when opened in a browser:
   `https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOpQrSt`
   → the ID is `1AbCdEfGhIjKlMnOpQrSt`. Copy both.

---

## Part 3: Push this code to GitHub

```bash
cd render_deploy
git init
git add .
git commit -m "Initial sales pipeline"
# create an empty repo on github.com first, then:
git remote add origin https://github.com/<your-username>/sales-pipeline.git
git push -u origin main
```

**Never commit your real service account JSON or a real `.env` file** —
`.env.example` is just a template. The actual secret goes into GitHub's
encrypted Secrets in Part 4.

---

## Part 4: Add secrets in GitHub (no card, ever)

1. On your repo's GitHub page → **Settings** tab → left sidebar → **Secrets
   and variables** → **Actions**.
2. Click **New repository secret** three times, adding:
   - `GOOGLE_SERVICE_ACCOUNT_JSON` → paste the **entire content** of the
     downloaded JSON key file as the value
   - `RAW_FOLDER_ID` → the raw folder's ID from Part 2
   - `OUTPUT_FOLDER_ID` → the output folder's ID from Part 2

These are encrypted and never visible in logs or to anyone browsing the repo.

---

## Part 5: That's it — it's already scheduled

The workflow file `.github/workflows/sales_pipeline.yml` is already in this
repo. As soon as you push it (Part 3) and add the secrets (Part 4), GitHub
will automatically run it daily at 8:00 AM IST (2:30 AM UTC).

### To test it right now, without waiting for the schedule:
Go to your repo → **Actions** tab → click **Sales Pipeline** in the left
list → **Run workflow** button (top right) → **Run workflow**. Watch it
run live and check the logs for each step (download, clean, upload).

### To change the schedule:
Edit the `cron:` line in `.github/workflows/sales_pipeline.yml`. It's in
**UTC**, same format as before:
- `"30 2 * * *"` → 8:00 AM IST
- `"0 4 * * *"` → 9:30 AM IST

Commit and push — GitHub picks up the new schedule automatically.

---

## A note on GitHub Actions' free limits

- **Public repositories:** unlimited free minutes for scheduled workflows.
- **Private repositories:** 2,000 free minutes/month on a free GitHub
  account — this pipeline takes well under a minute per run, so a daily run
  uses a tiny fraction of that. You will not hit the limit or be asked for a
  card at any point.

---

## Testing locally before pushing (optional)

```bash
pip install -r requirements.txt
export GOOGLE_SERVICE_ACCOUNT_JSON="$(cat /path/to/your-key.json)"
export RAW_FOLDER_ID="your_raw_folder_id"
export OUTPUT_FOLDER_ID="your_output_folder_id"
python sales_pipeline.py
```

---

## Files in this package

| File | Purpose |
|---|---|
| `sales_pipeline.py` | Main script: downloads raw files, cleans/aggregates, builds report, uploads it |
| `gdrive_utils.py` | Google Drive API helper functions (auth, list, download, upload) |
| `requirements.txt` | Python dependencies |
| `.github/workflows/sales_pipeline.yml` | GitHub Actions schedule — runs the pipeline daily, for free |
| `.env.example` | Template showing which secrets/env vars are needed |

---

## Notes on formulas in the output report

The report's summary sheet uses live Excel formulas (SUMIFS, SUM), not
hardcoded numbers. When opened in Excel, Google Sheets, or the Drive
preview, these formulas recalculate automatically.
