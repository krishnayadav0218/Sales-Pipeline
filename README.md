# Sales Reporting Pipeline — Render + Google Drive Deployment

This runs entirely in the cloud on Render as a scheduled job. No PC needs to
stay on, no desktop app needs to be installed. It talks to Google Drive
directly through the Drive API using a Service Account.

---

## Part 1: Google Cloud setup (one-time, ~10 minutes)

1. **Create a Google Cloud project**
   Go to https://console.cloud.google.com/ → top-left project dropdown →
   **New Project** → name it (e.g. `sales-pipeline`) → Create.

2. **Enable the Google Drive API**
   In the same project: search bar → "Google Drive API" → click it → **Enable**.

3. **Create a Service Account**
   Left menu → **IAM & Admin** → **Service Accounts** → **Create Service Account**.
   Name it anything (e.g. `sales-pipeline-bot`). Skip the optional role/access
   steps → **Done**.

4. **Generate a JSON key**
   Click the service account you just created → **Keys** tab → **Add Key** →
   **Create new key** → choose **JSON** → it downloads a `.json` file.
   **Keep this file safe — it's the credential the pipeline uses to log in.**

5. **Note the service account's email**
   Open the downloaded JSON — the `client_email` field looks like
   `sales-pipeline-bot@sales-pipeline.iam.gserviceaccount.com`. Copy it.

---

## Part 2: Google Drive folder setup

1. In your normal Google Drive, create two folders:
   - `Sales_Regional_Raw` — where managers upload their Excel files
   - `Sales_Reports_Output` — where the generated report will appear

2. **Share both folders** with the service account email from step 5 above,
   giving it **Editor** access (Share → paste the email → Editor → Send).
   Without this, the pipeline cannot see or write to your Drive — the service
   account is a separate "robot" account, not your personal account.

3. **Get each folder's ID** from its URL when you open it in a browser:
   `https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOpQrSt`
   → the ID is `1AbCdEfGhIjKlMnOpQrSt`. Copy both folder IDs.

---

## Part 3: Push this code to GitHub

Render deploys from a Git repo, so this folder needs to be a GitHub repo first.

```bash
cd render_deploy
git init
git add .
git commit -m "Initial sales pipeline"
# create an empty repo on github.com, then:
git remote add origin https://github.com/<your-username>/sales-pipeline.git
git push -u origin main
```

**Important:** never commit your actual service account JSON key or a real
`.env` file. `.env.example` is just a template — the real secret goes into
Render's dashboard in Part 4, step 3.

---

## Part 4: Deploy on Render

1. Go to https://dashboard.render.com/ → **New** → **Blueprint**.
2. Connect your GitHub account and select the `sales-pipeline` repo.
   Render will detect `render.yaml` automatically and show you the cron job
   it's about to create.
3. Before/after creating, open the service → **Environment** tab → add:
   - `GOOGLE_SERVICE_ACCOUNT_JSON` → paste the **entire content** of the JSON
     key file you downloaded in Part 1 (as one value)
   - `RAW_FOLDER_ID` → the raw folder's ID from Part 2
   - `OUTPUT_FOLDER_ID` → the output folder's ID from Part 2
4. Click **Deploy**. Render will install dependencies and the cron job will
   run on the schedule set in `render.yaml` (default: 8:00 AM IST daily).

---

## Changing the schedule

Edit the `schedule` line in `render.yaml`. It's a standard cron expression in
**UTC**, e.g.:
- `"30 2 * * *"` → 2:30 AM UTC = 8:00 AM IST
- `"0 4 * * *"` → 4:00 AM UTC = 9:30 AM IST

Commit and push the change — Render redeploys automatically.

---

## Testing locally before deploying

```bash
pip install -r requirements.txt
export GOOGLE_SERVICE_ACCOUNT_JSON="$(cat /path/to/your-key.json)"
export RAW_FOLDER_ID="your_raw_folder_id"
export OUTPUT_FOLDER_ID="your_output_folder_id"
python sales_pipeline.py
```

Check the console logs — it will print each step (download, clean, upload)
and any file that got skipped due to a formatting problem.

---

## Files in this package

| File | Purpose |
|---|---|
| `sales_pipeline.py` | Main script: downloads raw files, cleans/aggregates, builds report, uploads it |
| `gdrive_utils.py` | Google Drive API helper functions (auth, list, download, upload) |
| `requirements.txt` | Python dependencies Render will install |
| `render.yaml` | Render Blueprint — defines this as a scheduled Cron Job |
| `.env.example` | Template showing which environment variables are needed |

---

## Notes on formulas in the output report

The report's summary sheet uses live Excel formulas (SUMIFS, SUM), not
hardcoded numbers, so it stays accurate if you ever recompute by hand. When
opened in Excel, Google Sheets, or the Drive preview, these formulas
recalculate automatically — you don't need to do anything extra.
