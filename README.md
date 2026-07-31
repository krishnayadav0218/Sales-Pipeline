# Sales Reporting Pipeline — GitHub Actions + Google Drive (OAuth edition)

Runs for free on GitHub's servers, daily. This version authenticates as
**your own Google account** rather than a Service Account, because Service
Accounts have zero storage quota on personal Gmail and cannot upload files —
that's the `storageQuotaExceeded` error you may have hit.

---

## Part 1: Google Cloud setup

1. **Create a Google Cloud project** (skip if you already made one)
   https://console.cloud.google.com/ → New Project → name it → Create.

2. **Enable the Google Drive API**
   Search bar → "Google Drive API" → **Enable**.

3. **Configure the OAuth consent screen** (one-time)
   Left menu → **APIs & Services** → **OAuth consent screen** → choose
   **External** → fill in an app name (anything) and your email in the two
   required fields → Save through the remaining steps (Scopes, Test users —
   you can skip adding scopes/test users here, defaults are fine).

4. **Create an OAuth Client ID**
   **APIs & Services** → **Credentials** → **Create Credentials** →
   **OAuth client ID** → Application type: **Desktop app** → name it
   anything → **Create**. Click **Download JSON** on the credential you just
   created.

5. **Add yourself as a test user**
   Back in **OAuth consent screen** → **Audience** (or **Test users**
   section) → **Add users** → add your own Gmail address. Without this,
   Google will block the login since the app isn't published/verified.

---

## Part 2: Generate your refresh token (run once, on your own computer)

1. Rename the downloaded JSON from step 4 to `client_secret.json` and put it
   in the same folder as `get_refresh_token.py`.

2. Install the one extra library needed for this local step, then run it:
   ```bash
   pip install google-auth-oauthlib
   python get_refresh_token.py
   ```

3. Your browser opens, asks you to log in and approve access — approve it
   (you may see an "unverified app" warning since this is your own private
   app; click **Advanced → Go to [app name] (unsafe)** to proceed, it's safe
   since you created it).

4. The terminal prints three values:
   ```
   GOOGLE_CLIENT_ID     = ...
   GOOGLE_CLIENT_SECRET = ...
   GOOGLE_REFRESH_TOKEN = ...
   ```
   Copy all three — you'll paste them into GitHub Secrets in Part 4.

---

## Part 3: Google Drive folder setup

Create two folders in your normal Drive:
- `Sales_Regional_Raw` — where files get uploaded for processing
- `Sales_Reports_Output` — where the generated report appears

No sharing step needed this time — since the pipeline now acts as you, it
already has full access to your own folders.

Get each folder's ID from its URL:
`https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOpQrSt` → ID is
`1AbCdEfGhIjKlMnOpQrSt`.

---

## Part 4: Push to GitHub and add secrets

```bash
cd render_deploy
git init
git add .
git commit -m "Sales pipeline with OAuth auth"
git remote add origin https://github.com/<your-username>/sales-pipeline.git
git push -u origin main
```

**Never commit `client_secret.json` or your real tokens.**

On your repo → **Settings** → **Secrets and variables** → **Actions** → add
5 secrets:
- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`
- `GOOGLE_REFRESH_TOKEN`
- `RAW_FOLDER_ID`
- `OUTPUT_FOLDER_ID`

(If you still have the old `GOOGLE_SERVICE_ACCOUNT_JSON` secret from before,
you can delete it — it's no longer used.)

---

## Part 5: Test it

Repo → **Actions** tab → **Sales Pipeline** → **Run workflow** button.
Watch the logs — it should now get past the upload step and finish
successfully. Check `Sales_Reports_Output` in your Drive for
`Master_Report.xlsx`.

---

## Changing the schedule

Edit the `cron:` line in `.github/workflows/sales_pipeline.yml` (UTC time):
- `"30 2 * * *"` → 8:00 AM IST
- `"0 4 * * *"` → 9:30 AM IST

---

## Files in this package

| File | Purpose |
|---|---|
| `sales_pipeline.py` | Main script — auto-detects headers/columns in any xlsx/xls/csv, cleans, aggregates, builds report |
| `gdrive_utils.py` | Google Drive API helper (OAuth auth, list, download, upload) |
| `get_refresh_token.py` | **Run once, locally** — generates your refresh token |
| `requirements.txt` | Dependencies GitHub Actions installs |
| `.github/workflows/sales_pipeline.yml` | Daily schedule definition |
| `.env.example` | Template of required secrets |

---

## Why this approach instead of a Service Account

Service Accounts are great for *reading* shared files, but Google gives them
**0 bytes of storage quota** on personal (non-Workspace) Google accounts —
so any attempt to create/upload a file fails with `storageQuotaExceeded`,
regardless of folder permissions. The only ways around it are: a Google
Workspace **Shared Drive** (needs a paid Workspace plan), or authenticating
as a real user via OAuth — which is what this version does.
