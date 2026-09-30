# Ventas Daily Dashboard — Cloud Run deployment (always-on, no laptop needed)

This runs the exact same pipeline (`run_daily.py` + `ventas_slide_generator.py`) on Google Cloud,
triggered daily by Cloud Scheduler. It needs **no laptop** and **no Claude** — it lives entirely in
your GCP project.

## How it works
- **Cloud Run Job** runs the container once per invocation (fetch latest mail CSVs → build deck →
  upload to Drive → email it).
- **Cloud Scheduler** invokes that job every day at 10:00 (Asia/Kolkata).
- **Secret Manager** holds the credentials; nothing sensitive is baked into the image.
- **Auth = service account + domain-wide delegation**: the job impersonates the mailbox
  (`raj.mohan@6dtech.co.in`) to read the report email and send the deck. No user sign-in, no token to
  refresh, nothing expires.

## What only you / a Workspace admin can do
Creating the GCP project, enabling billing, creating the service account key, and authorizing
domain-wide delegation require your authenticated account and Workspace **admin** rights. Those steps
are below; the rest is automated by `deploy_cloudrun.sh`.

---

## Step A — Prerequisites (once)
1. Install the gcloud CLI and run `gcloud auth login`.
2. Have (or create) a GCP project with billing enabled.
3. Edit the variables at the top of `cloudrun/deploy_cloudrun.sh` (PROJECT_ID, REGION, recipients).

## Step B — Service account + domain-wide delegation (admin)
1. The deploy script creates the service account `ventas-job@<project>.iam.gserviceaccount.com`.
   Create a **JSON key** for it:
   ```
   gcloud iam service-accounts keys create sa-key.json \
     --iam-account=ventas-job@<PROJECT_ID>.iam.gserviceaccount.com
   ```
2. In the service account's details, note its **Client ID** (a long number), and enable
   "domain-wide delegation".
3. In **Google Workspace Admin** (admin.google.com) → Security → Access and data control →
   **API controls → Domain-wide delegation → Add new**:
   - Client ID = the service account's Client ID
   - OAuth scopes (comma-separated):
     ```
     https://www.googleapis.com/auth/gmail.readonly,
     https://www.googleapis.com/auth/gmail.send,
     https://www.googleapis.com/auth/drive.file
     ```
   - Authorize.
4. Store the key in Secret Manager (the deploy script references this secret):
   ```
   gcloud secrets create ventas-sa-key --data-file=sa-key.json
   ```
   Then delete the local `sa-key.json`.

> Alternative (no admin needed): skip the service account and instead store your existing
> `token_ventas.json` content as a secret named e.g. `ventas-token`, and in `deploy_cloudrun.sh`
> change `--set-secrets` to `TOKEN_JSON=ventas-token:latest` and drop `GMAIL_DELEGATED_USER`.
> Caveat: keep the OAuth app **Internal/Published** so the refresh token never lapses.

## Step C — Deploy
From the **project root** (folder with `run_daily.py`):
```
bash cloudrun/deploy_cloudrun.sh
```
This enables APIs, builds/pushes the image, creates the Cloud Run Job (wired to the secret and env
vars), and schedules it daily.

## Step D — Test immediately
```
gcloud run jobs execute ventas-daily-dashboard --region <REGION>
```
Check the run logs in the Cloud Run console; you should receive the email with the `.pptx` attached.

---

## Configuration (set at deploy time, no code change)
- `MAIL_TO`, `MAIL_CC`, `MAIL_BCC` — recipients (comma-separated).
- `GMAIL_DELEGATED_USER` — the mailbox to read/send as.
- `VENTAS_WORKDIR=/tmp` — writable scratch dir inside the container.
- Schedule/time-zone — `SCHEDULE` and `TZ` in the deploy script (Cloud Scheduler).

## Changing things later
- **Recipients / schedule:** re-run the relevant `gcloud run jobs update` / `gcloud scheduler jobs
  update` command (or edit in the console). No rebuild needed for env-var changes.
- **Deck design / logic:** edit `ventas_slide_generator.py` or `run_daily.py`, then re-run
  `deploy_cloudrun.sh` to rebuild and redeploy.

## Notes
- The container installs LibreOffice so the inline slide preview still renders in the email.
- Drive upload uses the impersonated user's access; if it ever fails the job still sends the email
  (the failure is logged as a warning).
- Cost at one run/day is effectively within free tier.
