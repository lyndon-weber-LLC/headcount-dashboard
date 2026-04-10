# Google Sheets Integration Setup

Follow these steps once your admin has enabled Google Cloud Platform for your account.

---

## Admin Step (one-time, done by your Google Workspace admin)
1. Go to **admin.google.com**
2. Click **Apps** → **Additional Google services**
3. Search for **"Google Cloud Platform"** → click it
4. Click **Turn ON for everyone** (or just for your account) → **Save**

---

## Your Steps

### Step 1 — Create a Google Cloud Project (~2 min)
1. Go to **console.cloud.google.com**
2. Click the project dropdown at the top → **New Project**
3. Name it `Headcount Dashboard` → click **Create**

### Step 2 — Enable the Google Sheets API (~1 min)
1. In the search bar at the top, type **"Google Sheets API"** → click it
2. Click **Enable**

### Step 3 — Create a Service Account (~3 min)
1. In the left sidebar go to **IAM & Admin** → **Service Accounts**
2. Click **+ Create Service Account**
3. Name it `headcount-dashboard` → click **Create and Continue**
4. On the "Grant this service account access" step, click the **Role** dropdown → select **Viewer** under Basic roles (read-only — it cannot modify anything)
5. Click **Continue** → **Done**
6. Click on the service account you just created
7. Go to the **Keys** tab → **Add Key** → **Create new key** → choose **JSON** → click **Create**
8. A JSON file will download automatically — that's your credentials file

> **Security note:** The Viewer role means this account can only read from your Cloud project. It can only access the specific Google Sheets you explicitly share with it — nothing else in your Drive or Google account is visible to it.

### Step 4 — Activate the integration
1. Move the downloaded JSON file into your `headcount-dashboard` folder
2. Rename it to exactly: `google_credentials.json`
3. Open each of your 12 Google Sheets and share them with the service account email
   - The email looks like: `headcount-dashboard@your-project.iam.gserviceaccount.com`
   - Add it as a **Viewer** (same as sharing with any person)

### Step 5 — Tell Claude
Once the above is done, let Claude know and the script will be updated with all 12 sheet IDs and tested automatically. No other changes needed on your end.

---

## Notes
- The `google_credentials.json` file activates the Google Sheets pipeline automatically on every dashboard run
- CSV files in the folder will continue to work as a fallback if Sheets are unreachable
- Subcontractor sheets with different pay cutoffs will work fine — data is pulled by date, not by pay period
