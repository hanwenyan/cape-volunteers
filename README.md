# Mrs. White’s Class Volunteer Schedule (2026–2027)

Zero-schema-change mobile web application (`index.html`) backed by a container-bound Google Apps Script JSON API (`Code.gs`) for `Mrs_White_Volunteer_Schedule_2026-27`.

Teachers and Room Parents continue editing the Google Sheet directly on desktop, while parents use the mobile app (in Safari, Chrome, or ParentSquare) to check shifts, claim open slots, request subs, and add shifts to their calendar.

---

## 5-Step Deployment Checklist

### Step 1: Bind `Code.gs` to the Google Sheet
1. Open `Mrs_White_Volunteer_Schedule_2026-27` in Google Sheets.
2. In the top menu, click **Extensions → Apps Script**.
3. Replace the default contents of `Code.gs` with the contents of [`Code.gs`](./Code.gs) from this repository and click **Save** (`Cmd+S`).
4. *(Optional verification)* In the Apps Script toolbar dropdown, select `inspectWorkbookSchema` and click **Run**. Check the **Execution log** to confirm all 9 monthly tabs (`Oct 2026` – `Jun 2027`) and row coordinates resolved cleanly.

### Step 2: Add the Installable `onEdit` Trigger (Desktop Cache Invalidation)
1. In the left sidebar of the Apps Script editor, click **Triggers** (the alarm clock icon).
2. Click **+ Add Trigger** in the bottom-right corner and configure:
   - **Choose which function to run:** `handleSheetEdit`
   - **Select event source:** `From spreadsheet`
   - **Select event type:** `On edit`
3. Click **Save** and approve the Google account permissions prompt. *(This ensures any direct desktop edit by Mrs. White or the Room Parent immediately invalidates that month’s cache.)*

### Step 3: Deploy the Apps Script Web App
1. In the top-right corner of the Apps Script editor, click **Deploy → New deployment**.
2. Click the gear icon next to **Select type** and choose **Web app**.
3. Set:
   - **Description:** `Volunteer Schedule Mobile API v1`
   - **Execute as:** `Me` (the spreadsheet owner/editor)
   - **Who has access:** `Anyone` *(required so parents don't hit a Google login wall inside the ParentSquare WebView)*
4. Click **Deploy** and copy the generated **Web app URL** (`https://script.google.com/macros/s/.../exec`).

### Step 4: Set `DEFAULT_APPS_SCRIPT_URL` in `index.html`
1. Open [`index.html`](./index.html) and find `DEFAULT_APPS_SCRIPT_URL` near the top of the `<script>` block:
   ```js
   const DEFAULT_APPS_SCRIPT_URL = "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec";
   ```
2. Paste your copied `/exec` URL and save. *(Optional: if you ever want to require a 4-digit Class PIN before saving edits, set `REQUIRE_PIN: true` and `CLASS_PIN: "2026"` at the top of `Code.gs`.)*

### Step 5: Publish on GitHub Pages
1. Push this repository (`index.html`, `Code.gs`, `README.md`) to GitHub.
2. In your GitHub repository, open **Settings → Pages**.
3. Under **Build and deployment**, select **Deploy from a branch** → `main` → `/ (root)` and click **Save**.
4. Share the published GitHub Pages link in ParentSquare!
