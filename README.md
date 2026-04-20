# Budget Tracker

A Google Apps Script web app that compares approved budgets against actual GL spending, organized by SuiteKey (cost center or project code). Data comes from a NetSuite GL Multicurrency export pasted into Google Sheets.

![Budget Tracker](budget-tracker-icon.png)

---

## Features

- **Dashboard** — overview cards per SuiteKey with budget vs. actuals, utilization bars, and color-coded status (on track / approaching / over budget)
- **Drill-down** — click any SuiteKey to see a category breakdown, then click any category to see individual GL transactions
- **Watchlist** — create rules to track spending by vendor, person, or keyword; scoped to a specific SuiteKey or across all; click any rule to see matching transactions
- **Budget template** — editable budget sheet per SuiteKey per fiscal year, pre-filled from the chart of accounts
- **Period filtering** — full fiscal year, quarter, month, or custom date range
- **Role-based access** — Finance sets a permissions sheet; owners can edit budgets, viewers can read

---

## Setup

### 1. Deploy the script

1. Open [Google Apps Script](https://script.google.com) and create a new project
2. Paste `Code.js` and `Index.html` into the editor (rename `Code.js` to `Code.gs` if needed)
3. Add `appsscript.json` as the manifest (enable **Show "appsscript.json" manifest file** in Project Settings)
4. **Deploy → New deployment → Web app**
   - Execute as: **User accessing the web app**
   - Who has access: **Anyone in your organization**
5. Copy the deployment URL

### 2. Create the permissions sheet (Finance)

Create a Google Sheet and share it as **Viewer with everyone in the org** (required — the script runs as each user, so they must be able to open it).

| Col A — SuiteKey | Col B — Viewer emails (comma-separated) | Col C — Owner email |
|---|---|---|
| WaxhawIT-U_SIL | viewer1@example.org, viewer2@example.org | owner@example.org |

### 3. Create the chart of accounts sheet (Finance, optional)

Paste a NetSuite COA export into a Google Sheet. Expected columns:

| Col A — NetSuite Category | Col B — Account Code | Col C — Account Name | Col D — Description |
|---|---|---|---|

Codes below 4000 and rows with "SUMMARY" in the name are skipped. If no COA sheet is linked, a built-in set of common accounts is used.

### 4. Configure Settings (Finance)

Open the web app, go to **Settings**, and paste the spreadsheet URLs for the permissions sheet and COA sheet under **Admin settings**.

### 5. Link your GL export (each user)

1. Export the **General Ledger Multicurrency** report from NetSuite
2. Paste it into a Google Sheet (one tab per fiscal year recommended)
3. In the web app, go to **Settings → Your GL transactions sheet**, paste the URL, select the tab, and save

### 6. Create budgets (SuiteKey owners)

Go to **Manage**, select a fiscal year, and click **Create budget** next to each SuiteKey you own. The budget sheet is created in your Google Drive, pre-filled with the chart of accounts. Enter amounts in the **Budget template** tab and save.

---

## GL export format

Export **General Ledger Multicurrency** from NetSuite. The app reads these columns (0-based):

| Column | Field |
|---|---|
| A (0) | Account section header |
| B (1) | Account code (detail rows) |
| D (3) | Posted date (DD-MM-YY) |
| F (5) | Document number |
| G (6) | Description |
| H (7) | Name / vendor / donor |
| K (10) | ICJE SuiteKey — must match the permissions sheet exactly |
| L (11) | Debit |
| M (12) | Credit |
| O (14) | Local amount |

> The SuiteKey column (K) must be populated for transactions to appear. Rows without a SuiteKey are ignored.

---

## Architecture

```
ScriptProperties (shared, set by Finance)
  PERMISSIONS_SHEET_ID   — spreadsheet ID of the permissions sheet
  COA_SHEET_ID           — spreadsheet ID of the chart of accounts

UserProperties (per user)
  GL_SHEET_ID            — spreadsheet ID of the user's GL export
  GL_TAB_NAME            — tab name within that spreadsheet
  BUDGET_FILE_ID_{key}_FY{year} — Drive file ID of each budget sheet
  WATCH_RULES            — JSON array of watchlist rules
```

Budget sheets are stored privately in each owner's Google Drive, named `Budget — {SuiteKey} — FY{year}`. Fiscal year starts in **October** (configurable via `FISCAL_YEAR_START_MONTH` in `Code.js`).

---

## Local development with clasp

```bash
npm install -g @google/clasp
clasp login
```

Create `.clasp.json` in the project root:

```json
{"scriptId": "YOUR_SCRIPT_ID", "rootDir": "."}
```

Then push changes:

```bash
clasp push
```

Use **Deploy → Test deployments** in the Apps Script editor to test the latest pushed code without creating a new versioned deployment.

---

## Permissions note

Because the web app runs as `USER_ACCESSING`, every user must have at least **Viewer** access to the permissions sheet in Google Drive. If they cannot open it, the app will show a clear error asking them to request access. This is by design — it ensures each user only sees SuiteKeys they are authorized for.
