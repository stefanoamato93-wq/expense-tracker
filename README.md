# Expense Tracker (lean)

A stripped-down, faster rebuild of the `google-sheets-form` app. Same iOS-style
roll-picker form and History, but **no query engine and no separate database**.

## What's different (and why it's faster)

| | Old (`google-sheets-form`) | This app |
|---|---|---|
| Query tab | Yes (full-DB scans, `queryMeta` cache) | **Removed** |
| Reads a separate big DB | Yes (`SpreadsheetApp.openById`) | **No** |
| OAuth scope | Broad (`spreadsheets`, needed for the other file) | **Narrow (current spreadsheet only)** |
| Data target | Sheet1 | Sheet1 |
| History source | Sheet1 (tail read) | Sheet1 (tail read) |
| Submit UX | Awaited (felt slow when the deployment was cold) | **Optimistic** (instant), idempotent server write |
| Delete UX | Awaited + full refetch | **Optimistic** (instant), content-verified server delete, restores on failure |

> v3 fixes a sequential-delete bug: rows are now matched by content (not by the
> sheet row number, which shifts as rows are deleted), in-flight deletes are
> tombstoned so a background refresh can't flash them back, and row numbers are
> reconciled by a background refetch after each delete.

Three things drive the speed:

1. **Smaller, single-scope backend.** Dropping `openById` means the script only
   needs the narrow "current spreadsheet" scope. That avoids the broad-scope
   re-authorization that was stalling writes, and a smaller script warms up
   faster on Google's side.
2. **Optimistic submit.** A new entry shows in History and clears the form
   immediately; the network write happens in the background. You never watch a
   spinner. Safe because the backend is **idempotent**: it dedupes on the entry's
   timestamp, so a retried send can never create a duplicate row. Interrupted or
   failed sends are kept locally and shown in History as "not saved · tap to
   retry" (nothing is lost).
3. **Cache-first History + tail read.** History paints instantly from a local
   cache, then refreshes; the backend reads only the last N rows of Sheet1.

Edits and deletes are still awaited (they need the real sheet row and are less
frequent). After a delete, History refetches so the remaining row numbers stay
correct.

## Data model (Sheet1, row 1 = headers, data from row 2)

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| Timestamp | Account (-) | Value | Label | Details | Account (+) | Assetclassdetails |

Timestamp format is `M/D/YYYY H:MM:SS`. Transfers write two rows (the second is
the mirror: swapped accounts, negated value).

## Setup

### 1. Google Sheet
Use a sheet whose first tab is named `Sheet1` with the headers above. You can
point this at the **same Sheet1** the old app already writes to (this app only
touches Sheet1, never the big DB).

### 2. Deploy the backend
1. In the sheet: **Extensions > Apps Script**.
2. Paste the contents of `apps-script.js` from this folder.
3. **Deploy > New deployment > Web app**, Execute as **Me**, Who has access
   **Anyone**. Authorize (you'll only be asked for access to *this* spreadsheet,
   not the broad scope).
4. Copy the **/exec URL**.
5. Verify: open that URL in a browser. You should see
   `{"status":"ok","backendVersion":"lean-v1",...}`.

### 3. Point the frontend at it
Open `index.html`, find:
```js
const SCRIPT_URL = 'PASTE_NEW_DEPLOYMENT_EXEC_URL_HERE';
```
and paste the /exec URL from step 2.

### 4. Host (optional)
Open `index.html` directly, or host it (e.g. GitHub Pages) and Add to Home
Screen on iPhone. This is a new app, so give it its own repo/URL rather than
overwriting the old one.

## Backend actions (`apps-script.js`)
- `submit` — append a row (plus mirror row for transfers). Idempotent: dedupes
  against the last 40 rows on `timestamp | account | value | label`.
- `update` — overwrite the row at `rowIndex`.
- `delete` — delete a row. If the caller passes the row content (account/value/
  label/details), it deletes at `rowIndex` only if that row still matches, else
  relocates the matching row in the tail, else treats it as already gone
  (idempotent). This keeps optimistic delete safe even after row numbers shift.
- `getRecent` — return the last `count` rows (default 500, most-recent-first)
  with absolute `rowIndex`.
- `doGet` — echoes `backendVersion` (`lean-v2`) for a quick "is it live?" check.

## Frontend notes
- Categories (Account/Label/Details/Asset Class) are editable per device and
  stored in `localStorage` (`expenseTrackerLists`), exactly like the old app,
  including label-dependent Details and the Export button.
- History cache + unsynced entries live in `localStorage`
  (`expenseTrackerRecentCache`). Stale in-flight entries from a previous session
  are flipped to "tap to retry" on load, never silently dropped.
- If a background send fails, open History and tap the red "not saved" entry to
  retry. The retry is idempotent server-side.

## Relationship to the old app
The old `google-sheets-form` app still exists and is unchanged. This is a
separate, self-contained app with its own deployment and (recommended) its own
host URL. Nothing here reads or writes the big historical DB.
