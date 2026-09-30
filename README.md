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

> **v4 (frontend) + lean-v3 (backend): faster History, no more stuck "not saved".**
> - **Outbox with automatic retry.** Unconfirmed entries stay in a local outbox
>   and are resent automatically (backoff 3s, 10s, 30s, 1m, then every 2m), plus
>   on app open, when the app returns to the foreground and when the phone comes
>   back online. v3 flipped any in-flight entry to "not saved" on every reload
>   and only retried on a manual tap.
> - **Tolerant matching.** v3 matched a local entry to the sheet by comparing raw
>   strings, so if the sheet displayed the timestamp differently (D/M locale, no
>   seconds) or the value with separators, an entry that had actually been saved
>   (response lost) stayed "not saved" forever and showed twice. v4 compares
>   normalised date/time parts (M/D or D/M) and numbers, and lean-v3 also returns
>   a normalised `ts` per row.
> - **One round trip.** New `sync` action: flushes the whole outbox and returns the
>   History tail in a single Apps Script call. Requests are single-flight, so
>   overlapping refreshes can't overwrite each other.
> - **Timeouts.** Every call has a hard timeout (20s reads, 35s writes), so a hung
>   request becomes a retry instead of a permanent "saving…". Non-JSON (HTML
>   error page) responses count as retryable failures.
> - **Faster render.** In-memory cache (no re-parsing localStorage), one HTML
>   string with one delegated click handler, only the first 100 rows rendered
>   ("Show more" adds 100), DOM write skipped when nothing changed, and History is
>   pre-fetched on app open so it is usually fresh when tapped.
> - **Backend write lock.** Every write runs under `LockService`; before, two
>   submits close together could read the same last row and overwrite each other.
>   Dedupe window widened from 40 to 300 rows so late retries still dedupe.
> - History header shows a status line ("Updated 14:05", "Refreshing…",
>   "2 not saved · tap to retry"). Tap it to retry or refresh. An unsaved entry
>   can be discarded with ✕.
> - The v4 frontend also works against the old lean-v2 backend (it detects the
>   missing `sync` action and falls back to submit + getRecent), but only lean-v3
>   gives the single-round-trip speed and the write lock. **Redeploy the backend**
>   (see "Updating the backend").

Three things drive the speed:

1. **Smaller, single-scope backend.** Dropping `openById` means the script only
   needs the narrow "current spreadsheet" scope. That avoids the broad-scope
   re-authorization that was stalling writes, and a smaller script warms up
   faster on Google's side.
2. **Optimistic submit.** A new entry shows in History and clears the form
   immediately; the network write happens in the background. You never watch a
   spinner. Safe because the backend is **idempotent**: it dedupes on the entry's
   timestamp, so a retried send can never create a duplicate row. Interrupted or
   failed sends stay in a local outbox and are retried automatically (nothing is
   lost).
3. **Cache-first History + tail read.** History paints instantly from a local
   cache, then refreshes; the backend reads only the last N rows of Sheet1.

Edits are still awaited (they need the real sheet row). Deletes are optimistic.
After either, History refetches so the remaining row numbers stay correct.

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

### Updating the backend (keeps the same URL)
1. Open the sheet, **Extensions > Apps Script**, replace the code with
   `apps-script.js` and save.
2. **Deploy > Manage deployments**, pencil icon on the existing web app,
   **Version: New version**, **Deploy**. Do not create a new deployment, or the
   /exec URL changes.
3. Open the /exec URL: it should echo the new `backendVersion` (currently
   `lean-v3`).

### 4. Host (optional)
Open `index.html` directly, or host it (e.g. GitHub Pages) and Add to Home
Screen on iPhone. This is a new app, so give it its own repo/URL rather than
overwriting the old one.

## Backend actions (`apps-script.js`)
All writes (`sync` with entries, `submit`, `update`, `delete`) run under a
script lock.
- `sync` — `{entries:[{id, ...submit payload}], count}`. Writes every entry
  (idempotent, same dedupe as `submit`), then returns
  `{results:[{id, success, rowIndex}], rows, total}`. This is what the frontend
  uses for both saving and refreshing History.
- `submit` — append a row (plus mirror row for transfers). Idempotent: dedupes
  against the last 300 rows on `timestamp | account | value | label` (also
  matching a day/month-swapped reading, in case the sheet locale auto-parsed the
  text timestamp as D/M).
- `update` — overwrite the row at `rowIndex`.
- `delete` — delete a row. If the caller passes the row content (account/value/
  label/details), it deletes at `rowIndex` only if that row still matches, else
  relocates the matching row in the tail, else treats it as already gone
  (idempotent). This keeps optimistic delete safe even after row numbers shift.
- `getRecent` — return the last `count` rows (default 500, most-recent-first)
  with absolute `rowIndex`, the display `timestamp` and a normalised `ts`
  (`M/D/YYYY H:MM:SS`).
- `doGet` — echoes `backendVersion` (`lean-v3`) for a quick "is it live?" check.

## Frontend notes
- Categories (Account/Label/Details/Asset Class) are editable per device and
  stored in `localStorage` (`expenseTrackerLists`), exactly like the old app,
  including label-dependent Details and the Export button.
- History cache + outbox live in `localStorage` (`expenseTrackerRecentCache`,
  also holds `syncedAt`) and in memory. An entry is unsynced while it carries a
  `_payload`; `_sending` / `_error` / `_attempts` only drive the status text.
  v3 caches are migrated on load (old `_pending` flags dropped, entries kept in
  the outbox and resent).
- Sync engine (`requestSync` / `runSync` / `doSync` in `index.html`): single
  flight, whole outbox plus History tail per call, backoff retry. Matching of
  local vs server rows is `sameEntry()` (normalised content + date/time parts).
- If a send keeps failing, History shows "not saved yet · retrying". Tap the entry
  or the status line to retry immediately, or ✕ to discard it.
- Edits are still awaited and followed by a background refresh. Deletes are
  optimistic and followed by a refresh that reconciles row numbers.

## Relationship to the old app
The old `google-sheets-form` app still exists and is unchanged. This is a
separate, self-contained app with its own deployment and (recommended) its own
host URL. Nothing here reads or writes the big historical DB.
