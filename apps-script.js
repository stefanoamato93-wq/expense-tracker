/**
 * Expense Tracker (lean) — Google Apps Script Web App backend.
 *
 * Design goals vs the old google-sheets-form backend:
 *  - Speed: no separate "big DB", no query/queryMeta engine, no getSuggestions.
 *    Fewer functions and NO SpreadsheetApp.openById means this script only needs
 *    the narrow "current spreadsheet" OAuth scope, so it warms up faster and does
 *    not trigger the broad-scope re-authorization that made writes stall.
 *  - Data fills Sheet1 and History reads Sheet1 (getRecent tail-read).
 *  - Safe optimistic UI: submit is idempotent (a retried submit with the same
 *    timestamp will NOT create a duplicate row), so the frontend can send in the
 *    background and retry failures without ever double-writing.
 *
 * SETUP:
 * 1. Google Sheet with headers in row 1 of tab "Sheet1":
 *    Timestamp | Account (-) | Value | Label | Details | Account (+) | Assetclassdetails
 * 2. Extensions > Apps Script, paste this file.
 * 3. Deploy > New deployment > Web app, Execute as: Me, Who has access: Anyone.
 * 4. Copy the /exec URL into index.html (SCRIPT_URL).
 * Visiting the /exec URL in a browser (GET) echoes BACKEND_VERSION so you can
 * confirm the deployed version matches this file after redeploying.
 */

const SHEET_NAME = 'Sheet1';
const BACKEND_VERSION = 'lean-v2';
const N_COLS = 7; // A..G

function doGet() {
  return jsonOut_({
    status: 'ok',
    backendVersion: BACKEND_VERSION,
    message: 'Use POST requests from the form app.'
  });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');

    switch (data.action) {
      case 'submit':   return jsonOut_(handleSubmit_(sheet, data));
      case 'update':   return jsonOut_(handleUpdate_(sheet, data));
      case 'delete':   return jsonOut_(handleDelete_(sheet, data));
      case 'getRecent':return jsonOut_(handleGetRecent_(sheet, data));
      default:
        return jsonOut_({ success: false, error: 'Unknown action', backendVersion: BACKEND_VERSION });
    }
  } catch (err) {
    return jsonOut_({ success: false, error: err.toString() });
  }
}

// ------------------------------------------------------------
// Actions
// ------------------------------------------------------------

function handleSubmit_(sheet, data) {
  const row = rowFromData_(data);

  // Idempotency: if an identical row (same timestamp/account/value/label) is
  // already among the last rows, this is a retry of a submit that actually
  // succeeded, return that row instead of writing a duplicate.
  const dup = findRecentDuplicate_(sheet, data.timestamp, data.accountMinus, data.value, data.label);
  if (dup) return { success: true, rowIndex: dup, deduped: true };

  const lastRow = sheet.getLastRow();
  const target = lastRow + 1;
  sheet.getRange(target, 1, 1, N_COLS).setValues([row]);

  // Transfer: also write the mirror row (swapped accounts, negated value).
  if (String(data.label).toLowerCase() === 'transfer' && data.accountPlus) {
    const mirror = [
      data.timestamp,
      data.accountPlus,
      -Number(data.value),
      data.label,
      data.details,
      data.accountMinus,
      data.assetclassdetails || ''
    ];
    sheet.getRange(target + 1, 1, 1, N_COLS).setValues([mirror]);
  }
  return { success: true, rowIndex: target };
}

function handleUpdate_(sheet, data) {
  const rowIndex = Number(data.rowIndex);
  if (!rowIndex || rowIndex < 2) throw new Error('Invalid rowIndex for update.');
  sheet.getRange(rowIndex, 1, 1, N_COLS).setValues([rowFromData_(data)]);
  return { success: true, rowIndex: rowIndex };
}

function handleDelete_(sheet, data) {
  const rowIndex = Number(data.rowIndex);

  // Content-aware delete (safe for optimistic UI). Deleting a row renumbers the
  // rows below it, so a cached rowIndex can go stale. If the caller passes the
  // row's content we (1) delete at rowIndex only if it still matches, else (2)
  // relocate the matching row in the tail, else (3) treat it as already gone
  // (idempotent). This guarantees we never delete the wrong row.
  const hasContent = data.label !== undefined && data.value !== undefined;
  if (hasContent) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, alreadyGone: true };

    // Match on account/value/label/details using DISPLAY values, which is what
    // getRecent returned to the client, so strings compare like-for-like.
    const eq = (dv) =>
      String(dv[1]) === String(data.accountMinus) &&
      Number(dv[2]) === Number(data.value) &&
      String(dv[3]) === String(data.label) &&
      String(dv[4]) === String(data.details);

    if (rowIndex >= 2 && rowIndex <= lastRow) {
      const cur = sheet.getRange(rowIndex, 1, 1, 5).getDisplayValues()[0];
      if (eq(cur)) { sheet.deleteRow(rowIndex); return { success: true, deletedRow: rowIndex }; }
    }

    // Relocate: scan the tail and pick the match closest to the requested row.
    const n = Math.min(120, lastRow - 1);
    const start = lastRow - n + 1;
    const dv = sheet.getRange(start, 1, n, 5).getDisplayValues();
    let best = 0, bestDist = Infinity;
    for (let i = 0; i < n; i++) {
      if (eq(dv[i])) {
        const abs = start + i;
        const dist = Math.abs(abs - rowIndex);
        if (dist < bestDist) { bestDist = dist; best = abs; }
      }
    }
    if (best) { sheet.deleteRow(best); return { success: true, deletedRow: best, relocated: true }; }
    return { success: true, alreadyGone: true };
  }

  // Legacy path: delete strictly by rowIndex.
  if (!rowIndex || rowIndex < 2) throw new Error('Invalid rowIndex for delete.');
  sheet.deleteRow(rowIndex);
  return { success: true, deletedRow: rowIndex };
}

function handleGetRecent_(sheet, data) {
  // Read only the last `count` data rows (default 500) instead of the whole
  // sheet, so this stays fast no matter how large Sheet1 grows.
  const lastRow = sheet.getLastRow();
  const total = lastRow - 1; // row 1 is the header
  if (total < 1) return { success: true, rows: [], total: 0 };

  const count = Math.max(1, Math.min(Number(data.count) || 500, total));
  const startRow = lastRow - count + 1;
  const rows = sheet.getRange(startRow, 1, count, N_COLS).getDisplayValues();
  const result = rows.map((r, i) => ({
    rowIndex: startRow + i, // absolute sheet row (needed for edit/delete)
    timestamp: r[0],
    accountMinus: r[1],
    value: r[2],
    label: r[3],
    details: r[4],
    accountPlus: r[5],
    assetclassdetails: r[6]
  }));
  return { success: true, count: count, total: total, rows: result.reverse() };
}

// ------------------------------------------------------------
// Helpers
// ------------------------------------------------------------

function rowFromData_(data) {
  return [
    data.timestamp,
    data.accountMinus,
    data.value,
    data.label,
    data.details,
    data.accountPlus || '',
    data.assetclassdetails || ''
  ];
}

// Look at the last N rows for an identical entry to make submit idempotent.
function findRecentDuplicate_(sheet, ts, account, value, label) {
  const lastRow = sheet.getLastRow();
  const n = Math.min(40, lastRow - 1);
  if (n < 1) return 0;
  const vals = sheet.getRange(lastRow - n + 1, 1, n, 4).getValues(); // A..D
  for (let i = n - 1; i >= 0; i--) {
    if (formatTs_(vals[i][0]) === String(ts) &&
        String(vals[i][1]) === String(account) &&
        Number(vals[i][2]) === Number(value) &&
        String(vals[i][3]) === String(label)) {
      return lastRow - n + 1 + i;
    }
  }
  return 0;
}

// Normalise a timestamp cell (Date or text) to the client's
// "M/D/YYYY H:MM:SS" format so duplicate detection compares like-for-like.
function formatTs_(v) {
  const d = (v instanceof Date && !isNaN(v.getTime())) ? v : parseMDY_(v);
  if (!d) return String(v);
  const pad = (x) => ('0' + x).slice(-2);
  return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear() +
         ' ' + d.getHours() + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
}

function parseMDY_(v) {
  if (v === null || v === undefined || v === '') return null;
  const m = String(v).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{1,2}):(\d{1,2}))?/);
  if (!m) { const d = new Date(v); return isNaN(d.getTime()) ? null : d; }
  return new Date(+m[3], +m[1] - 1, +m[2], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
