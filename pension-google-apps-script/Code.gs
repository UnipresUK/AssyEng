// ─── Configuration ───────────────────────────────────────────────
const SPREADSHEET_ID = '1bAZ-kfiPKjSfiJu7vsnXyFRfxpW-TbVzfvD1Op108dY';
const STATE_SHEET    = 'DashboardState';
const STATE_KEY      = 'state';

// ─── Helpers ─────────────────────────────────────────────────────
function jsonResp(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
function ok(data)  { return jsonResp({ ok: true,  data: data }); }
function fail(msg) { return jsonResp({ ok: false, error: msg }); }

function getOrCreateSheet(name) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(['key', 'value']);
  }
  return sheet;
}

// ─── Routing ─────────────────────────────────────────────────────
function doGet(e)  { return route(e.parameter); }

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (_) {}
  // Also merge URL params (so ?action=save works alongside POST body)
  if (e.parameter) Object.keys(e.parameter).forEach(k => { if (p[k] === undefined) p[k] = e.parameter[k]; });
  return route(p);
}

function route(p) {
  try {
    switch (p.action) {
      case 'load': return doLoad();
      case 'save': return doSave(p);
      default:     return fail('Unknown action: ' + p.action);
    }
  } catch (err) {
    return fail(err.toString());
  }
}

// ─── Load ─────────────────────────────────────────────────────────
// Returns the full dashboard state JSON, or null if nothing saved yet.
function doLoad() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(STATE_SHEET);
  if (!sheet) return ok(null);

  const data = sheet.getDataRange().getValues();
  // Find the row whose first column matches STATE_KEY
  const row = data.find(r => r[0] === STATE_KEY);
  if (!row || !row[1]) return ok(null);

  try {
    return ok(JSON.parse(row[1]));
  } catch (e) {
    return fail('Corrupt state: ' + e.toString());
  }
}

// ─── Save ─────────────────────────────────────────────────────────
// Expects: { action: 'save', state: '<JSON string of full dashboard state>' }
function doSave(p) {
  if (!p.state) return fail('No state provided');

  // Validate JSON
  try { JSON.parse(p.state); } catch (e) { return fail('Invalid JSON: ' + e.toString()); }

  const sheet = getOrCreateSheet(STATE_SHEET);
  const data  = sheet.getDataRange().getValues();

  // Find existing state row (skip header at index 0)
  let stateRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === STATE_KEY) { stateRowIndex = i + 1; break; } // 1-based sheet row
  }

  if (stateRowIndex === -1) {
    // Append new row
    sheet.appendRow([STATE_KEY, p.state]);
  } else {
    // Update existing value cell
    sheet.getRange(stateRowIndex, 2).setValue(p.state);
  }

  return ok({ saved: true, ts: new Date().toISOString() });
}
