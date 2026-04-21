// ─── NCR Tracker — Google Apps Script Backend ──────────────────
const SPREADSHEET_ID = '17WxtqSYbL4wt7bnAmHavbSM8MYE3CnRdkikeZpiHIuA';
const PROGRESS_SHEET = 'NCR_Progress';

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function jsonResp(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function ok(data)  { return jsonResp({ ok: true,  data: data  }); }
function fail(msg) { return jsonResp({ ok: false, error: msg  }); }

// ─── CORS-friendly: all calls arrive as GET params ──────────────
function doGet(e) {
  return route(e.parameter || {});
}

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (_) {}
  if (e.parameter) Object.keys(e.parameter).forEach(k => { if (!p[k]) p[k] = e.parameter[k]; });
  return route(p);
}

function route(p) {
  try {
    switch (p.action) {
      case 'getAll': return doGetAll();
      case 'update': return doUpdate(p);
      default:       return fail('Unknown action: ' + (p.action || 'none'));
    }
  } catch (err) {
    return fail(err.toString());
  }
}

// ─── GET ALL progress records ───────────────────────────────────
function doGetAll() {
  const sheet = getOrCreateSheet(PROGRESS_SHEET, ['NCR_Number','Investigator','Status','Notes','Category','LastUpdated']);
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return ok([]);
  const headers = data[0];
  const records = data.slice(1)
    .filter(row => row[0] !== '')
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i] !== undefined ? String(row[i]) : '');
      return obj;
    });
  return ok(records);
}

// ─── UPSERT a progress record ───────────────────────────────────
function doUpdate(p) {
  const ncr = (p.ncr || '').toString().trim();
  if (!ncr) return fail('ncr is required');

  const sheet   = getOrCreateSheet(PROGRESS_SHEET, ['NCR_Number','Investigator','Status','Notes','Category','LastUpdated']);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const ncrCol  = headers.indexOf('NCR_Number');

  const rowValues = [
    ncr,
    p.investigator || '',
    p.status       || '',
    p.notes        || '',
    p.category     || '',
    new Date().toISOString()
  ];

  // Find existing row (1-indexed, row 1 = headers)
  let found = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][ncrCol]) === ncr) { found = i + 1; break; }
  }

  if (found > 0) {
    sheet.getRange(found, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return ok({ ncr: ncr });
}
