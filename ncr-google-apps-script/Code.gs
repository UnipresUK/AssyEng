// ─── NCR Tracker — Google Apps Script Backend ──────────────────
const SPREADSHEET_ID = '17WxtqSYbL4wt7bnAmHavbSM8MYE3CnRdkikeZpiHIuA';
const PROGRESS_SHEET  = 'NCR_Progress';
const NCR_DATA_SHEET  = 'All NCRs';
const EXPECTED_HEADERS = ['NCR_Number','Investigator','Status','Notes','Category','Dept',
                          'TS_Investigated','TS_Countermeasure','TS_Failed','TS_Complete','DueDate','LastUpdated'];

// ─── Get sheet, adding any missing header columns automatically ─
function getSheet() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet   = ss.getSheetByName(PROGRESS_SHEET);

  if (!sheet) {
    // Brand new sheet — create with all headers
    sheet = ss.insertSheet(PROGRESS_SHEET);
    sheet.appendRow(EXPECTED_HEADERS);
    sheet.getRange(1, 1, 1, EXPECTED_HEADERS.length).setFontWeight('bold');
    return sheet;
  }

  // Sheet exists — check for any missing columns and append them
  const lastCol      = sheet.getLastColumn() || 0;
  const existingHdrs = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String)
    : [];

  EXPECTED_HEADERS.forEach(h => {
    if (!existingHdrs.includes(h)) {
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(h).setFontWeight('bold');
      existingHdrs.push(h); // keep in sync for subsequent iterations
    }
  });

  return sheet;
}

// ─── JSON helpers ───────────────────────────────────────────────
function jsonResp(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
function ok(data)  { return jsonResp({ ok: true,  data: data  }); }
function fail(msg) { return jsonResp({ ok: false, error: msg  }); }

// ─── CORS-friendly: all calls arrive as GET params ──────────────
function doGet(e)  { return route(e.parameter || {}); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (_) {}
  if (e.parameter) Object.keys(e.parameter).forEach(k => { if (!p[k]) p[k] = e.parameter[k]; });
  return route(p);
}

function route(p) {
  try {
    switch (p.action) {
      case 'getAll':      return doGetAll();
      case 'update':      return doUpdate(p);
      case 'getNCRData':  return doGetNCRData();
      case 'setNCRData':  return doSetNCRData(p);
      default:            return fail('Unknown action: ' + (p.action || 'none'));
    }
  } catch (err) {
    return fail(err.toString());
  }
}

// ─── GET ALL progress records ───────────────────────────────────
function doGetAll() {
  const sheet = getSheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return ok([]);

  const headers = data[0].map(String);
  const records = data.slice(1)
    .filter(row => row[0] !== '')
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? String(row[i]) : ''; });
      return obj;
    });
  return ok(records);
}

// ─── GET NCR list data from All NCRs sheet ──────────────────────
function doGetNCRData() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(NCR_DATA_SHEET);
  if (!sheet) return ok([]);
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return ok([]);

  // Find the header row — scan first 10 rows for one containing 'NCR'
  let headerIdx = -1;
  for (var i = 0; i < Math.min(data.length, 10); i++) {
    if (data[i].some(function(c) { return /ncr/i.test(String(c)); })) {
      headerIdx = i; break;
    }
  }
  if (headerIdx < 0) return ok([]);

  const headers = data[headerIdx].map(String);
  const records = data.slice(headerIdx + 1)
    .filter(function(row) { return row.some(function(c) { return c !== '' && c !== null; }); })
    .map(function(row) {
      const obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i] !== undefined ? String(row[i]) : ''; });
      return obj;
    });
  return ok(records);
}

// ─── WRITE NCR list data to NCR_Data sheet ──────────────────────
function doSetNCRData(p) {
  const records = JSON.parse(p.records || '[]');
  if (!records.length) return fail('No records provided');

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet   = ss.getSheetByName(NCR_DATA_SHEET);
  if (!sheet) sheet = ss.insertSheet(NCR_DATA_SHEET);
  else sheet.clearContents();

  const headers = ['NCR Number','Logged Date','To','Part Number','Fault Code',
                   'Status','Workcentre Found','Sign Off - Supervisor Name'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');

  if (records.length > 0) {
    const rows = records.map(r => headers.map(h => r[h] || ''));
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  return ok({ written: records.length });
}

// ─── UPSERT a progress record ───────────────────────────────────
function doUpdate(p) {
  const ncr = (p.ncr || '').toString().trim();
  if (!ncr) return fail('ncr is required');

  const sheet   = getSheet();
  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(String);

  // Build header → column index map (0-based)
  const colIdx = {};
  headers.forEach((h, i) => { colIdx[h] = i; });

  // Values we want to write keyed by header name
  const values = {
    'NCR_Number':        ncr,
    'Investigator':      p.investigator      || '',
    'Status':            p.status            || '',
    'Notes':             p.notes             || '',
    'Category':          p.category          || '',
    'Dept':              p.dept              || '',
    'TS_Investigated':   p.ts_investigated   || '',
    'TS_Countermeasure': p.ts_countermeasure || '',
    'TS_Failed':         p.ts_failed         || '',
    'TS_Complete':       p.ts_complete       || '',
    'DueDate':           p.due_date          || '',
    'LastUpdated':       new Date().toISOString(),
  };

  // Build a row array aligned to current column positions
  const numCols  = headers.length;
  const rowArray = new Array(numCols).fill('');
  Object.keys(values).forEach(h => {
    if (colIdx[h] !== undefined) rowArray[colIdx[h]] = values[h];
  });

  // Find existing row for this NCR (skip header row 0)
  const ncrCol = colIdx['NCR_Number'];
  let found = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][ncrCol]) === ncr) { found = i + 1; break; } // 1-indexed sheet row
  }

  if (found > 0) {
    sheet.getRange(found, 1, 1, numCols).setValues([rowArray]);
  } else {
    sheet.appendRow(rowArray);
  }

  return ok({ ncr: ncr });
}
