// ─── Configuration ───────────────────────────────────────────────
const SPREADSHEET_ID = '1ujAk9ZYupXSx9q0YWuglgc9GvAZ6xqkTx4UF9oQXLqI';

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet not found: ' + name);
  return sheet;
}

// ─── Helpers ─────────────────────────────────────────────────────
function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).filter(row => row[0] !== '').map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function uid() {
  return 'id' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function simpleHash(str) {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charCodeAt(i);
    hash = ((hash << 5) - hash) + ch;
    hash |= 0;
  }
  return 'h' + Math.abs(hash).toString(16);
}

function jsonResp(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function ok(data)    { return jsonResp({ ok: true, data: data }); }
function fail(msg)   { return jsonResp({ ok: false, error: msg }); }

// ─── CORS-friendly GET + POST handlers ──────────────────────────
function doGet(e)  { return route(e.parameter); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (_) {}
  // merge query params
  if (e.parameter) Object.keys(e.parameter).forEach(k => p[k] = p[k] || e.parameter[k]);
  return route(p);
}

// ─── Router ─────────────────────────────────────────────────────
function route(p) {
  try {
    switch (p.action) {
      case 'read':          return doRead(p);
      case 'create':        return doCreate(p);
      case 'update':        return doUpdate(p);
      case 'delete':        return doDelete(p);
      case 'login':         return doLogin(p);
      case 'addUser':       return doAddUser(p);
      case 'updateUser':    return doUpdateUser(p);
      case 'deleteUser':    return doDeleteUser(p);
      default:              return fail('Unknown action: ' + p.action);
    }
  } catch (err) {
    return fail(err.toString());
  }
}

// ─── READ ───────────────────────────────────────────────────────
function doRead(p) {
  const sheet = getSheet(p.sheet);
  return ok(sheetToObjects(sheet));
}

// ─── CREATE ─────────────────────────────────────────────────────
function doCreate(p) {
  const sheet = getSheet(p.sheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = JSON.parse(p.row);
  if (!row.id) row.id = uid();
  const newRow = headers.map(h => row[h] !== undefined ? row[h] : '');
  sheet.appendRow(newRow);
  return ok(row);
}

// ─── UPDATE ─────────────────────────────────────────────────────
function doUpdate(p) {
  const sheet = getSheet(p.sheet);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const updates = JSON.parse(p.row);
  const idCol = headers.indexOf('id');
  for (let r = 1; r < data.length; r++) {
    if (data[r][idCol] === updates.id) {
      headers.forEach((h, c) => {
        if (updates[h] !== undefined) {
          sheet.getRange(r + 1, c + 1).setValue(updates[h]);
        }
      });
      return ok(updates);
    }
  }
  return fail('Row not found: ' + updates.id);
}

// ─── DELETE ─────────────────────────────────────────────────────
function doDelete(p) {
  const sheet = getSheet(p.sheet);
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('id');
  for (let r = 1; r < data.length; r++) {
    if (data[r][idCol] === p.id) {
      sheet.deleteRow(r + 1);
      return ok({ deleted: p.id });
    }
  }
  return fail('Row not found: ' + p.id);
}

// ─── LOGIN ──────────────────────────────────────────────────────
function doLogin(p) {
  const sheet = getSheet('Users');
  const users = sheetToObjects(sheet);
  const user = users.find(u => u.username && u.username.toLowerCase() === p.username.toLowerCase());
  if (!user) return fail('User not found');
  const hash = simpleHash(p.password);
  if (user.passwordHash !== hash) return fail('Invalid password');
  return ok({ id: user.id, name: user.name, role: user.role, email: user.email, color: user.color });
}

// ─── USER MANAGEMENT ────────────────────────────────────────────
function doAddUser(p) {
  const sheet = getSheet('Users');
  const row = JSON.parse(p.row);
  if (!row.id) row.id = uid();
  if (row.password) {
    row.passwordHash = simpleHash(row.password);
    delete row.password;
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newRow = headers.map(h => row[h] !== undefined ? row[h] : '');
  sheet.appendRow(newRow);
  return ok({ id: row.id, name: row.name, role: row.role });
}

function doUpdateUser(p) {
  const sheet = getSheet('Users');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const updates = JSON.parse(p.row);
  if (updates.password) {
    updates.passwordHash = simpleHash(updates.password);
    delete updates.password;
  }
  const idCol = headers.indexOf('id');
  for (let r = 1; r < data.length; r++) {
    if (data[r][idCol] === updates.id) {
      headers.forEach((h, c) => {
        if (updates[h] !== undefined) {
          sheet.getRange(r + 1, c + 1).setValue(updates[h]);
        }
      });
      return ok(updates);
    }
  }
  return fail('User not found: ' + updates.id);
}

function doDeleteUser(p) {
  return doDelete({ sheet: 'Users', id: p.id });
}
