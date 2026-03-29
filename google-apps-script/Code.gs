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

  // Notify assignee when a task is created and assigned to them
  if (p.sheet === 'Tasks' && row.memberId && row.title) {
    try {
      sendAssignmentEmail(row.title, row.memberId, row.due || '');
    } catch (emailErr) {
      Logger.log('Assignment email error: ' + emailErr.toString());
    }
  }

  return ok(row);
}

// ─── ASSIGNMENT NOTIFICATION ────────────────────────────────────
function sendAssignmentEmail(taskTitle, memberId, dueDate) {
  const users = sheetToObjects(getSheet('Users'));
  const assignee = users.find(u => u.id === memberId);
  if (!assignee || !assignee.email) return;

  const dueStr = dueDate ? ' (Due: ' + dueDate + ')' : '';
  MailApp.sendEmail({
    to: assignee.email,
    subject: '📋 New Task Assigned: ' + taskTitle,
    htmlBody: '<div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;padding:24px;">'
      + '<div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:8px;padding:20px;">'
      + '<h2 style="margin:0 0 8px;color:#1e40af;">📋 New Task Assigned</h2>'
      + '<p style="margin:0 0 16px;color:#374151;">Hi ' + assignee.name + ',</p>'
      + '<div style="background:white;border-radius:6px;padding:16px;border:1px solid #e5e7eb;">'
      + '<p style="margin:0;font-size:16px;font-weight:600;color:#111827;">' + taskTitle + dueStr + '</p>'
      + '</div>'
      + '<p style="margin:16px 0 0;color:#374151;">A new task has been assigned to you.</p>'
      + '</div>'
      + '<p style="margin:16px 0 0;font-size:12px;color:#9ca3af;text-align:center;">AssyEng Task Manager</p>'
      + '</div>'
  });
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
      // Check if task is being marked complete
      const wasStatus = data[r][headers.indexOf('status')];
      const isTaskComplete = p.sheet === 'Tasks' && updates.status === 'done' && wasStatus !== 'done';

      headers.forEach((h, c) => {
        if (updates[h] !== undefined) {
          sheet.getRange(r + 1, c + 1).setValue(updates[h]);
        }
      });

      // Send notification emails on task completion
      if (isTaskComplete) {
        try {
          const taskTitle = data[r][headers.indexOf('title')] || 'Untitled';
          const memberId = updates.memberId || data[r][headers.indexOf('memberId')];
          sendCompletionEmails(taskTitle, memberId, updates.id);
        } catch (emailErr) {
          // Don't fail the update if email fails
          Logger.log('Email notification error: ' + emailErr.toString());
        }
      }

      return ok(updates);
    }
  }
  return fail('Row not found: ' + updates.id);
}

// ─── EMAIL NOTIFICATIONS ────────────────────────────────────────
function sendCompletionEmails(taskTitle, memberId, taskId) {
  const usersSheet = getSheet('Users');
  const users = sheetToObjects(usersSheet);

  const assignee = users.find(u => u.id === memberId);
  const managers = users.filter(u => u.role === 'manager');

  const subject = '✅ Task Completed: ' + taskTitle;

  // Email the assignee
  if (assignee && assignee.email) {
    MailApp.sendEmail({
      to: assignee.email,
      subject: subject,
      htmlBody: buildEmailHtml(taskTitle, assignee.name, 'Your task has been marked as complete.', false)
    });
  }

  // Email all managers (skip if they are also the assignee)
  for (const mgr of managers) {
    if (mgr.email && (!assignee || mgr.id !== assignee.id)) {
      const completedBy = assignee ? assignee.name : 'Unknown';
      MailApp.sendEmail({
        to: mgr.email,
        subject: subject,
        htmlBody: buildEmailHtml(taskTitle, mgr.name, 'Task completed by ' + completedBy + '.', true)
      });
    }
  }
}

function buildEmailHtml(taskTitle, recipientName, message, isManagerView) {
  return '<div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;padding:24px;">'
    + '<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;padding:20px;">'
    + '<h2 style="margin:0 0 8px;color:#166534;">✅ Task Complete</h2>'
    + '<p style="margin:0 0 16px;color:#374151;">Hi ' + recipientName + ',</p>'
    + '<div style="background:white;border-radius:6px;padding:16px;border:1px solid #e5e7eb;">'
    + '<p style="margin:0;font-size:16px;font-weight:600;color:#111827;">' + taskTitle + '</p>'
    + '</div>'
    + '<p style="margin:16px 0 0;color:#374151;">' + message + '</p>'
    + '</div>'
    + '<p style="margin:16px 0 0;font-size:12px;color:#9ca3af;text-align:center;">AssyEng Task Manager</p>'
    + '</div>';
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
