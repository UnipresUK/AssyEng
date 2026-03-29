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

// Parse notifyPrefs JSON safely, default all on
function getNotifyPrefs(user) {
  const defaults = { assigned: true, completed: true, weeklySummary: true, managerBcc: true };
  if (!user.notifyPrefs) return defaults;
  try {
    const parsed = typeof user.notifyPrefs === 'string' ? JSON.parse(user.notifyPrefs) : user.notifyPrefs;
    return Object.assign(defaults, parsed);
  } catch (_) { return defaults; }
}

// ─── CORS-friendly GET + POST handlers ──────────────────────────
function doGet(e)  { return route(e.parameter); }
function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (_) {}
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

  const prefs = getNotifyPrefs(assignee);
  if (!prefs.assigned) return;

  const dueStr = dueDate ? ' (Due: ' + dueDate + ')' : '';
  MailApp.sendEmail({
    to: assignee.email,
    subject: 'New Task Assigned: ' + taskTitle,
    htmlBody: buildEmailHtml('New Task Assigned', taskTitle, assignee.name,
      'A new task has been assigned to you.' + (dueStr ? '<br>Due: ' + dueDate : ''),
      '#eff6ff', '#bfdbfe', '#1e40af')
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
      const wasStatus = data[r][headers.indexOf('status')];
      const isTaskComplete = p.sheet === 'Tasks' && updates.status === 'done' && wasStatus !== 'done';

      headers.forEach((h, c) => {
        if (updates[h] !== undefined) {
          sheet.getRange(r + 1, c + 1).setValue(updates[h]);
        }
      });

      if (isTaskComplete) {
        try {
          const taskTitle = data[r][headers.indexOf('title')] || 'Untitled';
          const memberId = updates.memberId || data[r][headers.indexOf('memberId')];
          sendCompletionEmails(taskTitle, memberId, updates.id);
        } catch (emailErr) {
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
  const users = sheetToObjects(getSheet('Users'));
  const assignee = users.find(u => u.id === memberId);
  const managers = users.filter(u => u.role === 'manager');

  if (!assignee || !assignee.email) return;
  const assigneePrefs = getNotifyPrefs(assignee);

  // Build BCC list: managers who opted in to managerBcc (exclude assignee if they're a manager)
  const bccList = managers
    .filter(mgr => mgr.id !== assignee.id && mgr.email)
    .filter(mgr => { const p = getNotifyPrefs(mgr); return p.managerBcc !== false; })
    .map(mgr => mgr.email);

  // Email the assignee (if opted in), BCC managers who opted in
  if (assigneePrefs.completed) {
    const emailOpts = {
      to: assignee.email,
      subject: 'Task Completed: ' + taskTitle,
      htmlBody: buildEmailHtml('Task Complete', taskTitle, assignee.name,
        'Your task has been marked as complete.',
        '#f0fdf4', '#bbf7d0', '#166534')
    };
    if (bccList.length > 0) emailOpts.bcc = bccList.join(',');
    MailApp.sendEmail(emailOpts);
  } else if (bccList.length > 0) {
    // Assignee opted out, but managers still want to know — send direct to managers
    const completedBy = assignee.name || 'Unknown';
    for (const mgr of managers) {
      if (mgr.id === assignee.id || !mgr.email) continue;
      const mgrPrefs = getNotifyPrefs(mgr);
      if (mgrPrefs.managerBcc === false) continue;
      MailApp.sendEmail({
        to: mgr.email,
        subject: 'Task Completed: ' + taskTitle,
        htmlBody: buildEmailHtml('Task Complete', taskTitle, mgr.name,
          'Task completed by ' + completedBy + '.',
          '#f0fdf4', '#bbf7d0', '#166534')
      });
    }
  }
}

function buildEmailHtml(heading, taskTitle, recipientName, message, bgColor, borderColor, headingColor) {
  return '<div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;padding:24px;">'
    + '<div style="background:' + bgColor + ';border:1px solid ' + borderColor + ';border-radius:8px;padding:20px;">'
    + '<h2 style="margin:0 0 8px;color:' + headingColor + ';">' + heading + '</h2>'
    + '<p style="margin:0 0 16px;color:#374151;">Hi ' + recipientName + ',</p>'
    + '<div style="background:white;border-radius:6px;padding:16px;border:1px solid #e5e7eb;">'
    + '<p style="margin:0;font-size:16px;font-weight:600;color:#111827;">' + taskTitle + '</p>'
    + '</div>'
    + '<p style="margin:16px 0 0;color:#374151;">' + message + '</p>'
    + '</div>'
    + '<p style="margin:16px 0 0;font-size:12px;color:#9ca3af;text-align:center;">AssyEng Task Manager</p>'
    + '</div>';
}

// ─── WEEKLY SUMMARY EMAIL (trigger every Monday morning) ────────
function sendWeeklySummary() {
  const users = sheetToObjects(getSheet('Users'));
  const allTasks = sheetToObjects(getSheet('Tasks'));

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const endOfWeek = new Date(today);
  endOfWeek.setDate(endOfWeek.getDate() + (5 - endOfWeek.getDay())); // Friday
  endOfWeek.setHours(23, 59, 59, 999);

  for (const user of users) {
    if (!user.email) continue;
    const prefs = getNotifyPrefs(user);
    if (!prefs.weeklySummary) continue;

    // Get this user's open tasks
    const myTasks = allTasks.filter(t =>
      t.memberId === user.id && t.status !== 'done'
    );
    if (myTasks.length === 0) continue;

    // Categorise: overdue, this week, later
    const overdue = [];
    const thisWeek = [];
    const later = [];
    const noDue = [];

    for (const t of myTasks) {
      if (!t.due) { noDue.push(t); continue; }
      const d = new Date(t.due);
      d.setHours(0, 0, 0, 0);
      if (d < today) overdue.push(t);
      else if (d <= endOfWeek) thisWeek.push(t);
      else later.push(t);
    }

    // Sort by priority then due date
    const priOrder = { high: 0, medium: 1, low: 2 };
    const sortFn = (a, b) => {
      const pa = priOrder[a.priority] !== undefined ? priOrder[a.priority] : 1;
      const pb = priOrder[b.priority] !== undefined ? priOrder[b.priority] : 1;
      if (pa !== pb) return pa - pb;
      if (a.due && b.due) return new Date(a.due) - new Date(b.due);
      return a.due ? -1 : 1;
    };
    overdue.sort(sortFn);
    thisWeek.sort(sortFn);
    later.sort(sortFn);

    // Build email HTML
    let body = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;padding:24px;">';
    body += '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:24px;">';
    body += '<h2 style="margin:0 0 4px;color:#1e293b;">Weekly Task Summary</h2>';
    body += '<p style="margin:0 0 20px;color:#64748b;font-size:14px;">' + formatDate(today) + ' — ' + formatDate(endOfWeek) + '</p>';
    body += '<p style="margin:0 0 20px;color:#374151;">Hi ' + user.name + ', here\'s your week at a glance:</p>';

    // Stats bar
    body += '<div style="display:flex;gap:12px;margin-bottom:20px;">';
    body += statBadge(overdue.length, 'Overdue', '#fef2f2', '#dc2626');
    body += statBadge(thisWeek.length, 'This Week', '#eff6ff', '#2563eb');
    body += statBadge(later.length, 'Upcoming', '#f0fdf4', '#16a34a');
    body += '</div>';

    if (overdue.length > 0) {
      body += taskSection('Overdue — Action Required', overdue, '#dc2626', '#fef2f2');
    }
    if (thisWeek.length > 0) {
      body += taskSection('Due This Week', thisWeek, '#2563eb', '#eff6ff');
    }
    if (later.length > 0 && later.length <= 10) {
      body += taskSection('Upcoming', later, '#16a34a', '#f0fdf4');
    } else if (later.length > 10) {
      body += taskSection('Upcoming (next 10)', later.slice(0, 10), '#16a34a', '#f0fdf4');
      body += '<p style="color:#64748b;font-size:13px;">+ ' + (later.length - 10) + ' more upcoming tasks</p>';
    }
    if (noDue.length > 0) {
      body += taskSection('No Due Date', noDue.slice(0, 5), '#64748b', '#f8fafc');
      if (noDue.length > 5) {
        body += '<p style="color:#64748b;font-size:13px;">+ ' + (noDue.length - 5) + ' more without due dates</p>';
      }
    }

    body += '</div>';
    body += '<p style="margin:16px 0 0;font-size:12px;color:#9ca3af;text-align:center;">AssyEng Task Manager — Weekly Summary</p>';
    body += '</div>';

    MailApp.sendEmail({
      to: user.email,
      subject: 'Your Week: ' + overdue.length + ' overdue, ' + thisWeek.length + ' due this week',
      htmlBody: body
    });
  }
}

// ─── Weekly summary helpers ─────────────────────────────────────
function formatDate(d) {
  const days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return days[d.getDay()] + ' ' + d.getDate() + ' ' + months[d.getMonth()];
}

function statBadge(count, label, bg, color) {
  return '<div style="background:' + bg + ';border-radius:8px;padding:12px 16px;text-align:center;flex:1;">'
    + '<div style="font-size:24px;font-weight:700;color:' + color + ';">' + count + '</div>'
    + '<div style="font-size:12px;color:' + color + ';opacity:0.8;">' + label + '</div>'
    + '</div>';
}

function taskSection(title, taskList, color, bg) {
  const priLabels = { high: 'HIGH', medium: 'MED', low: 'LOW' };
  const priColors = { high: '#dc2626', medium: '#f59e0b', low: '#6b7280' };
  let html = '<div style="margin-bottom:16px;">';
  html += '<h3 style="margin:0 0 8px;color:' + color + ';font-size:14px;text-transform:uppercase;letter-spacing:0.5px;">' + title + '</h3>';
  html += '<div style="background:white;border-radius:6px;border:1px solid #e5e7eb;overflow:hidden;">';
  for (let i = 0; i < taskList.length; i++) {
    const t = taskList[i];
    const border = i > 0 ? 'border-top:1px solid #f3f4f6;' : '';
    const pri = t.priority || 'medium';
    const dueStr = t.due ? new Date(t.due).toLocaleDateString('en-GB', { weekday: 'short', day: 'numeric', month: 'short' }) : '';
    html += '<div style="padding:10px 14px;' + border + 'display:flex;align-items:center;gap:10px;">';
    html += '<span style="font-size:10px;font-weight:700;color:' + (priColors[pri] || '#6b7280') + ';background:' + bg + ';padding:2px 6px;border-radius:4px;">' + (priLabels[pri] || 'MED') + '</span>';
    html += '<span style="flex:1;color:#111827;font-size:14px;">' + (t.title || 'Untitled') + '</span>';
    if (dueStr) {
      html += '<span style="font-size:12px;color:#64748b;white-space:nowrap;">' + dueStr + '</span>';
    }
    html += '</div>';
  }
  html += '</div></div>';
  return html;
}

// ─── Setup trigger (run once manually) ──────────────────────────
function setupWeeklyTrigger() {
  // Delete existing weekly triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendWeeklySummary') ScriptApp.deleteTrigger(t);
  });
  // Create new trigger: every Monday at 7am
  ScriptApp.newTrigger('sendWeeklySummary')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();
  Logger.log('Weekly summary trigger created for Monday 7am');
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
  return ok({
    id: user.id, name: user.name, role: user.role, email: user.email,
    color: user.color, notifyPrefs: getNotifyPrefs(user)
  });
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
