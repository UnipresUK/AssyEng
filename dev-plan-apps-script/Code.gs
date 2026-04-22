// Engineer Development Plan - Google Apps Script Backend
var SPREADSHEET_ID = '1G55YaGAiMOuZZTfZnI3xFVPRHV780-kBA2MMzReHq4I';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getOrCreate(name, headers) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var results = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === '' || data[i][0] === null || data[i][0] === undefined) continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j] !== undefined ? data[i][j] : '';
    }
    results.push(obj);
  }
  return results;
}

function jsonResp(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function ok(data)  { return jsonResp({ ok: true,  data: data  }); }
function fail(msg) { return jsonResp({ ok: false, error: msg  }); }

// Entry points
function doGet(e)  { return route(e.parameter || {}); }
function doPost(e) {
  var p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) {}
  if (e.parameter) {
    for (var k in e.parameter) {
      if (!p[k]) p[k] = e.parameter[k];
    }
  }
  return route(p);
}

function route(p) {
  try {
    var action = p.action || '';
    if (action === 'getAll')      return doGetAll();
    if (action === 'updateSkill') return doUpdateSkill(p);
    if (action === 'updateGoal')  return doUpdateGoal(p);
    if (action === 'updateNotes') return doUpdateNotes(p);
    if (action === 'addGoal')     return doAddGoal(p);
    if (action === 'setup')       return doSetup();
    return fail('Unknown action: ' + action);
  } catch (err) {
    return fail(err.toString());
  }
}

// GET ALL
function doGetAll() {
  var engSheet   = getOrCreate('Engineers', ['id','name','initials','tier','role','started','manager','color','colorBg','notes']);
  var skillSheet = getOrCreate('Skills',    ['engineerId','drw','asm','qlt','hse','erp','rca','lean','com','tool','prod']);
  var goalSheet  = getOrCreate('Goals',     ['id','engineerId','title','target','status']);

  var engineers = sheetToObjects(engSheet);
  var skillRows = sheetToObjects(skillSheet);
  var goalRows  = sheetToObjects(goalSheet);

  var result = [];
  for (var i = 0; i < engineers.length; i++) {
    var eng = engineers[i];

    // Find skill row for this engineer
    var skillRow = null;
    for (var s = 0; s < skillRows.length; s++) {
      if (String(skillRows[s].engineerId) === String(eng.id)) {
        skillRow = skillRows[s];
        break;
      }
    }
    if (!skillRow) skillRow = {};

    var skills = {
      drw:  Number(skillRow.drw  || 0),
      asm:  Number(skillRow.asm  || 0),
      qlt:  Number(skillRow.qlt  || 0),
      hse:  Number(skillRow.hse  || 0),
      erp:  Number(skillRow.erp  || 0),
      rca:  Number(skillRow.rca  || 0),
      lean: Number(skillRow.lean || 0),
      com:  Number(skillRow.com  || 0),
      tool: Number(skillRow.tool || 0),
      prod: Number(skillRow.prod || 0)
    };

    // Collect goals for this engineer
    var goals = [];
    for (var g = 0; g < goalRows.length; g++) {
      if (String(goalRows[g].engineerId) === String(eng.id)) {
        var t = goalRows[g].target;
        var tStr = (t instanceof Date) ? t.toISOString().split('T')[0] : String(t);
        goals.push({
          id:     goalRows[g].id,
          title:  goalRows[g].title,
          target: tStr,
          status: goalRows[g].status
        });
      }
    }

    // Build output object without spread operator
    var out = {};
    for (var key in eng) { out[key] = eng[key]; }
    out.skills = skills;
    out.goals  = goals;
    result.push(out);
  }

  return ok(result);
}

// UPDATE SKILL
function doUpdateSkill(p) {
  var sheet   = getOrCreate('Skills', ['engineerId','drw','asm','qlt','hse','erp','rca','lean','com','tool','prod']);
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol   = headers.indexOf('engineerId');
  var skillCol = headers.indexOf(p.skill);
  if (skillCol < 0) return fail('Unknown skill: ' + p.skill);

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idCol]) === String(p.engineerId)) {
      sheet.getRange(r + 1, skillCol + 1).setValue(Number(p.level));
      return ok({ engineerId: p.engineerId, skill: p.skill, level: p.level });
    }
  }
  // Row not found - create it
  var newRow = [];
  for (var h = 0; h < headers.length; h++) {
    if (headers[h] === 'engineerId') newRow.push(p.engineerId);
    else if (headers[h] === p.skill) newRow.push(Number(p.level));
    else newRow.push(0);
  }
  sheet.appendRow(newRow);
  return ok({ engineerId: p.engineerId, skill: p.skill, level: p.level, created: true });
}

// UPDATE GOAL STATUS
function doUpdateGoal(p) {
  var sheet   = getOrCreate('Goals', ['id','engineerId','title','target','status']);
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol     = headers.indexOf('id');
  var statusCol = headers.indexOf('status');

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idCol]) === String(p.goalId)) {
      sheet.getRange(r + 1, statusCol + 1).setValue(p.status);
      return ok({ goalId: p.goalId, status: p.status });
    }
  }
  return fail('Goal not found: ' + p.goalId);
}

// UPDATE NOTES
function doUpdateNotes(p) {
  var sheet   = getOrCreate('Engineers', ['id','name','initials','tier','role','started','manager','color','colorBg','notes']);
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol   = headers.indexOf('id');
  var notesCol = headers.indexOf('notes');

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idCol]) === String(p.engineerId)) {
      sheet.getRange(r + 1, notesCol + 1).setValue(p.notes);
      return ok({ engineerId: p.engineerId });
    }
  }
  return fail('Engineer not found: ' + p.engineerId);
}

// ADD GOAL
function doAddGoal(p) {
  var sheet  = getOrCreate('Goals', ['id','engineerId','title','target','status']);
  var goalId = 'G' + Date.now().toString(36);
  sheet.appendRow([goalId, p.engineerId, p.title, p.target, p.status || 'pending']);
  return ok({ id: goalId, engineerId: p.engineerId, title: p.title, target: p.target, status: p.status || 'pending' });
}

// SETUP - run once manually to seed sample data
function doSetup() {
  var engSheet   = getOrCreate('Engineers', ['id','name','initials','tier','role','started','manager','color','colorBg','notes']);
  var skillSheet = getOrCreate('Skills',    ['engineerId','drw','asm','qlt','hse','erp','rca','lean','com','tool','prod']);
  var goalSheet  = getOrCreate('Goals',     ['id','engineerId','title','target','status']);

  var existing = engSheet.getDataRange().getValues();
  if (existing.length > 1) {
    return ok({ message: 'Already seeded - ' + (existing.length - 1) + ' engineers found. Clear sheets to reset.' });
  }

  var engineers = [
    ['E01','Tom Ashworth','TA','junior','Junior Assembly Engineer','2024-09-01','S. Richardson','#f59e0b','rgba(245,158,11,.15)','Tom joined straight from college. Strong hands-on ability; needs to build confidence reading complex drawings.'],
    ['E02','Priya Nair','PN','graduate','Graduate Engineer (Out of Time)','2023-09-01','S. Richardson','#06b6d4','rgba(6,182,212,.15)','Priya completed apprenticeship Sept 2023. Technically solid; developing confidence in problem solving and RCA.'],
    ['E03','Jake Holden','JH','junior','Junior Assembly Engineer','2025-01-06','S. Richardson','#f59e0b','rgba(245,158,11,.15)','New starter January 2025. Focus on structured shadowing and safety compliance first.'],
    ['E04','Amelia Foster','AF','graduate','Graduate Engineer (Out of Time)','2023-09-01','S. Richardson','#06b6d4','rgba(6,182,212,.15)','Progressing well and close to full sign-off. ERP and Lean are the remaining development areas.'],
    ['E05','Ryan Blackwell','RB','junior','Junior Assembly Engineer','2024-06-03','S. Richardson','#f59e0b','rgba(245,158,11,.15)','Good practical ability and safety-conscious. Needs to develop quality awareness and drawing interpretation.'],
    ['E06','Sophie Marsh','SM','mid','Assembly Engineer','2022-09-01','S. Richardson','#10b981','rgba(16,185,129,.15)','Most experienced on the plan; acts as informal mentor. Targeted for IEng application this year.']
  ];

  var skills = [
    ['E01',1,2,1,2,1,1,1,2,2,1],
    ['E02',3,2,2,3,2,2,2,3,2,2],
    ['E03',0,1,1,2,0,0,0,1,1,0],
    ['E04',3,3,3,3,2,2,2,2,3,2],
    ['E05',2,2,1,3,1,1,1,2,2,1],
    ['E06',4,4,3,4,3,3,3,3,4,3]
  ];

  var goals = [
    ['G001','E01','Complete Engineering Drawing Fundamentals module','2025-06-30','progress'],
    ['G002','E01','Shadow QC inspection on 3 builds','2025-05-31','complete'],
    ['G003','E01','Pass HSE induction and toolbox talks','2025-04-30','complete'],
    ['G004','E01','Complete ERP training (SAP basics)','2025-07-31','pending'],
    ['G005','E01','Achieve level 2 on all core skills','2025-12-31','pending'],
    ['G006','E02','Lead full build sign-off independently','2025-05-31','complete'],
    ['G007','E02','Complete Lean Six Sigma Yellow Belt','2025-07-31','progress'],
    ['G008','E02','RCA on last 2 NCRs raised in section','2025-04-30','overdue'],
    ['G009','E02','Mentor new junior engineer for 1 quarter','2025-09-30','pending'],
    ['G010','E02','Achieve level 3 on all core skills','2025-12-31','pending'],
    ['G011','E03','Complete site induction and HSE mandatory training','2025-01-31','complete'],
    ['G012','E03','Shadow assembly on PZ1D build all stations','2025-04-30','progress'],
    ['G013','E03','Complete drawing and GD&T awareness session','2025-05-31','pending'],
    ['G014','E03','First solo workstation sign-off with supervisor','2025-07-31','pending'],
    ['G015','E04','Complete full NCR root-cause investigation','2025-03-31','complete'],
    ['G016','E04','Achieve ERP level 3 transactions and reporting','2025-06-30','progress'],
    ['G017','E04','Lean improvement project submission','2025-08-31','pending'],
    ['G018','E04','Lead junior engineer development for 1 quarter','2025-09-30','pending'],
    ['G019','E04','Full competency sign-off by manager','2025-12-31','pending'],
    ['G020','E05','Complete HSE advanced module','2025-02-28','complete'],
    ['G021','E05','QC inspection shadowing 5 builds logged','2025-05-31','progress'],
    ['G022','E05','Drawing reading assessment pass 70 percent','2025-06-30','pending'],
    ['G023','E05','ERP system access and basic transactions','2025-07-31','pending'],
    ['G024','E06','Lean project lead 5S implementation','2025-03-31','complete'],
    ['G025','E06','Mentor two junior engineers through Q2','2025-06-30','progress'],
    ['G026','E06','Complete IEng application submission','2025-09-30','pending'],
    ['G027','E06','Full advanced sign-off on all skill areas','2025-12-31','pending']
  ];

  engSheet.getRange(2, 1, engineers.length, engineers[0].length).setValues(engineers);
  skillSheet.getRange(2, 1, skills.length, skills[0].length).setValues(skills);
  goalSheet.getRange(2, 1, goals.length, goals[0].length).setValues(goals);

  return ok({ message: 'Setup complete - ' + engineers.length + ' engineers seeded.' });
}
