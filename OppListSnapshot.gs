function OppListSnapshot_createWeekly(deptKey) {
  var rows = OppListReader_getLiveRows(deptKey).rows || [];
  var sheet = getSharedSheet(OPP_LIST_SNAPSHOT_SHEET_NAME);
  if (!sheet) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    sheet = ss.insertSheet(OPP_LIST_SNAPSHOT_SHEET_NAME);
    sheet.getRange(1, 1, 1, 5).setValues([['snapshot_at', 'snapshot_date', 'opp_id', 'dept', 'payload_json']]);
  }

  var now = new Date();
  var dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  var appendRows = (rows || []).map(function(row) {
    var payload = JSON.parse(JSON.stringify(row));
    payload.snapshotDate = dateStr;
    // Task 4 guard: omit proposalProductIds for depts without proposalProducts feature
    if (!DEPT_CONFIG[deptKey].features.proposalProducts) {
      delete payload.proposalProductIds;
    }
    return [now, dateStr, row.oppId || '', deptKey, JSON.stringify(payload)];
  });

  OppListSnapshot_deleteByDate_(deptKey, sheet, dateStr);
  if (appendRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, 5).setValues(appendRows);
  }
  OppListSnapshot_trimOld_(deptKey, sheet);
  return { ok: true, date: dateStr, count: appendRows.length };
}

function OppListSnapshot_getSnapshotDates(deptKey) {
  var sheet = getSharedSheet(OPP_LIST_SNAPSHOT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  var seen = {};
  var dates = [];
  values.forEach(function(row) {
    var dateStr = String(row[1] || '').trim();
    var dept = String(row[3] || '').trim();
    if (!dateStr || dept !== deptKey || seen[dateStr]) return;
    seen[dateStr] = true;
    dates.push(dateStr);
  });
  dates.sort(function(a, b) { return a > b ? -1 : a < b ? 1 : 0; });
  return dates;
}

function OppListSnapshot_getByDate(deptKey, dateStr) {
  var sheet = getSharedSheet(OPP_LIST_SNAPSHOT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { date: dateStr, rows: [] };

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  var rows = [];
  values.forEach(function(row) {
    var dept = String(row[3] || '').trim();
    if (String(row[1] || '').trim() !== dateStr) return;
    if (dept !== deptKey) return;
    try {
      rows.push(JSON.parse(String(row[4] || '{}')));
    } catch (e) {}
  });
  return { date: dateStr, rows: rows };
}

function OppListSnapshot_setupWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'createOppSnapshot') ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger('createOppSnapshot')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(3)
    .create();
  return { ok: true };
}

function OppListSnapshot_deleteByDate_(deptKey, sheet, dateStr) {
  if (sheet.getLastRow() < 2) return;
  for (var row = sheet.getLastRow(); row >= 2; row--) {
    var rowValues = sheet.getRange(row, 2, 1, 3).getValues()[0];
    var rowDate = String(rowValues[0] || '').trim();
    var rowDept = String(rowValues[2] || '').trim();
    if (rowDate === dateStr && rowDept === deptKey) {
      sheet.deleteRow(row);
    }
  }
}

function OppListSnapshot_trimOld_(deptKey, sheet) {
  var dates = OppListSnapshot_getSnapshotDates(deptKey);
  if (dates.length <= 52) return;
  var keep = {};
  dates.slice(0, 52).forEach(function(dateStr) { keep[dateStr] = true; });
  for (var row = sheet.getLastRow(); row >= 2; row--) {
    var rowValues = sheet.getRange(row, 2, 1, 3).getValues()[0];
    var dateStr = String(rowValues[0] || '').trim();
    var rowDept = String(rowValues[2] || '').trim();
    if (rowDept === deptKey && dateStr && !keep[dateStr]) sheet.deleteRow(row);
  }
}
