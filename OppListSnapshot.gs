var OPP_LIST_SNAPSHOT_SHEET_NAME = '案件リストスナップショット';

function OppListSnapshot_createWeekly() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var rows = OppListReader_getLiveRows().rows || [];
  var sheet = ss.getSheetByName(OPP_LIST_SNAPSHOT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(OPP_LIST_SNAPSHOT_SHEET_NAME);
    sheet.getRange(1, 1, 1, 5).setValues([['snapshot_at', 'snapshot_date', 'opp_id', 'dept', 'payload_json']]);
  }

  var now = new Date();
  var dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  var appendRows = (rows || []).map(function(row) {
    var payload = JSON.parse(JSON.stringify(row));
    payload.snapshotDate = dateStr;
    return [now, dateStr, row.oppId || '', row.dept || '', JSON.stringify(payload)];
  });

  OppListSnapshot_deleteByDate_(sheet, dateStr);
  if (appendRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, 5).setValues(appendRows);
  }
  OppListSnapshot_trimOld_(sheet);
  return { ok: true, date: dateStr, count: appendRows.length };
}

function OppListSnapshot_getSnapshotDates() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(OPP_LIST_SNAPSHOT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  var seen = {};
  var dates = [];
  values.forEach(function(row) {
    var dateStr = String(row[0] || '').trim();
    if (!dateStr || seen[dateStr]) return;
    seen[dateStr] = true;
    dates.push(dateStr);
  });
  dates.sort(function(a, b) { return a > b ? -1 : a < b ? 1 : 0; });
  return dates;
}

function OppListSnapshot_getByDate(dateStr) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(OPP_LIST_SNAPSHOT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { date: dateStr, rows: [] };

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  var rows = [];
  values.forEach(function(row) {
    if (String(row[1] || '').trim() !== dateStr) return;
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

function OppListSnapshot_deleteByDate_(sheet, dateStr) {
  if (sheet.getLastRow() < 2) return;
  for (var row = sheet.getLastRow(); row >= 2; row--) {
    if (String(sheet.getRange(row, 2).getValue() || '').trim() === dateStr) {
      sheet.deleteRow(row);
    }
  }
}

function OppListSnapshot_trimOld_(sheet) {
  var dates = OppListSnapshot_getSnapshotDates();
  if (dates.length <= 52) return;
  var keep = {};
  dates.slice(0, 52).forEach(function(dateStr) { keep[dateStr] = true; });
  for (var row = sheet.getLastRow(); row >= 2; row--) {
    var dateStr = String(sheet.getRange(row, 2).getValue() || '').trim();
    if (dateStr && !keep[dateStr]) sheet.deleteRow(row);
  }
}
