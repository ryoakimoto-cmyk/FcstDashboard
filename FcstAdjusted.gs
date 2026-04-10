var FCST_STATE_HEADERS = ['担当者', '期間', 'Net', 'New+Exp', 'Churn', '最終更新日時', 'Note', 'dept'];

function FcstAdjusted_getState(deptKey) {
  return FcstAdjusted_getState_(deptKey);
}

function FcstAdjusted_getAll(deptKey) {
  return FcstAdjusted_getState_(deptKey).adjusted;
}

function FcstAdjusted_getNotes(deptKey) {
  return FcstAdjusted_getState_(deptKey).notes;
}

function FcstAdjusted_save(deptKey, p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = FcstAdjusted_getOrCreateSheet_(deptKey);
    var lastRow = sheet.getLastRow();
    var name = String(p.name || '').trim();
    var period = String(p.period || '').trim();
    if (!name || !period) throw new Error('name / period is required');

    var existing = FcstAdjusted_findRow_(sheet, deptKey, name, period, lastRow);
    var rowIndex = existing.rowIndex;
    var prev = existing.record;
    var now = new Date();
    var record = [
      name,
      period,
      p.hasOwnProperty('net') ? (Number(p.net) || 0) : prev.net,
      p.hasOwnProperty('newExp') ? (Number(p.newExp) || 0) : prev.newExp,
      p.hasOwnProperty('churn') ? (Number(p.churn) || 0) : prev.churn,
      now,
      p.hasOwnProperty('note') ? String(p.note || '') : prev.note,
      deptKey
    ];

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, FCST_STATE_HEADERS.length).setValues([record]);
    } else {
      sheet.appendRow(record);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function FcstAdjusted_getState_(deptKey) {
  var sheet = FcstAdjusted_getSheet_(deptKey);
  if (!sheet) return { adjusted: {}, notes: {} };

  FcstAdjusted_ensureSchema_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { adjusted: {}, notes: {} };

  var values = sheet.getRange(2, 1, lastRow - 1, FCST_STATE_HEADERS.length).getValues();
  var adjusted = {};
  var notes = {};
  values.forEach(function(row) {
    var name = String(row[0] || '').trim();
    var period = String(row[1] || '').trim();
    var dept = String(row[7] || '').trim();
    if (dept !== deptKey || !name || !period) return;
    var key = name + '|' + period;
    adjusted[key] = {
      net: Number(row[2]) || 0,
      newExp: Number(row[3]) || 0,
      churn: Number(row[4]) || 0
    };
    notes[key] = String(row[6] || '');
  });

  return { adjusted: adjusted, notes: notes };
}

function FcstAdjusted_getOrCreateSheet_(deptKey) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getSharedSheet(FCST_ADJUSTED_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FCST_ADJUSTED_SHEET_NAME);
    sheet.getRange(1, 1, 1, FCST_STATE_HEADERS.length).setValues([FCST_STATE_HEADERS]);
  } else {
    FcstAdjusted_ensureSchema_(sheet);
  }
  return sheet;
}

function FcstAdjusted_getSheet_(deptKey) {
  var sheet = getSharedSheet(FCST_ADJUSTED_SHEET_NAME);
  if (sheet) FcstAdjusted_ensureSchema_(sheet);
  return sheet;
}

function FcstAdjusted_ensureSchema_(sheet) {
  var headerWidth = Math.max(sheet.getLastColumn(), FCST_STATE_HEADERS.length);
  if (headerWidth < FCST_STATE_HEADERS.length) headerWidth = FCST_STATE_HEADERS.length;
  var headers = sheet.getRange(1, 1, 1, headerWidth).getValues()[0];
  for (var i = 0; i < FCST_STATE_HEADERS.length; i++) {
    if (String(headers[i] || '') !== FCST_STATE_HEADERS[i]) {
      sheet.getRange(1, i + 1).setValue(FCST_STATE_HEADERS[i]);
    }
  }
  if (headers.indexOf('dept') === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('dept');
  }
}

function FcstAdjusted_findRow_(sheet, deptKey, name, period, lastRow) {
  var result = {
    rowIndex: -1,
    record: { net: 0, newExp: 0, churn: 0, note: '' }
  };
  if (lastRow < 2) return result;

  var values = sheet.getRange(2, 1, lastRow - 1, FCST_STATE_HEADERS.length).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][7] || '').trim() !== deptKey) continue;
    if (String(values[i][0] || '').trim() === name && String(values[i][1] || '').trim() === period) {
      result.rowIndex = i + 2;
      result.record = {
        net: Number(values[i][2]) || 0,
        newExp: Number(values[i][3]) || 0,
        churn: Number(values[i][4]) || 0,
        note: String(values[i][6] || '')
      };
      return result;
    }
  }
  return result;
}
