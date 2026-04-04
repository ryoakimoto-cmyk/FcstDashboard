function FcstSnapshot_create(members, notesMap) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(FCST_SNAPSHOT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(FCST_SNAPSHOT_SHEET_NAME);
    sheet.getRange(1, 1, 1, 4).setValues([['\u65e5\u6642', '\u62c5\u5f53\u8005', '\u671f\u9593', '\u30c7\u30fc\u30bf']]);
  }
  var now = new Date();
  var periods = ['Q', 'M5', 'M6', 'M7'];
  var prevMetricMap = FcstSnapshot_getLatestMetricMap_();
  var notes = notesMap || {};
  var rows = [];
  (members || []).forEach(function(member) {
    if (!member || !member.name) return;
    periods.forEach(function(period) {
      var metric = member[period] || {};
      var mapKey = member.name + '|' + period;
      var payload = {};
      Object.keys(metric).forEach(function(k) { payload[k] = metric[k]; });
      payload.weekOverWeek = FcstSnapshot_buildWeekOverWeek_(metric, prevMetricMap[mapKey] || {});
      payload.note = String(notes[mapKey] || '');
      payload.__meta = {
        isTotal: !!member.isTotal,
        group: member.group || '',
        dept: member.dept || '',
        name: member.name,
      };
      rows.push([now, member.name, period, JSON.stringify(payload)]);
    });
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  }
  FcstSnapshot_trimOld_(sheet);
  return { ok: true, count: rows.length };
}

function FcstSnapshot_getLatestMetricMap_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(FCST_SNAPSHOT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 1) return {};
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange(1, 1, lastRow, 4).getValues();
  var latestKey = FcstSnapshot_getLatestTimestampKey_(values);
  if (!latestKey) return {};
  var metricMap = {};
  values.forEach(function(row) {
    var d = row[0];
    if (!(d instanceof Date) || isNaN(d)) return;
    if (Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') !== latestKey) return;
    var name = String(row[1] || '').trim();
    var period = String(row[2] || '').trim();
    var metric;
    try {
      metric = JSON.parse(String(row[3] || '{}'));
    } catch (e) {
      metric = {};
    }
    metricMap[name + '|' + period] = FcstSnapshot_extractMetricPayload_(metric);
  });
  return metricMap;
}

function FcstSnapshot_trimOld_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return;
  var dates = sheet.getRange(1, 1, lastRow, 1).getValues().map(function(r) { return r[0]; });
  var seenDates = [];
  dates.forEach(function(d) {
    if (!(d instanceof Date) || isNaN(d)) return;
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    if (seenDates.indexOf(key) === -1) seenDates.push(key);
  });
  if (seenDates.length <= 52) return;
  var oldKey = seenDates[0];
  for (var i = lastRow; i >= 1; i--) {
    var d = dates[i - 1];
    if (!(d instanceof Date) || isNaN(d)) continue;
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    if (key === oldKey) sheet.deleteRow(i);
  }
}

function FcstSnapshot_getWeekOverWeek() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(FCST_SNAPSHOT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 1) return {};
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange(1, 1, lastRow, 4).getValues();
  var dateKeys = [];
  values.forEach(function(row) {
    var d = row[0];
    if (!(d instanceof Date) || isNaN(d)) return;
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    if (dateKeys.indexOf(key) === -1) dateKeys.push(key);
  });
  dateKeys.sort(function(a, b) { return b < a ? -1 : b > a ? 1 : 0; });
  if (dateKeys.length < 2) return {};
  var latestKey = dateKeys[0];
  var prevKey = dateKeys[1];
  var latestMap = {};
  var prevMap = {};
  values.forEach(function(row) {
    var d = row[0];
    if (!(d instanceof Date) || isNaN(d)) return;
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    var name = String(row[1] || '').trim();
    var period = String(row[2] || '').trim();
    var mapKey = name + '|' + period;
    var metric;
    try {
      metric = JSON.parse(String(row[3] || '{}'));
    } catch (e) {
      metric = {};
    }
    if (key === latestKey) latestMap[mapKey] = metric;
    if (key === prevKey) prevMap[mapKey] = metric;
  });
  var result = {};
  Object.keys(latestMap).forEach(function(mapKey) {
    var latest = latestMap[mapKey];
    var prev = prevMap[mapKey] || {};
    result[mapKey] = FcstSnapshot_buildWeekOverWeek_(latest, prev);
  });
  return result;
}

// 最新スナップショットの members データを返す（live view の先週比計算用）
function FcstSnapshot_getLatestMembers() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(FCST_SNAPSHOT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 1) return null;
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange(1, 1, lastRow, 4).getValues();
  var latestKey = FcstSnapshot_getLatestTimestampKey_(values);
  if (!latestKey) return null;
  var data = FcstSnapshot_getDataByTimestampKey_(latestKey, values);
  return { members: data.members, date: data.date };
}

function FcstSnapshot_getSnapshotDates() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(FCST_SNAPSHOT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 1) return [];
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange(1, 1, lastRow, 1).getValues();
  var seen = {};
  var dates = [];
  values.forEach(function(row) {
    var d = row[0];
    if (!(d instanceof Date) || isNaN(d)) return;
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
    if (!seen[key]) {
      seen[key] = true;
      dates.push(key);
    }
  });
  dates.sort(function(a, b) { return a > b ? -1 : a < b ? 1 : 0; });
  return dates;
}

function FcstSnapshot_getDataByDate(dateStr) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(FCST_SNAPSHOT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 1) return { members: [], fcstAdjusted: {}, weekOverWeekMap: {}, date: dateStr };
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange(1, 1, lastRow, 4).getValues();
  var latestKeyForDate = '';
  values.forEach(function(row) {
    var d = row[0];
    if (!(d instanceof Date) || isNaN(d)) return;
    if (Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') !== dateStr) return;
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    if (!latestKeyForDate || key > latestKeyForDate) latestKeyForDate = key;
  });
  if (!latestKeyForDate) return { members: [], fcstAdjusted: {}, weekOverWeekMap: {}, notes: {}, date: dateStr };
  return FcstSnapshot_getDataByTimestampKey_(latestKeyForDate, values);
}

function FcstSnapshot_getDataByTimestampKey_(timestampKey, valuesOpt) {
  var rows = valuesOpt;
  if (!rows) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(FCST_SNAPSHOT_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 1) return { members: [], fcstAdjusted: {}, weekOverWeekMap: {}, notes: {}, date: '' };
    rows = sheet.getRange(1, 1, sheet.getLastRow(), 4).getValues();
  }
  var memberMap = {};
  var fcstAdjusted = {};
  var weekOverWeekMap = {};
  var notes = {};
  rows.forEach(function(row) {
    var d = row[0];
    if (!(d instanceof Date) || isNaN(d)) return;
    if (Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') !== timestampKey) return;
    var name = String(row[1] || '').trim();
    var period = String(row[2] || '').trim();
    var payload;
    try {
      payload = JSON.parse(String(row[3] || '{}'));
    } catch (e) {
      payload = {};
    }

    if (!memberMap[name]) {
      var meta = payload.__meta || {};
      memberMap[name] = {
        name: name,
        isTotal: !!meta.isTotal,
        group: meta.group || '',
        dept: meta.dept || '',
        sortOrder: 0,
        Q: {},
        M5: {},
        M6: {},
        M7: {},
      };
    }

    var metric = {};
    Object.keys(payload).forEach(function(k) {
      if (k !== '__meta' && k !== 'weekOverWeek' && k !== 'note') metric[k] = payload[k];
    });
    memberMap[name][period] = metric;

    var mapKey = name + '|' + period;
    fcstAdjusted[mapKey] = payload.fcstAdjusted || { net: 0, newExp: 0, churn: 0 };
    weekOverWeekMap[mapKey] = payload.weekOverWeek || {};
    notes[mapKey] = String(payload.note || '');
  });

  var members = Object.keys(memberMap).map(function(n) { return memberMap[n]; });
  return {
    members: members,
    fcstAdjusted: fcstAdjusted,
    weekOverWeekMap: weekOverWeekMap,
    notes: notes,
    date: timestampKey.slice(0, 10),
    timestampKey: timestampKey
  };
}

function FcstSnapshot_setupWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'createSnapshot') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('createSnapshot')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(3)
    .create();
  return { ok: true };
}

function FcstSnapshot_getLatestTimestampKey_(values) {
  var latestKey = '';
  (values || []).forEach(function(row) {
    var d = row[0];
    if (!(d instanceof Date) || isNaN(d)) return;
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    if (!latestKey || key > latestKey) latestKey = key;
  });
  return latestKey;
}

function FcstSnapshot_extractMetricPayload_(payload) {
  var metric = {};
  Object.keys(payload || {}).forEach(function(k) {
    if (k === '__meta' || k === 'weekOverWeek' || k === 'note') return;
    metric[k] = payload[k];
  });
  return metric;
}

function FcstSnapshot_buildWeekOverWeek_(currentMetric, prevMetric) {
  var metricKeys = ['fcstAdjusted', 'fcstCommit', 'confirmed', 'expectedMrr', 'received', 'debtMgmt', 'debtMgmtLite', 'expense'];
  var result = {};
  metricKeys.forEach(function(metricKey) {
    result[metricKey] = FcstSnapshot_diffBreakdown_(currentMetric && currentMetric[metricKey], prevMetric && prevMetric[metricKey]);
  });
  return result;
}

function FcstSnapshot_diffBreakdown_(currentValue, prevValue) {
  var cur = currentValue || {};
  var prev = prevValue || {};
  return {
    net: (Number(cur.net) || 0) - (Number(prev.net) || 0),
    newExp: (Number(cur.newExp) || 0) - (Number(prev.newExp) || 0),
    churn: (Number(cur.churn) || 0) - (Number(prev.churn) || 0),
  };
}
