function FcstSnapshot_create(deptKey, members, notesMap, periodKeys) {
  var sheet = getSharedSheet(FCST_SNAPSHOT_SHEET_NAME);
  if (!sheet) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    sheet = ss.insertSheet(FCST_SNAPSHOT_SHEET_NAME);
    sheet.getRange(1, 1, 1, 4).setValues([['\u65e5\u6642', '\u62c5\u5f53\u8005', '\u671f\u9593', '\u30c7\u30fc\u30bf']]);
  }
  var now = new Date();
  var periods = periodKeys || [];
  var prevMetricMap = FcstSnapshot_getLatestMetricMap_(deptKey);
  var notes = notesMap || {};
  var rows = [];

  var metricKeys = ['fcstAdjusted', 'fcstCommit', 'confirmed', 'expectedMrr'];
  if (DEPT_CONFIG[deptKey].features.proposalProducts) {
    metricKeys = metricKeys.concat(['received', 'debtMgmt', 'debtMgmtLite', 'expense']);
  }

  (members || []).forEach(function(member) {
    if (!member || !member.name) return;
    periods.forEach(function(period) {
      var metric = member[period] || {};
      var mapKey = member.name + '|' + period;
      var payload = {};
      Object.keys(metric).forEach(function(k) {
        if (metricKeys.indexOf(k) !== -1 || k === 'target' || k === 'fcstMax' || k === 'keyDeals') {
          payload[k] = metric[k];
        }
      });
      payload.weekOverWeek = FcstSnapshot_buildWeekOverWeek_(metric, prevMetricMap[mapKey] || {}, metricKeys);
      payload.note = String(notes[mapKey] || '');
      payload.__meta = {
        isTotal: !!member.isTotal,
        group: member.group || '',
        dept: member.dept || '',
        name: member.name,
      };
      var timestampKey = deptKey + ':' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
      rows.push([now, deptKey + ':' + member.name, period, JSON.stringify(payload)]);
    });
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  }
  FcstSnapshot_trimOld_(deptKey, sheet);
  return { ok: true, count: rows.length };
}

function FcstSnapshot_getLatestMetricMap_(deptKey) {
  var sheet = getSharedSheet(FCST_SNAPSHOT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 1) return {};
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange(1, 1, lastRow, 4).getValues();
  var latestKey = FcstSnapshot_getLatestTimestampKey_(deptKey, values);
  if (!latestKey) return {};
  var metricMap = {};
  values.forEach(function(row) {
    var d = row[0];
    if (!(d instanceof Date) || isNaN(d)) return;
    var nameRaw = String(row[1] || '').trim();
    if (!nameRaw.startsWith(deptKey + ':')) return;
    var name = nameRaw.slice(deptKey.length + 1);
    if (Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') !== latestKey) return;
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

function FcstSnapshot_trimOld_(deptKey, sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return;
  var allValues = sheet.getRange(1, 1, lastRow, 2).getValues();
  var seenDates = [];
  allValues.forEach(function(row) {
    var d = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) return;
    if (!nameRaw.startsWith(deptKey + ':')) return;
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    if (seenDates.indexOf(key) === -1) seenDates.push(key);
  });
  if (seenDates.length <= 52) return;
  var oldKey = seenDates[0];
  for (var i = lastRow; i >= 1; i--) {
    var d = allValues[i - 1][0];
    var nameRaw = String(allValues[i - 1][1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) continue;
    if (!nameRaw.startsWith(deptKey + ':')) continue;
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    if (key === oldKey) sheet.deleteRow(i);
  }
}

function FcstSnapshot_getWeekOverWeek(deptKey) {
  var sheet = getSharedSheet(FCST_SNAPSHOT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 1) return {};
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange(1, 1, lastRow, 4).getValues();
  var dateKeys = [];
  values.forEach(function(row) {
    var d = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) return;
    if (!nameRaw.startsWith(deptKey + ':')) return;
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
    var nameRaw = String(row[1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) return;
    if (!nameRaw.startsWith(deptKey + ':')) return;
    var name = nameRaw.slice(deptKey.length + 1);
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
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
    result[mapKey] = FcstSnapshot_buildWeekOverWeek_(latest, prev, null);
  });
  return result;
}

// 最新スナップショットの members データを返す（live view の先週比計算用）
function FcstSnapshot_getLatestMembers(deptKey) {
  var sheet = getSharedSheet(FCST_SNAPSHOT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 1) return null;
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange(1, 1, lastRow, 4).getValues();
  var latestKey = FcstSnapshot_getLatestTimestampKey_(deptKey, values);
  if (!latestKey) return null;
  var data = FcstSnapshot_getDataByTimestampKey_(deptKey, latestKey, values);
  return { members: data.members, date: data.date, periodOptions: data.periodOptions };
}

function FcstSnapshot_getSnapshotDates(deptKey) {
  var sheet = getSharedSheet(FCST_SNAPSHOT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 1) return [];
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange(1, 1, lastRow, 2).getValues();
  var seen = {};
  var dates = [];
  values.forEach(function(row) {
    var d = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) return;
    if (!nameRaw.startsWith(deptKey + ':')) return;
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
    if (!seen[key]) {
      seen[key] = true;
      dates.push(key);
    }
  });
  dates.sort(function(a, b) { return a > b ? -1 : a < b ? 1 : 0; });
  return dates;
}

function FcstSnapshot_getDataByDate(deptKey, dateStr) {
  var sheet = getSharedSheet(FCST_SNAPSHOT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 1) return { members: [], fcstAdjusted: {}, weekOverWeekMap: {}, date: dateStr, periodOptions: [] };
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange(1, 1, lastRow, 4).getValues();
  var latestKeyForDate = '';
  values.forEach(function(row) {
    var d = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) return;
    if (!nameRaw.startsWith(deptKey + ':')) return;
    if (Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') !== dateStr) return;
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    if (!latestKeyForDate || key > latestKeyForDate) latestKeyForDate = key;
  });
  if (!latestKeyForDate) return { members: [], fcstAdjusted: {}, weekOverWeekMap: {}, notes: {}, date: dateStr, periodOptions: [] };
  return FcstSnapshot_getDataByTimestampKey_(deptKey, latestKeyForDate, values);
}

function FcstSnapshot_getDataByTimestampKey_(deptKey, timestampKey, valuesOpt) {
  var rows = valuesOpt;
  if (!rows) {
    var sheet = getSharedSheet(FCST_SNAPSHOT_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 1) return { members: [], fcstAdjusted: {}, weekOverWeekMap: {}, notes: {}, date: '', periodOptions: [] };
    rows = sheet.getRange(1, 1, sheet.getLastRow(), 4).getValues();
  }
  var memberMap = {};
  var fcstAdjusted = {};
  var weekOverWeekMap = {};
  var notes = {};
  var monthKeyMap = {};
  rows.forEach(function(row) {
    var d = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) return;
    if (!nameRaw.startsWith(deptKey + ':')) return;
    var name = nameRaw.slice(deptKey.length + 1);
    if (Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') !== timestampKey) return;
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
        sortOrder: 0
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
    if (/^\d{4}-\d{2}$/.test(period)) monthKeyMap[period] = true;
  });

  var members = Object.keys(memberMap).map(function(n) { return memberMap[n]; });
  var periodOptions = FcstPeriods_buildDefinitionsFromMonthKeys_(Object.keys(monthKeyMap));
  return {
    members: members,
    fcstAdjusted: fcstAdjusted,
    weekOverWeekMap: weekOverWeekMap,
    notes: notes,
    date: timestampKey.slice(0, 10),
    timestampKey: timestampKey,
    periodOptions: periodOptions
  };
}

function FcstSnapshot_getTrendData(deptKey, periodKey, liveData) {
  var sheet = getSharedSheet(FCST_SNAPSHOT_SHEET_NAME);
  var values = (!sheet || sheet.getLastRow() < 1)
    ? []
    : sheet.getRange(1, 1, sheet.getLastRow(), 4).getValues();
  var targetPeriod = String(periodKey || '').trim();
  var data = { weeks: [], dates: [], series: { target: [], fcstAdjusted: [], fcstCommit: [], confirmed: [], expectedMrr: [] } };
  var snapshotMap = {};

  values.forEach(function(row) {
    var d = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) return;
    if (!nameRaw.startsWith(deptKey + ':')) return;
    var period = String(row[2] || '').trim();
    if (period !== targetPeriod) return;

    var payload;
    try {
      payload = JSON.parse(String(row[3] || '{}'));
    } catch (e) {
      payload = {};
    }

    var meta = payload.__meta || {};
    if (!meta.isTotal || meta.group !== deptKey) return;
    var timestampKey = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    snapshotMap[timestampKey] = { payload: payload, date: d };
  });

  Object.keys(snapshotMap).sort().forEach(function(timestampKey) {
    var entry = snapshotMap[timestampKey] || {};
    var payload = entry.payload || {};
    data.dates.push(timestampKey.slice(0, 10));
    data.weeks.push(Utilities.formatDate(entry.date, 'Asia/Tokyo', 'M/d'));
    data.series.target.push(FcstSnapshot_getTrendMetricNet_(payload.target));
    data.series.fcstAdjusted.push(FcstSnapshot_getTrendMetricNet_(payload.fcstAdjusted));
    data.series.fcstCommit.push(FcstSnapshot_getTrendMetricNet_(payload.fcstCommit));
    data.series.confirmed.push(FcstSnapshot_getTrendMetricNet_(payload.confirmed));
    data.series.expectedMrr.push(FcstSnapshot_getTrendMetricNet_(payload.expectedMrr));
  });

  var livePeriod = FcstSnapshot_findTrendLivePeriod_(liveData, targetPeriod);
  var liveMetric = FcstSnapshot_findTrendLiveMetric_(liveData, livePeriod, deptKey);
  if (liveMetric) {
    data.dates.push('live');
    data.weeks.push('現在');
    data.series.target.push(FcstSnapshot_getTrendMetricNet_(liveMetric.target));
    data.series.fcstAdjusted.push(FcstSnapshot_getTrendMetricNet_(liveMetric.fcstAdjusted));
    data.series.fcstCommit.push(FcstSnapshot_getTrendMetricNet_(liveMetric.fcstCommit));
    data.series.confirmed.push(FcstSnapshot_getTrendMetricNet_(liveMetric.confirmed));
    data.series.expectedMrr.push(FcstSnapshot_getTrendMetricNet_(liveMetric.expectedMrr));
  }

  return data;
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

function FcstSnapshot_getLatestTimestampKey_(deptKey, values) {
  var latestKey = '';
  (values || []).forEach(function(row) {
    var d = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) return;
    if (deptKey && !nameRaw.startsWith(deptKey + ':')) return;
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    if (!latestKey || key > latestKey) latestKey = key;
  });
  return latestKey;
}

function FcstSnapshot_findTrendLivePeriod_(liveData, targetPeriod) {
  var period = String(targetPeriod || '').trim();
  var periodOptions = (liveData && liveData.periodOptions) || [];
  if (period && FcstPeriods_expandKeys_(periodOptions).indexOf(period) !== -1) return period;
  return periodOptions.length ? periodOptions[0].key : '';
}

function FcstSnapshot_findTrendLiveMetric_(liveData, periodKey, deptKey) {
  var members = (liveData && liveData.members) || [];
  var total = members.find(function(member) {
    return member && member.isTotal && member.group === deptKey;
  });
  return total && periodKey ? total[periodKey] : null;
}

function FcstSnapshot_getTrendMetricNet_(value) {
  if (typeof value === 'number') return Number(value) || 0;
  return Number(value && value.net) || 0;
}

function FcstSnapshot_extractMetricPayload_(payload) {
  var metric = {};
  Object.keys(payload || {}).forEach(function(k) {
    if (k === '__meta' || k === 'weekOverWeek' || k === 'note') return;
    metric[k] = payload[k];
  });
  return metric;
}

function FcstSnapshot_buildWeekOverWeek_(currentMetric, prevMetric, metricKeysOpt) {
  var metricKeys = metricKeysOpt || ['fcstAdjusted', 'fcstCommit', 'confirmed', 'expectedMrr'];
  ['received', 'debtMgmt', 'debtMgmtLite', 'expense'].forEach(function(metricKey) {
    if ((currentMetric && currentMetric.hasOwnProperty(metricKey)) || (prevMetric && prevMetric.hasOwnProperty(metricKey))) {
      metricKeys.push(metricKey);
    }
  });
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
