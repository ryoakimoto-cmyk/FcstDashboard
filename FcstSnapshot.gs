function FcstSnapshot_create(deptKey, members, notesMap, periodKeys) {
  return FcstSnapshot_createAt_(deptKey, members, notesMap, periodKeys, new Date(), { captureMode: 'scheduled' });
}

function FcstSnapshot_createAt_(deptKey, members, notesMap, periodKeys, snapshotAt, meta) {
  var sheet = getSharedSheet(FCST_SNAPSHOT_SHEET_NAME);
  if (!sheet) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    sheet = ss.insertSheet(FCST_SNAPSHOT_SHEET_NAME);
    sheet.getRange(1, 1, 1, 4).setValues([['\u65e5\u6642', '\u62c5\u5f53\u8005', '\u671f\u9593', '\u30c7\u30fc\u30bf']]);
  }

  var snapshotDate = (snapshotAt instanceof Date && !isNaN(snapshotAt)) ? new Date(snapshotAt.getTime()) : new Date();
  var options = meta || {};
  var periods = periodKeys || [];
  var existingValues = sheet.getLastRow() > 0 ? sheet.getRange(1, 1, sheet.getLastRow(), 4).getValues() : [];
  var timestampKey = Utilities.formatDate(snapshotDate, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  var captureMode = FcstSnapshot_normalizeCaptureMode_(options.captureMode);

  if (!options.force) {
    for (var p = 0; p < periods.length; p++) {
      if (FcstSnapshot_hasRowAt_(deptKey, timestampKey, periods[p], existingValues)) {
        return { ok: true, skipped: true, count: 0 };
      }
    }
  }

  var prevMetricMap = FcstSnapshot_getLatestMetricMap_(deptKey);
  var notes = notesMap || {};
  var rows = [];
  var metricKeys = ['fcstAdjusted', 'fcstCommit', 'fcstMin', 'fcstMax', 'confirmed', 'expectedMrr'];

  if (isProposalProductsEnabled_(deptKey)) {
    metricKeys = metricKeys.concat(['received', 'debtMgmt', 'debtMgmtLite', 'expense']);
  }

  (members || []).forEach(function(member) {
    if (!member || !member.name) return;
    periods.forEach(function(period) {
      var metric = member[period] || {};
      var mapKey = member.name + '|' + period;
      var payload = {};

      Object.keys(metric).forEach(function(k) {
        if (metricKeys.indexOf(k) !== -1 || k === 'target' || k === 'keyDeals') {
          payload[k] = metric[k];
        }
      });

      payload.weekOverWeek = FcstSnapshot_buildWeekOverWeek_(metric, prevMetricMap[mapKey] || {}, metricKeys);
      payload.note = String(notes[mapKey] || '');
      payload.__meta = {
        isTotal: !!member.isTotal,
        group: member.group || '',
        groupCode: member.groupCode || '',
        dept: member.dept || '',
        totalKind: member.totalKind || '',
        captureMode: captureMode,
        name: member.name
      };
      if (options.backfilled) payload.__meta.backfilled = true;

      rows.push([snapshotDate, deptKey + ':' + member.name, period, JSON.stringify(payload)]);
    });
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  }
  FcstSnapshot_trimOld_(deptKey, sheet);
  return { ok: true, skipped: false, count: rows.length };
}

function FcstSnapshot_normalizeCaptureMode_(captureMode) {
  var mode = String(captureMode || '').trim();
  if (mode === 'scheduled' || mode === 'auto-recovery' || mode === 'manual-backfill') return mode;
  return 'scheduled';
}

function FcstSnapshot_buildSnapshotInputFromLive_(liveData) {
  var members = (liveData && liveData.members) || [];
  var notesMap = (liveData && liveData.notes) || {};
  var periodKeys = [];

  if (liveData && liveData.periodOptions && liveData.periodOptions.length && typeof FcstPeriods_expandKeys_ === 'function') {
    periodKeys = FcstPeriods_expandKeys_(liveData.periodOptions);
  }
  if (!periodKeys.length) periodKeys = FcstSnapshot_collectPeriodKeysFromMembers_(members);

  return {
    members: members,
    notesMap: notesMap,
    periodKeys: periodKeys
  };
}

function FcstSnapshot_collectPeriodKeysFromMembers_(members) {
  var seen = {};
  var keys = [];

  (members || []).forEach(function(member) {
    Object.keys(member || {}).forEach(function(key) {
      var metric = member[key];
      if (!metric || typeof metric !== 'object' || Array.isArray(metric)) return;
      if (!Object.prototype.hasOwnProperty.call(metric, 'target') &&
          !Object.prototype.hasOwnProperty.call(metric, 'fcstAdjusted') &&
          !Object.prototype.hasOwnProperty.call(metric, 'fcstCommit') &&
          !Object.prototype.hasOwnProperty.call(metric, 'confirmed') &&
          !Object.prototype.hasOwnProperty.call(metric, 'expectedMrr')) {
        return;
      }
      if (seen[key]) return;
      seen[key] = true;
      keys.push(key);
    });
  });

  return keys.sort(function(a, b) { return a < b ? -1 : a > b ? 1 : 0; });
}

function FcstSnapshot_getMondayAt3AM_(date) {
  var baseDate = (date instanceof Date && !isNaN(date)) ? date : new Date();
  var dateKey = Utilities.formatDate(baseDate, 'Asia/Tokyo', 'yyyy-MM-dd');
  var monday = new Date(dateKey + 'T03:00:00+09:00');
  var dayOfWeek = Number(Utilities.formatDate(baseDate, 'Asia/Tokyo', 'u')) || 1;
  monday.setTime(monday.getTime() - (dayOfWeek - 1) * 86400000);
  return monday;
}

function FcstSnapshot_hasRowAt_(deptKey, timestampKey, periodKey, valuesOpt) {
  var values = valuesOpt || [];
  return values.some(function(row) {
    var d = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) return false;
    if (!nameRaw.startsWith(deptKey + ':')) return false;
    if (String(row[2] || '').trim() !== String(periodKey || '').trim()) return false;
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') === timestampKey;
  });
}

function FcstSnapshot_hasDateRow_(deptKey, dateKey, valuesOpt) {
  var values = valuesOpt || [];
  return values.some(function(row) {
    var d = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) return false;
    if (!nameRaw.startsWith(deptKey + ':')) return false;
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') === dateKey;
  });
}

function FcstSnapshot_shouldAutoRecoverWeekly_(now) {
  var baseDate = (now instanceof Date && !isNaN(now)) ? now : new Date();
  return Utilities.formatDate(baseDate, 'Asia/Tokyo', 'u') === '1' &&
    Number(Utilities.formatDate(baseDate, 'Asia/Tokyo', 'H')) >= 3;
}

function FcstSnapshot_autoRecoverWeeklyIfMissing_(deptKey) {
  if (!deptKey) return;

  var monday = FcstSnapshot_getMondayAt3AM_(new Date());
  var mondayKey = Utilities.formatDate(monday, 'Asia/Tokyo', 'yyyy-MM-dd');
  var sheet = getSharedSheet(FCST_SNAPSHOT_SHEET_NAME);
  var values = (!sheet || sheet.getLastRow() < 1)
    ? []
    : sheet.getRange(1, 1, sheet.getLastRow(), 4).getValues();

  if (FcstSnapshot_hasDateRow_(deptKey, mondayKey, values)) return;

  var live = AggregatedCache_read(deptKey) || null;
  if (!live) return;

  var input = FcstSnapshot_buildSnapshotInputFromLive_(live);
  FcstSnapshot_createAt_(deptKey, input.members, input.notesMap, input.periodKeys, monday, { captureMode: 'auto-recovery' });
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
  seenDates.sort(function(a, b) { return a < b ? -1 : a > b ? 1 : 0; });
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
        groupCode: meta.groupCode || '',
        dept: meta.dept || '',
        totalKind: meta.totalKind || '',
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

function FcstSnapshot_resolveTrendPeriodKey_(trendBlock, liveData) {
  var token = String(trendBlock || '').trim();
  var periodOptions = (liveData && liveData.periodOptions) || [];
  var defaultOption = periodOptions.length ? periodOptions[0] : null;
  var expandedKeys = FcstPeriods_expandKeys_(periodOptions);
  if (token && expandedKeys.indexOf(token) !== -1) return token;
  if (token === 'Q') return defaultOption && defaultOption.key ? String(defaultOption.key) : '';

  var monthMatch = token.match(/^M(\d+)$/);
  if (monthMatch && defaultOption) {
    var requestedMonth = Number(monthMatch[1]);
    var months = defaultOption.months || [];
    for (var i = 0; i < months.length; i++) {
      var parsed = FcstPeriods_parseMonthKey_(months[i]);
      if (parsed && parsed.month === requestedMonth) return months[i];
    }

    var fallbackIndexMap = { M5: 0, M6: 1, M7: 2 };
    if (Object.prototype.hasOwnProperty.call(fallbackIndexMap, token)) {
      return months[fallbackIndexMap[token]] || '';
    }
  }

  return defaultOption && defaultOption.key ? String(defaultOption.key) : '';
}

function FcstSnapshot_runCreateSnapshot_forDate(dateStr, deptKey, force) {
  var targetDate = new Date(String(dateStr || '') + 'T00:00:00+09:00');
  if (isNaN(targetDate)) throw new Error('\u65e5\u4ed8\u5f62\u5f0f\u304c\u4e0d\u6b63\u3067\u3059');

  var snapshotAt = FcstSnapshot_getMondayAt3AM_(targetDate);
  var deptKeys = deptKey ? [deptKey] : getDeptKeys_();
  var results = [];

  deptKeys.forEach(function(dk) {
    var live = AggregatedCache_read(dk);
    if (!live) live = AggregatedCache_refresh(dk);
    var input = FcstSnapshot_buildSnapshotInputFromLive_(live);
    var created = FcstSnapshot_createAt_(dk, input.members, input.notesMap, input.periodKeys, snapshotAt, {
      captureMode: 'manual-backfill',
      force: !!force,
      backfilled: true
    });
    results.push({
      deptKey: dk,
      snapshotAt: Utilities.formatDate(snapshotAt, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'),
      skipped: !!(created && created.skipped),
      count: Number(created && created.count) || 0
    });
  });

  return { ok: true, results: results };
}

function FcstSnapshot_runBackfillMissingWeeks(deptKey, weeksBack) {
  var totalWeeks = Number(weeksBack);
  if (!totalWeeks || totalWeeks < 1) totalWeeks = 12;

  var deptKeys = deptKey ? [deptKey] : getDeptKeys_();
  var currentMonday = FcstSnapshot_getMondayAt3AM_(new Date());
  var filled = [];

  deptKeys.forEach(function(dk) {
    var live = AggregatedCache_read(dk);
    if (!live) live = AggregatedCache_refresh(dk);
    var input = FcstSnapshot_buildSnapshotInputFromLive_(live);
    var sheet = getSharedSheet(FCST_SNAPSHOT_SHEET_NAME);
    var values = (!sheet || sheet.getLastRow() < 1)
      ? []
      : sheet.getRange(1, 1, sheet.getLastRow(), 4).getValues();

    for (var i = totalWeeks - 1; i >= 0; i--) {
      var monday = new Date(currentMonday.getTime() - i * 7 * 86400000);
      var mondayKey = Utilities.formatDate(monday, 'Asia/Tokyo', 'yyyy-MM-dd');
      if (FcstSnapshot_hasDateRow_(dk, mondayKey, values)) continue;

      var created = FcstSnapshot_createAt_(dk, input.members, input.notesMap, input.periodKeys, monday, {
        captureMode: 'manual-backfill',
        backfilled: true
      });
      if (created && !created.skipped) {
        Logger.log('filled missing snapshot: ' + dk + ' ' + mondayKey);
        filled.push({ deptKey: dk, date: mondayKey, count: Number(created.count) || 0 });
      }
    }
  });

  return { ok: true, filled: filled, totalFilled: filled.length };
}

function FcstSnapshot_getTrendData(deptKey, periodKey, liveData) {
  var sheet = getSharedSheet(FCST_SNAPSHOT_SHEET_NAME);
  var values = (!sheet || sheet.getLastRow() < 1)
    ? []
    : sheet.getRange(1, 1, sheet.getLastRow(), 4).getValues();
  var targetPeriod = FcstSnapshot_resolveTrendPeriodKey_(periodKey, liveData) || String(periodKey || '').trim();
  var data = {
    labels: [],
    weeks: [],
    dates: [],
    series: { target: [], fcstAdjusted: [], fcstCommit: [], confirmed: [], expectedMrr: [] },
    points: []
  };
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
    if (!meta.isTotal || meta.totalKind !== 'department' || meta.dept !== deptKey) return;
    var timestampKey = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    snapshotMap[timestampKey] = { payload: payload, date: d };
  });

  Object.keys(snapshotMap).sort().forEach(function(timestampKey) {
    var entry = snapshotMap[timestampKey] || {};
    var payload = entry.payload || {};
    var label = Utilities.formatDate(entry.date, 'Asia/Tokyo', 'M/d');
    var dateKey = timestampKey.slice(0, 10);
    var metrics = FcstSnapshot_extractTrendMetrics_(payload);
    var keyDeals = Array.isArray(payload.keyDeals) ? payload.keyDeals : [];

    data.dates.push(dateKey);
    data.weeks.push(label);
    data.series.target.push(metrics.target);
    data.series.fcstAdjusted.push(metrics.fcstAdjusted);
    data.series.fcstCommit.push(metrics.fcstCommit);
    data.series.confirmed.push(metrics.confirmed);
    data.series.expectedMrr.push(metrics.expectedMrr);
    data.points.push({
      snapshotKey: timestampKey,
      date: dateKey,
      label: label,
      isLive: false,
      isBackfilled: !!(payload.__meta && payload.__meta.backfilled),
      metrics: metrics,
      keyDealCount: keyDeals.length,
      keyDealPreview: FcstSnapshot_extractKeyDealPreview_(keyDeals)
    });
  });

  var livePeriod = FcstSnapshot_findTrendLivePeriod_(liveData, targetPeriod);
  var liveMetric = FcstSnapshot_findTrendLiveMetric_(liveData, livePeriod, deptKey);
  if (liveMetric) {
    var liveMetrics = FcstSnapshot_extractTrendMetrics_(liveMetric);
    var liveKeyDeals = Array.isArray(liveMetric.keyDeals) ? liveMetric.keyDeals : [];
    data.dates.push('live');
    data.weeks.push('\u73fe\u5728');
    data.series.target.push(liveMetrics.target);
    data.series.fcstAdjusted.push(liveMetrics.fcstAdjusted);
    data.series.fcstCommit.push(liveMetrics.fcstCommit);
    data.series.confirmed.push(liveMetrics.confirmed);
    data.series.expectedMrr.push(liveMetrics.expectedMrr);
    data.points.push({
      snapshotKey: 'live',
      date: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'),
      label: '\u73fe\u5728',
      isLive: true,
      isBackfilled: false,
      metrics: liveMetrics,
      keyDealCount: liveKeyDeals.length,
      keyDealPreview: FcstSnapshot_extractKeyDealPreview_(liveKeyDeals)
    });
  }

  data.labels = data.weeks.slice();
  return data;
}

function FcstSnapshot_getTrendWeekDetails(deptKey, periodKey, snapshotKey) {
  var liveData = AggregatedCache_read(deptKey) || null;
  var resolverData = liveData || FcstSnapshot_getLatestMembers(deptKey) || { periodOptions: [] };
  var targetPeriod = FcstSnapshot_resolveTrendPeriodKey_(periodKey, resolverData) || String(periodKey || '').trim();
  var result = {
    snapshotKey: snapshotKey,
    isLive: snapshotKey === 'live',
    metrics: FcstSnapshot_extractTrendMetrics_(null),
    keyDeals: []
  };

  if (snapshotKey === 'live') {
    var livePeriod = FcstSnapshot_findTrendLivePeriod_(liveData, targetPeriod);
    var liveMetric = FcstSnapshot_findTrendLiveMetric_(liveData, livePeriod, deptKey);
    result.metrics = FcstSnapshot_extractTrendMetrics_(liveMetric);
    result.keyDeals = FcstSnapshot_normalizeKeyDeals_(liveMetric && liveMetric.keyDeals);
    return result;
  }

  if (!targetPeriod) return result;

  var sheet = getSharedSheet(FCST_SNAPSHOT_SHEET_NAME);
  var values = (!sheet || sheet.getLastRow() < 1)
    ? []
    : sheet.getRange(1, 1, sheet.getLastRow(), 4).getValues();

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var d = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) continue;
    if (!nameRaw.startsWith(deptKey + ':')) continue;
    if (String(row[2] || '').trim() !== targetPeriod) continue;
    if (Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') !== snapshotKey) continue;

    var payload;
    try {
      payload = JSON.parse(String(row[3] || '{}'));
    } catch (e) {
      payload = {};
    }

    var meta = payload.__meta || {};
    if (!meta.isTotal || meta.totalKind !== 'department' || meta.dept !== deptKey) continue;
    return {
      snapshotKey: snapshotKey,
      isLive: false,
      metrics: FcstSnapshot_extractTrendMetrics_(payload),
      keyDeals: FcstSnapshot_normalizeKeyDeals_(payload.keyDeals)
    };
  }

  return result;
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
    return member && member.isTotal && member.totalKind === 'department' && member.dept === deptKey;
  });
  return total && periodKey ? total[periodKey] : null;
}

function FcstSnapshot_getTrendMetricNet_(value) {
  if (typeof value === 'number') return Number(value) || 0;
  return Number(value && value.net) || 0;
}

function FcstSnapshot_extractTrendMetrics_(metric) {
  return {
    target: FcstSnapshot_getTrendMetricNet_(metric && metric.target),
    confirmed: FcstSnapshot_getTrendMetricNet_(metric && metric.confirmed),
    fcstAdjusted: FcstSnapshot_getTrendMetricNet_(metric && metric.fcstAdjusted),
    fcstCommit: FcstSnapshot_getTrendMetricNet_(metric && metric.fcstCommit),
    expectedMrr: FcstSnapshot_getTrendMetricNet_(metric && metric.expectedMrr)
  };
}

function FcstSnapshot_normalizeKeyDeal_(rawKeyDeal) {
  var deal = rawKeyDeal || {};
  var company = String(deal.company || deal.companyName || deal.accountName || deal.name || '').trim();
  var monthlyMrr = 0;

  ['monthlyMrr', 'mrr', 'amount'].some(function(key) {
    if (!Object.prototype.hasOwnProperty.call(deal, key)) return false;
    var rawValue = deal[key];
    if (rawValue === null || rawValue === undefined || String(rawValue).trim() === '') return false;
    var value = Number(rawValue);
    if (isNaN(value)) return false;
    monthlyMrr = value;
    return true;
  });

  return {
    company: company,
    monthlyMrr: monthlyMrr,
    phase: String(deal.phase || '').trim(),
    fcst: Number(deal.fcst) || 0,
    oppId: String(deal.oppId || '').trim()
  };
}

function FcstSnapshot_normalizeKeyDeals_(keyDeals) {
  return (Array.isArray(keyDeals) ? keyDeals : []).map(function(keyDeal) {
    return FcstSnapshot_normalizeKeyDeal_(keyDeal);
  });
}

function FcstSnapshot_extractKeyDealPreview_(keyDeals) {
  return FcstSnapshot_normalizeKeyDeals_(keyDeals)
    .sort(function(a, b) { return b.monthlyMrr - a.monthlyMrr; })
    .slice(0, 3)
    .map(function(keyDeal) {
      return {
        company: keyDeal.company,
        monthlyMrr: keyDeal.monthlyMrr
      };
    });
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
  var metricKeys = metricKeysOpt || ['fcstAdjusted', 'fcstCommit', 'fcstMin', 'fcstMax', 'confirmed', 'expectedMrr'];
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
  if (typeof currentValue === 'number' || typeof prevValue === 'number') {
    return {
      net: (Number(currentValue) || 0) - (Number(prevValue) || 0),
      newExp: 0,
      churn: 0
    };
  }
  var cur = currentValue || {};
  var prev = prevValue || {};
  return {
    net: (Number(cur.net) || 0) - (Number(prev.net) || 0),
    newExp: (Number(cur.newExp) || 0) - (Number(prev.newExp) || 0),
    churn: (Number(cur.churn) || 0) - (Number(prev.churn) || 0)
  };
}
