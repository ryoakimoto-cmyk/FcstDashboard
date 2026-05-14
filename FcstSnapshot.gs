function FcstSnapshot_create(deptKey, members, notesMap, periodKeys) {
  return FcstSnapshot_createAt_(deptKey, members, notesMap, periodKeys, new Date(), { captureMode: 'scheduled' });
}

function FcstSnapshot_createAt_(deptKey, members, notesMap, periodKeys, snapshotAt, meta) {
  var snapshotDate = (snapshotAt instanceof Date && !isNaN(snapshotAt)) ? new Date(snapshotAt.getTime()) : new Date();
  var options = meta || {};
  var periods = FcstSnapshot_filterSnapshotPeriodKeys_(periodKeys || []);
  var dateKey = Utilities.formatDate(snapshotDate, 'Asia/Tokyo', 'yyyy-MM-dd');
  var timestampKey = Utilities.formatDate(snapshotDate, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  var captureMode = FcstSnapshot_normalizeCaptureMode_(options.captureMode);

  var prevMetricMap = FcstSnapshot_getLatestMetricMap_(deptKey, dateKey);
  var notes = notesMap || {};
  var rows = [];
  var metricKeys = ['fcstAdjusted', 'fcstCommit', 'fcstMin', 'fcstMax', 'confirmed', 'expectedMrr'];

  if (isProposalProductsEnabled_(deptKey)) {
    metricKeys = metricKeys.concat(PROPOSAL_PRODUCT_METRIC_KEYS, CONFIRMED_PROPOSAL_PRODUCT_METRIC_KEYS);
  }

  (members || []).forEach(function(member) {
    if (!member || !member.name) return;
    periods.forEach(function(period) {
      var metric = member[period] || {};
      var mapKey = member.name + '|' + period;
      var payload = {};

      Object.keys(metric).forEach(function(k) {
        if (metricKeys.indexOf(k) !== -1 || k === 'target') {
          payload[k] = metric[k];
        }
      });

      payload.weekOverWeek = FcstSnapshot_buildWeekOverWeek_(metric, prevMetricMap[mapKey] || {}, metricKeys);
      payload.note = String(notes[mapKey] || '');
      payload.__meta = FcstSnapshot_buildPayloadMeta_(member, options);

      rows.push([snapshotDate, deptKey + ':' + member.name, period, JSON.stringify(payload)]);
    });
  });

  if (rows.length) {
    FcstSnapshot_getSheets_().forEach(function(sheet) {
      FcstSnapshot_deleteByDate_(deptKey, sheet, dateKey);
    });
    var writeResult = SnapshotStorage_appendRows_(FCST_SNAPSHOT_SHEET_NAME, FcstSnapshot_headers_(), rows);
    FcstSnapshot_trimOld_(deptKey, writeResult.sheet);
    return {
      ok: true,
      skipped: false,
      count: rows.length,
      date: dateKey,
      snapshotAt: timestampKey,
      captureMode: captureMode,
      storageFileId: writeResult.fileId,
      storageFileUrl: writeResult.fileUrl,
      storageSheetName: writeResult.sheetName,
      storageRolledOver: !!writeResult.rolledOver
    };
  }
  return { ok: true, skipped: false, count: rows.length, date: dateKey, snapshotAt: timestampKey, captureMode: captureMode };
}

function FcstSnapshot_headers_() {
  return ['日時', '担当者', '期間', 'データ'];
}

function FcstSnapshot_filterSnapshotPeriodKeys_(periodKeys) {
  var seen = {};
  var result = [];
  (periodKeys || []).forEach(function(periodKey) {
    var key = String(periodKey || '').trim();
    if (!FcstPeriods_parseMonthKey_(key) || seen[key]) return;
    seen[key] = true;
    result.push(key);
  });
  return result;
}

function FcstSnapshot_getSheets_() {
  return SnapshotStorage_getReadSheets_(FCST_SNAPSHOT_SHEET_NAME, FcstSnapshot_headers_());
}

function FcstSnapshot_getAllValues_(columnCount) {
  return SnapshotStorage_getAllValues_(FCST_SNAPSHOT_SHEET_NAME, FcstSnapshot_headers_(), columnCount || 4);
}

function FcstSnapshot_normalizeCaptureMode_(captureMode) {
  var mode = String(captureMode || '').trim();
  if (mode === 'scheduled' || mode === 'auto-recovery' || mode === 'manual-backfill') return mode;
  return 'scheduled';
}

function FcstSnapshot_buildPayloadMeta_(member, options) {
  var meta = {};
  var totalKind = String(member && member.totalKind || '').trim();
  if (!totalKind) {
    totalKind = member && member.isTotal ? SHARED_TOTAL_KIND.GROUP : SHARED_TOTAL_KIND.INDIVIDUAL;
  }
  meta.totalKind = totalKind;

  var group = String(member && member.group || '').trim();
  if (group) meta.group = group;

  if (options && options.backfilled) meta.backfilled = true;
  return meta;
}

function FcstSnapshot_parseRowName_(nameRaw) {
  var text = String(nameRaw || '').trim();
  var idx = text.indexOf(':');
  if (idx < 0) return { deptKey: '', name: text };
  return {
    deptKey: text.slice(0, idx),
    name: text.slice(idx + 1)
  };
}

function FcstSnapshot_normalizeMeta_(payload, rowDeptKey, rowName) {
  var raw = payload && payload.__meta || {};
  var name = String(rowName || raw.name || '').trim();
  var totalKind = String(raw.totalKind || '').trim();
  var legacyIsTotal = raw.isTotal === true;

  if (!totalKind) {
    if (legacyIsTotal) {
      totalKind = (name === 'department-total' || String(raw.group || '') === SHARED_ALL_GROUP_LABEL)
        ? SHARED_TOTAL_KIND.DEPARTMENT
        : SHARED_TOTAL_KIND.GROUP;
    } else {
      totalKind = SHARED_TOTAL_KIND.INDIVIDUAL;
    }
  }

  var isTotal = totalKind !== SHARED_TOTAL_KIND.INDIVIDUAL;
  var dept = String(raw.dept || rowDeptKey || '').trim();
  var group = String(raw.group || '').trim();
  var groupCode = String(raw.groupCode || '').trim();

  if (!group && totalKind === SHARED_TOTAL_KIND.DEPARTMENT) group = SHARED_ALL_GROUP_LABEL;
  if (!group && totalKind === SHARED_TOTAL_KIND.GROUP) group = name;
  if (!groupCode && totalKind === SHARED_TOTAL_KIND.DEPARTMENT) groupCode = dept;
  if (!groupCode) groupCode = group;

  return {
    isTotal: isTotal,
    totalKind: totalKind,
    dept: dept,
    group: group,
    groupCode: groupCode,
    backfilled: raw.backfilled === true
  };
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
    members: FcstSnapshot_mergeAdjustedIntoMembers_(members, liveData && liveData.fcstAdjusted, periodKeys),
    notesMap: notesMap,
    periodKeys: periodKeys
  };
}

function FcstSnapshot_mergeAdjustedIntoMembers_(members, adjustedMap, periodKeys) {
  var adjusted = adjustedMap || {};
  return (members || []).map(function(member) {
    var cloned = {};
    Object.keys(member || {}).forEach(function(key) {
      cloned[key] = member[key];
    });
    (periodKeys || []).forEach(function(periodKey) {
      if (!cloned[periodKey]) return;
      var metric = {};
      Object.keys(cloned[periodKey] || {}).forEach(function(metricKey) {
        metric[metricKey] = cloned[periodKey][metricKey];
      });
      var mapKey = String(member.name || '') + '|' + String(periodKey || '');
      metric.fcstAdjusted = adjusted[mapKey] || { net: 0, newExp: 0, churn: 0 };
      cloned[periodKey] = metric;
    });
    return cloned;
  });
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
  var values = FcstSnapshot_getAllValues_(4);

  if (FcstSnapshot_hasDateRow_(deptKey, mondayKey, values)) return;

  var live = AggregatedCache_read(deptKey) || null;
  if (!live) return;

  var input = FcstSnapshot_buildSnapshotInputFromLive_(live);
  FcstSnapshot_createAt_(deptKey, input.members, input.notesMap, input.periodKeys, monday, { captureMode: 'auto-recovery' });
}

function FcstSnapshot_getLatestMetricMap_(deptKey, beforeDateKey) {
  var values = FcstSnapshot_getAllValues_(4);
  if (!values.length) return {};
  var latestKey = FcstSnapshot_getLatestDateKey_(deptKey, values, beforeDateKey);
  if (!latestKey) return {};
  var metricMap = {};
  values.slice().sort(function(a, b) {
    var ad = a && a[0] instanceof Date ? a[0].getTime() : 0;
    var bd = b && b[0] instanceof Date ? b[0].getTime() : 0;
    return ad - bd;
  }).forEach(function(row) {
    var d = row[0];
    if (!(d instanceof Date) || isNaN(d)) return;
    var nameRaw = String(row[1] || '').trim();
    if (!nameRaw.startsWith(deptKey + ':')) return;
    var name = nameRaw.slice(deptKey.length + 1);
    if (Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') !== latestKey) return;
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
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
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
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
    if (key === oldKey) sheet.deleteRow(i);
  }
}

function FcstSnapshot_deleteByDate_(deptKey, sheet, dateKey) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return;
  var allValues = sheet.getRange(1, 1, lastRow, 2).getValues();
  var deleteRows = [];
  for (var i = 0; i < allValues.length; i++) {
    var d = allValues[i][0];
    var nameRaw = String(allValues[i][1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) continue;
    if (!nameRaw.startsWith(deptKey + ':')) continue;
    if (Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') === dateKey) deleteRows.push(i + 1);
  }
  FcstSnapshot_deleteRows_(sheet, deleteRows);
}

function FcstSnapshot_deleteRows_(sheet, rowNumbers) {
  if (!rowNumbers || !rowNumbers.length) return;
  var sorted = rowNumbers.slice().sort(function(a, b) { return b - a; });
  var start = sorted[0];
  var count = 1;
  for (var i = 1; i < sorted.length; i++) {
    var rowNumber = sorted[i];
    if (rowNumber === start - 1) {
      start = rowNumber;
      count++;
      continue;
    }
    sheet.deleteRows(start, count);
    start = rowNumber;
    count = 1;
  }
  sheet.deleteRows(start, count);
}

function FcstSnapshot_getWeekOverWeek(deptKey, opts) {
  var activePeriodSet = opts && opts.activePeriodKeys ? opts.activePeriodKeys : null;
  var values = FcstSnapshot_getAllValues_(4);
  if (!values.length) return {};
  var dateKeys = [];
  var seenDate = {};
  values.forEach(function(row) {
    var d = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) return;
    if (!nameRaw.startsWith(deptKey + ':')) return;
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
    if (!seenDate[key]) { seenDate[key] = true; dateKeys.push(key); }
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
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
    if (key !== latestKey && key !== prevKey) return;
    var period = String(row[2] || '').trim();
    if (activePeriodSet && !activePeriodSet[period]) return;
    var name = nameRaw.slice(deptKey.length + 1);
    var mapKey = name + '|' + period;
    var metric;
    try {
      metric = JSON.parse(String(row[3] || '{}'));
    } catch (e) {
      metric = {};
    }
    if (key === latestKey) latestMap[mapKey] = metric;
    else prevMap[mapKey] = metric;
  });
  var result = {};
  Object.keys(latestMap).forEach(function(mapKey) {
    var latest = latestMap[mapKey];
    var prev = prevMap[mapKey] || {};
    result[mapKey] = FcstSnapshot_buildWeekOverWeek_(latest, prev, null);
  });
  return result;
}

function FcstSnapshot_getLatestMembers(deptKey, options) {
  var values = FcstSnapshot_getAllValues_(4);
  if (!values.length) return null;
  var latestKey = FcstSnapshot_getLatestTimestampKey_(deptKey, values);
  if (!latestKey) return null;
  var data = FcstSnapshot_getDataByTimestampKey_(deptKey, latestKey, values, options || {});
  return { members: data.members, date: data.date, periodOptions: data.periodOptions };
}

function FcstSnapshot_getSnapshotDates(deptKey, opts) {
  var limit = opts && opts.limit ? Number(opts.limit) : 0;
  var values = FcstSnapshot_getAllValues_(2);
  if (!values.length) return [];
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
  if (limit > 0 && dates.length > limit) dates = dates.slice(0, limit);
  return dates;
}

function FcstSnapshot_getDataByDate(deptKey, dateStr, options) {
  var values = FcstSnapshot_getAllValues_(4);
  if (!values.length) return { members: [], fcstAdjusted: {}, weekOverWeekMap: {}, date: dateStr, periodOptions: [] };
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
  return FcstSnapshot_getDataByTimestampKey_(deptKey, latestKeyForDate, values, options || {});
}

function FcstSnapshot_getDataByTimestampKey_(deptKey, timestampKey, valuesOpt, options) {
  var rows = valuesOpt;
  if (!rows) {
    rows = FcstSnapshot_getAllValues_(4);
    if (!rows.length) return { members: [], fcstAdjusted: {}, weekOverWeekMap: {}, notes: {}, date: '', periodOptions: [] };
  }
  var activePeriodSet = options && options.activePeriodKeys ? options.activePeriodKeys : null;
  var memberMap = {};
  var fcstAdjusted = {};
  var weekOverWeekMap = {};
  var notes = {};
  var monthKeyMap = {};
  rows.forEach(function(row) {
    var d = row[0];
    var nameRaw = String(row[1] || '').trim();
    var rowInfo = FcstSnapshot_parseRowName_(nameRaw);
    if (!(d instanceof Date) || isNaN(d)) return;
    if (rowInfo.deptKey !== deptKey) return;
    var name = rowInfo.name;
    if (Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') !== timestampKey) return;
    var period = String(row[2] || '').trim();
    if (activePeriodSet && !activePeriodSet[period]) return;
    var payload;
    try {
      payload = JSON.parse(String(row[3] || '{}'));
    } catch (e) {
      payload = {};
    }

    if (!memberMap[name]) {
      var meta = FcstSnapshot_normalizeMeta_(payload, rowInfo.deptKey, name);
      memberMap[name] = {
        name: name,
        isTotal: meta.isTotal,
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
  FcstSnapshot_addDerivedQuarterMetrics_(members, periodOptions, fcstAdjusted, weekOverWeekMap, notes);
  if (!options || options.includeKeyDeals !== false) {
    FcstSnapshot_attachSnapshotKeyDealsToData_(deptKey, timestampKey.slice(0, 10), {
      members: members,
      periodOptions: periodOptions
    });
  }
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

function FcstSnapshot_addDerivedQuarterMetrics_(members, periodOptions, fcstAdjusted, weekOverWeekMap, notes) {
  (periodOptions || []).forEach(function(option) {
    var quarterKey = String(option && option.key || '').trim();
    var months = (option && option.months || []).filter(function(monthKey) {
      return !!FcstPeriods_parseMonthKey_(monthKey);
    });
    if (!quarterKey || !months.length) return;

    (members || []).forEach(function(member) {
      if (!member) return;
      if (!member[quarterKey]) {
        var monthMetrics = months.map(function(monthKey) { return member[monthKey]; }).filter(Boolean);
        if (monthMetrics.length) member[quarterKey] = FcstSnapshot_sumMetricList_(monthMetrics);
      }

      if (!member[quarterKey]) return;
      var mapKey = member.name + '|' + quarterKey;
      if (!fcstAdjusted[mapKey]) {
        fcstAdjusted[mapKey] = FcstSnapshot_sumBreakdownList_(months.map(function(monthKey) {
          return fcstAdjusted[member.name + '|' + monthKey];
        }));
      }
      if (!weekOverWeekMap[mapKey]) {
        weekOverWeekMap[mapKey] = FcstSnapshot_sumMetricList_(months.map(function(monthKey) {
          return weekOverWeekMap[member.name + '|' + monthKey];
        }).filter(Boolean));
      }
      if (!Object.prototype.hasOwnProperty.call(notes, mapKey)) notes[mapKey] = '';
    });
  });
}

function FcstSnapshot_sumMetricList_(metrics) {
  var total = {};
  (metrics || []).forEach(function(metric) {
    Object.keys(metric || {}).forEach(function(key) {
      if (key === 'keyDeals') return;
      FcstSnapshot_addMetricValue_(total, key, metric[key]);
    });
  });
  return total;
}

function FcstSnapshot_addMetricValue_(target, key, value) {
  if (value === null || value === undefined || Array.isArray(value)) return;
  if (typeof value === 'number') {
    target[key] = (Number(target[key]) || 0) + value;
    return;
  }
  if (typeof value !== 'object') return;

  if (!target[key] || typeof target[key] !== 'object' || Array.isArray(target[key])) target[key] = {};
  Object.keys(value).forEach(function(part) {
    var n = Number(value[part]);
    if (isNaN(n)) return;
    target[key][part] = (Number(target[key][part]) || 0) + n;
  });
}

function FcstSnapshot_sumBreakdownList_(values) {
  var total = { net: 0, newExp: 0, churn: 0 };
  (values || []).forEach(function(value) {
    if (!value) return;
    if (typeof value === 'number') {
      total.net += Number(value) || 0;
      return;
    }
    total.net += Number(value.net) || 0;
    total.newExp += Number(value.newExp) || 0;
    total.churn += Number(value.churn) || 0;
  });
  return total;
}

function FcstSnapshot_attachCurrentKeyDealsToData_(deptKey, data) {
  return FcstSnapshot_attachOppKeyDealsToData_(data, FcstSnapshot_getCurrentOppKeyDealRows_(deptKey));
}

function FcstSnapshot_attachSnapshotKeyDealsToData_(deptKey, dateKey, data) {
  return FcstSnapshot_attachOppKeyDealsToData_(data, FcstSnapshot_getSnapshotOppKeyDealRows_(deptKey, dateKey));
}

function FcstSnapshot_attachCurrentKeyDealsToRows_(deptKey, periodKey, rows) {
  return FcstSnapshot_attachOppKeyDealsToRows_(rows, periodKey, FcstSnapshot_getCurrentOppKeyDealRows_(deptKey));
}

function FcstSnapshot_attachSnapshotKeyDealsToRows_(deptKey, dateKey, periodKey, rows) {
  return FcstSnapshot_attachOppKeyDealsToRows_(rows, periodKey, FcstSnapshot_getSnapshotOppKeyDealRows_(deptKey, dateKey));
}

function FcstSnapshot_attachOppKeyDealsToData_(data, oppRows) {
  if (!data || !Array.isArray(data.members)) return data;
  var periods = FcstSnapshot_collectPeriodKeysForData_(data);
  periods.forEach(function(periodKey) {
    FcstSnapshot_attachOppKeyDealsForPeriod_(data.members, periodKey, oppRows);
  });
  return data;
}

function FcstSnapshot_attachOppKeyDealsToRows_(rows, periodKey, oppRows) {
  var list = Array.isArray(rows) ? rows : [];
  if (!periodKey) return list;
  var byMember = FcstSnapshot_buildKeyDealsByMemberForPeriod_(list, periodKey, oppRows);
  list.forEach(function(row) {
    if (!row) return;
    var deals = FcstSnapshot_collectKeyDealsForMember_(row, list, byMember);
    row.metric = row.metric || {};
    row.metric.keyDeals = deals;
    row.keyDeals = deals;
  });
  return list;
}

function FcstSnapshot_attachOppKeyDealsForPeriod_(members, periodKey, oppRows) {
  var byMember = FcstSnapshot_buildKeyDealsByMemberForPeriod_(members, periodKey, oppRows);
  (members || []).forEach(function(member) {
    if (!member || !member[periodKey]) return;
    member[periodKey].keyDeals = FcstSnapshot_collectKeyDealsForMember_(member, members, byMember);
  });
}

function FcstSnapshot_collectKeyDealsForMember_(member, allMembers, byMember) {
  if (!member) return [];
  if (!member.isTotal) return (byMember[member.name] || []).slice();

  var deals = [];
  (allMembers || []).forEach(function(row) {
    if (!row || row.isTotal) return;
    if (member.totalKind === SHARED_TOTAL_KIND.GROUP && !FcstSnapshot_sameGroup_(member, row)) return;
    deals = deals.concat(byMember[row.name] || []);
  });
  return FcstSnapshot_sortKeyDeals_(deals);
}

function FcstSnapshot_sameGroup_(totalMember, member) {
  var totalGroupCode = String(totalMember && totalMember.groupCode || '').trim();
  var memberGroupCode = String(member && member.groupCode || '').trim();
  if (totalGroupCode && memberGroupCode) return totalGroupCode === memberGroupCode;
  return String(totalMember && totalMember.group || '').trim() === String(member && member.group || '').trim();
}

function FcstSnapshot_buildKeyDealsByMemberForPeriod_(members, periodKey, oppRows) {
  var owners = {};
  (members || []).forEach(function(member) {
    if (!member || member.isTotal || !member.name) return;
    owners[String(member.name).trim()] = true;
  });

  var byMember = {};
  (oppRows || []).forEach(function(row) {
    if (!row || row.keyDeal !== true) return;
    if (!FcstSnapshot_oppRowMatchesPeriod_(row, periodKey)) return;
    var owner = String(row.subOwner || '').trim();
    if (!owner || !owners[owner]) return;
    if (!byMember[owner]) byMember[owner] = [];
    byMember[owner].push(FcstSnapshot_mapOppRowToKeyDeal_(row));
  });

  Object.keys(byMember).forEach(function(owner) {
    byMember[owner] = FcstSnapshot_sortKeyDeals_(byMember[owner]);
  });
  return byMember;
}

function FcstSnapshot_sortKeyDeals_(deals) {
  return (deals || []).slice().sort(function(a, b) {
    return Math.abs(Number(b.monthlyMrr) || 0) - Math.abs(Number(a.monthlyMrr) || 0);
  });
}

function FcstSnapshot_mapOppRowToKeyDeal_(row) {
  return {
    company: String(row && row.dealName || '').trim(),
    monthlyMrr: Number(row && row.mrr) || 0,
    phase: String(row && row.phase || '').trim(),
    fcst: Number(row && row.fcstCommit) || 0,
    oppId: String(row && row.oppId || '').trim()
  };
}

function FcstSnapshot_oppRowMatchesPeriod_(row, periodKey) {
  var monthKey = FcstSnapshot_normalizeOppMonthKey_(row && row.completedMonth);
  if (!monthKey) return false;
  var monthSet = FcstSnapshot_getPeriodMonthSet_(periodKey);
  return !!monthSet[monthKey];
}

function FcstSnapshot_getPeriodMonthSet_(periodKey) {
  var key = String(periodKey || '').trim();
  var months = {};
  if (FcstPeriods_parseMonthKey_(key)) {
    months[key] = true;
    return months;
  }
  var quarter = FcstPeriods_getQuarterDefinitionByKey_(key);
  (quarter && quarter.months || []).forEach(function(monthKey) {
    months[monthKey] = true;
  });
  return months;
}

function FcstSnapshot_periodMatchesTarget_(periodKey, targetPeriodKey) {
  var period = String(periodKey || '').trim();
  var target = String(targetPeriodKey || '').trim();
  if (!period || !target) return false;
  if (period === target) return true;
  return !!FcstSnapshot_getPeriodMonthSet_(target)[period];
}

function FcstSnapshot_normalizeOppMonthKey_(value) {
  if (value instanceof Date && !isNaN(value)) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM');
  }
  var text = String(value || '').trim();
  if (!text) return '';
  if (/^\d{4}-\d{2}$/.test(text)) return text;
  var compact = text.match(/^(\d{4})(\d{2})$/);
  if (compact) return compact[1] + '-' + compact[2];
  var match = text.match(/^(\d{4})[-\/年](\d{1,2})/);
  if (!match) return '';
  return match[1] + '-' + String(Number(match[2])).padStart(2, '0');
}

function FcstSnapshot_collectPeriodKeysForData_(data) {
  var seen = {};
  var keys = [];
  FcstPeriods_expandKeys_((data && data.periodOptions) || []).forEach(function(periodKey) {
    if (!periodKey || seen[periodKey]) return;
    seen[periodKey] = true;
    keys.push(periodKey);
  });
  (data && data.members || []).forEach(function(member) {
    Object.keys(member || {}).forEach(function(key) {
      if (!member[key] || typeof member[key] !== 'object' || Array.isArray(member[key])) return;
      if (!FcstPeriods_parseMonthKey_(key) && !FcstPeriods_getQuarterDefinitionByKey_(key)) return;
      if (seen[key]) return;
      seen[key] = true;
      keys.push(key);
    });
  });
  return keys;
}

function FcstSnapshot_getSnapshotOppKeyDealRows_(deptKey, dateKey) {
  var rows = [];
  FcstSnapshot_getOppDeptKeysForFcstDept_(deptKey).forEach(function(oppDeptKey) {
    var result = OppListSnapshot_getByDate(oppDeptKey, dateKey) || {};
    rows = rows.concat(result.rows || []);
  });
  return rows.filter(function(row) { return row && row.keyDeal === true; });
}

function FcstSnapshot_getSnapshotKeyDealsForPeriod_(deptKey, dateKey, periodKey) {
  return FcstSnapshot_getSnapshotOppKeyDealRows_(deptKey, dateKey).filter(function(row) {
    return FcstSnapshot_oppRowMatchesPeriod_(row, periodKey);
  }).map(function(row) {
    return FcstSnapshot_mapOppRowToKeyDeal_(row);
  }).sort(function(a, b) {
    return Math.abs(Number(b.monthlyMrr) || 0) - Math.abs(Number(a.monthlyMrr) || 0);
  });
}

function FcstSnapshot_getCurrentOppKeyDealRows_(deptKey) {
  var rows = [];
  FcstSnapshot_getOppDeptKeysForFcstDept_(deptKey).forEach(function(oppDeptKey) {
    var result = AppDataCache_getOpportunities(oppDeptKey);
    if (!result || result.error) {
      throw new Error('Current Opp Key Deal source failed: ' + oppDeptKey + ' / ' + (result && result.error ? result.error : 'empty payload'));
    }
    rows = rows.concat(result.rows || []);
  });
  return rows.filter(function(row) { return row && row.keyDeal === true; });
}

function FcstSnapshot_getCurrentKeyDealsForPeriod_(deptKey, periodKey) {
  return FcstSnapshot_getCurrentOppKeyDealRows_(deptKey).filter(function(row) {
    return FcstSnapshot_oppRowMatchesPeriod_(row, periodKey);
  }).map(function(row) {
    return FcstSnapshot_mapOppRowToKeyDeal_(row);
  }).sort(function(a, b) {
    return Math.abs(Number(b.monthlyMrr) || 0) - Math.abs(Number(a.monthlyMrr) || 0);
  });
}

function FcstSnapshot_getOppDeptKeysForFcstDept_(deptKey) {
  var targetDept = String(deptKey || '').trim();
  var seen = {};
  var keys = [];
  OrgMasterReader_getRows().forEach(function(row) {
    if (!row || String(row.departmentCode || '').trim() !== targetDept) return;
    var groupName = String(row.groupName || '').trim();
    if (!groupName || seen[groupName]) return;
    seen[groupName] = true;
    keys.push(groupName);
  });
  if (!keys.length) throw new Error('Opp group mapping missing for FCST dept: ' + targetDept);
  return keys;
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
    var values = FcstSnapshot_getAllValues_(4);

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
  var values = FcstSnapshot_getAllValues_(4);
  var targetPeriod = FcstSnapshot_resolveTrendPeriodKey_(periodKey, liveData) || String(periodKey || '').trim();
  var data = {
    labels: [],
    weeks: [],
    dates: [],
    series: { target: [], fcstAdjusted: [], fcstCommit: [], confirmed: [], expectedMrr: [] },
    points: []
  };
  var snapshotMap = {};

  values.slice().sort(function(a, b) {
    var ad = a && a[0] instanceof Date ? a[0].getTime() : 0;
    var bd = b && b[0] instanceof Date ? b[0].getTime() : 0;
    return ad - bd;
  }).forEach(function(row) {
    var d = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) return;
    if (!nameRaw.startsWith(deptKey + ':')) return;
    var period = String(row[2] || '').trim();
    if (!FcstSnapshot_periodMatchesTarget_(period, targetPeriod)) return;

    var payload;
    try {
      payload = JSON.parse(String(row[3] || '{}'));
    } catch (e) {
      payload = {};
    }

    if (!FcstSnapshot_isDepartmentTotalRowForDept_(payload, deptKey, nameRaw)) return;
    var dateKey = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
    if (!snapshotMap[dateKey]) snapshotMap[dateKey] = { payloads: [], date: d, backfilled: false };
    snapshotMap[dateKey].payloads.push(payload);
    if (payload.__meta && payload.__meta.backfilled) snapshotMap[dateKey].backfilled = true;
  });

  Object.keys(snapshotMap).sort().forEach(function(dateKey) {
    var entry = snapshotMap[dateKey] || {};
    var payload = FcstSnapshot_sumMetricList_(entry.payloads || []);
    var label = Utilities.formatDate(entry.date, 'Asia/Tokyo', 'M/d');
    var metrics = FcstSnapshot_extractTrendMetrics_(payload);
    var keyDeals = FcstSnapshot_getSnapshotKeyDealsForPeriod_(deptKey, dateKey, targetPeriod);

    data.dates.push(dateKey);
    data.weeks.push(label);
    data.series.target.push(metrics.target);
    data.series.fcstAdjusted.push(metrics.fcstAdjusted);
    data.series.fcstCommit.push(metrics.fcstCommit);
    data.series.confirmed.push(metrics.confirmed);
    data.series.expectedMrr.push(metrics.expectedMrr);
    data.points.push({
      snapshotKey: dateKey,
      date: dateKey,
      label: label,
      isLive: false,
      isBackfilled: !!entry.backfilled,
      metrics: metrics,
      keyDealCount: keyDeals.length,
      keyDealPreview: FcstSnapshot_extractKeyDealPreview_(keyDeals)
    });
  });

  var livePeriod = FcstSnapshot_findTrendLivePeriod_(liveData, targetPeriod);
  var liveMetric = FcstSnapshot_findTrendLiveMetric_(liveData, livePeriod, deptKey);
  if (liveMetric) {
    var liveMetrics = FcstSnapshot_extractTrendMetrics_(liveMetric);
    var liveKeyDeals = FcstSnapshot_getCurrentKeyDealsForPeriod_(deptKey, livePeriod);
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
  var resolverData = liveData || FcstSnapshot_getLatestMembers(deptKey, { includeKeyDeals: false }) || { periodOptions: [] };
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
    result.keyDeals = FcstSnapshot_getCurrentKeyDealsForPeriod_(deptKey, livePeriod);
    return result;
  }

  if (!targetPeriod) return result;

  var values = FcstSnapshot_getAllValues_(4).slice().sort(function(a, b) {
    var ad = a && a[0] instanceof Date ? a[0].getTime() : 0;
    var bd = b && b[0] instanceof Date ? b[0].getTime() : 0;
    return ad - bd;
  });
  var matchedPayloads = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var d = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) continue;
    if (!nameRaw.startsWith(deptKey + ':')) continue;
    if (!FcstSnapshot_periodMatchesTarget_(String(row[2] || '').trim(), targetPeriod)) continue;
    if (Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') !== snapshotKey) continue;

    var payload;
    try {
      payload = JSON.parse(String(row[3] || '{}'));
    } catch (e) {
      payload = {};
    }

    if (!FcstSnapshot_isDepartmentTotalRowForDept_(payload, deptKey, nameRaw)) continue;
    matchedPayloads.push(payload);
  }

  if (matchedPayloads.length) {
    return {
      snapshotKey: snapshotKey,
      isLive: false,
      metrics: FcstSnapshot_extractTrendMetrics_(FcstSnapshot_sumMetricList_(matchedPayloads)),
      keyDeals: FcstSnapshot_getSnapshotKeyDealsForPeriod_(deptKey, snapshotKey, targetPeriod)
    };
  }

  return result;
}

function FcstSnapshot_isDepartmentTotalRowForDept_(payload, deptKey, nameRaw) {
  var expectedDept = String(deptKey || '').trim();
  var rowInfo = FcstSnapshot_parseRowName_(nameRaw);
  var rowDept = String(rowInfo.deptKey || '').trim();
  var rawMeta = payload && payload.__meta || {};
  var metaDept = String(rawMeta.dept || '').trim();
  var meta = FcstSnapshot_normalizeMeta_(payload, rowDept || expectedDept, rowInfo.name);
  if (!expectedDept) return false;
  if (rowDept && rowDept !== expectedDept) return false;
  if (metaDept && metaDept !== expectedDept) return false;
  return meta.totalKind === SHARED_TOTAL_KIND.DEPARTMENT;
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

function FcstSnapshot_getLatestDateKey_(deptKey, values, beforeDateKey) {
  var latestKey = '';
  var upperBound = String(beforeDateKey || '').trim();
  (values || []).forEach(function(row) {
    var d = row[0];
    var nameRaw = String(row[1] || '').trim();
    if (!(d instanceof Date) || isNaN(d)) return;
    if (deptKey && !nameRaw.startsWith(deptKey + ':')) return;
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
    if (upperBound && key >= upperBound) return;
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
  PROPOSAL_PRODUCT_METRIC_KEYS.concat(CONFIRMED_PROPOSAL_PRODUCT_METRIC_KEYS).forEach(function(metricKey) {
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
