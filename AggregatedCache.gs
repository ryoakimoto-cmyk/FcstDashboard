function AggregatedCache_read(deptKey) {
  var sheet = AggregatedCache_getSheet_(deptKey);
  if (!sheet || sheet.getLastRow() < 1) return null;

  var values = sheet.getRange(1, 1, sheet.getLastRow(), 3).getValues();
  var map = AggregatedCache_decodeMap_(deptKey, values);

  if (!map.hasOwnProperty('members')) return null;
  if (!Array.isArray(map.members)) return null;
  if (map.hasOwnProperty('periodOptions') && !Array.isArray(map.periodOptions)) return null;

  return {
    members: map.hasOwnProperty('members') ? map.members : [],
    fcstAdjusted: map.hasOwnProperty('adjusted') ? map.adjusted : {},
    notes: map.hasOwnProperty('notes') ? map.notes : {},
    weekOverWeekMap: map.hasOwnProperty('wowMap') ? map.wowMap : {},
    snapshotDates: map.hasOwnProperty('snapshotDates') ? map.snapshotDates : [],
    previousSnapshot: map.hasOwnProperty('previousSnapshot') ? map.previousSnapshot : null,
    periodOptions: map.hasOwnProperty('periodOptions') ? map.periodOptions : [],
    lastUpdated: String(map.lastUpdated || ''),
    latestSnapshotData: map.hasOwnProperty('latestSnapshotData') ? map.latestSnapshotData : null,
    sfLastUpdated: String(map.sfLastUpdated || map.lastUpdated || '')
  };
}

function AggregatedCache_write(deptKey, data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = AggregatedCache_getOrCreateSheet_(deptKey);
    var cachedAt = new Date().toISOString();
    var prefix = deptKey + ':';
    var rows = [];

    rows = rows.concat(AggregatedCache_encodeRows_(prefix + 'members', JSON.stringify(data.members || []), cachedAt));
    rows = rows.concat(AggregatedCache_encodeRows_(prefix + 'adjusted', JSON.stringify(data.fcstAdjusted || {}), cachedAt));
    rows = rows.concat(AggregatedCache_encodeRows_(prefix + 'notes', JSON.stringify(data.notes || {}), cachedAt));
    rows = rows.concat(AggregatedCache_encodeRows_(prefix + 'wowMap', JSON.stringify(data.weekOverWeekMap || {}), cachedAt));
    rows = rows.concat(AggregatedCache_encodeRows_(prefix + 'snapshotDates', JSON.stringify(data.snapshotDates || []), cachedAt));
    rows = rows.concat(AggregatedCache_encodeRows_(prefix + 'previousSnapshot', JSON.stringify(data.previousSnapshot || null), cachedAt));
    rows = rows.concat(AggregatedCache_encodeRows_(prefix + 'periodOptions', JSON.stringify(data.periodOptions || []), cachedAt));
    rows = rows.concat(AggregatedCache_encodeRows_(prefix + 'lastUpdated', String(data.lastUpdated || ''), cachedAt));
    rows = rows.concat(AggregatedCache_encodeRows_(prefix + 'sfLastUpdated', String(data.sfLastUpdated || data.lastUpdated || ''), cachedAt));
    rows = rows.concat(AggregatedCache_encodeRows_(prefix + 'latestSnapshotData', JSON.stringify(data.latestSnapshotData || null), cachedAt));
    rows = rows.concat(AggregatedCache_encodeRows_(prefix + 'cachedAt', cachedAt, cachedAt));

    var existingValues = sheet.getLastRow() > 0 ? sheet.getRange(1, 1, sheet.getLastRow(), 3).getValues() : [];
    var otherRows = existingValues.filter(function(row) {
      var key = String(row[0] || '');
      return key && key.indexOf(prefix) !== 0;
    });

    sheet.clearContents();
    var allRows = otherRows.concat(rows);
    if (allRows.length > 0) {
      sheet.getRange(1, 1, allRows.length, 3).setValues(allRows);
    }
    return { ok: true, cachedAt: cachedAt };
  } finally {
    lock.releaseLock();
  }
}

function AggregatedCache_writeKey(deptKey, dataKey, value) {
  var sheet = AggregatedCache_getOrCreateSheet_(deptKey);
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var key = deptKey + ':' + dataKey;
    var rowIdx = AggregatedCache_findRow_(sheet, key);
    var now = new Date();
    if (rowIdx > 0) {
      sheet.getRange(rowIdx, 2).setValue(JSON.stringify(value));
      sheet.getRange(rowIdx, 3).setValue(now);
    } else {
      sheet.appendRow([key, JSON.stringify(value), now]);
    }
  } finally {
    lock.releaseLock();
  }
}

function AggregatedCache_refresh(deptKey) {
  try {
    var context = AssignmentMaster_getContext(deptKey);
    var result = SfDataReader_getAggregated(deptKey, context);
    var fcstState = FcstAdjusted_getState(deptKey);
    result.notes = fcstState.notes;
    result.fcstAdjusted = fcstState.adjusted;
    result.weekOverWeekMap = FcstSnapshot_getWeekOverWeek(deptKey);
    result.snapshotDates = FcstSnapshot_getSnapshotDates(deptKey);
    result.previousSnapshot = FcstSnapshot_getLatestMembers(deptKey);
    result.latestSnapshotData = result.snapshotDates.length
      ? FcstSnapshot_getDataByDate(deptKey, result.snapshotDates[0])
      : null;
    result.sfLastUpdated = result.lastUpdated || AggregatedCache_getSfLastUpdated_(deptKey);

    if (!isProposalProductsEnabled_(deptKey)) {
      AggregatedCache_stripProposalProductFields_(result);
    }

    AggregatedCache_write(deptKey, result);
    try { CacheLayer_write(deptKey, 'initData', result); } catch (e) {}
    return result;
  } catch (e) {
    throw new Error('集計キャッシュ更新失敗: [' + (e && e.message ? e.message : String(e)) + ']');
  }
}

function AggregatedCache_stripProposalProductFields_(result) {
  var proposalFields = ['received', 'debtMgmt', 'debtMgmtLite', 'expense'];
  var periodKeys = [];

  (result.periodOptions || []).forEach(function(opt) {
    if (!opt || !opt.key) return;
    periodKeys.push(opt.key);
    periodKeys = periodKeys.concat(opt.months || []);
  });

  (result.members || []).forEach(function(member) {
    periodKeys.forEach(function(periodKey) {
      if (!member || !member[periodKey]) return;
      proposalFields.forEach(function(field) {
        delete member[periodKey][field];
      });
    });
  });

  if (result.latestSnapshotData && Array.isArray(result.latestSnapshotData.members)) {
    (result.latestSnapshotData.members || []).forEach(function(member) {
      periodKeys.forEach(function(periodKey) {
        if (!member || !member[periodKey]) return;
        proposalFields.forEach(function(field) {
          delete member[periodKey][field];
        });
      });
    });
  }
}

function AggregatedCache_runScheduledUpdate() {
  var hour = new Date().getHours();
  var isBusinessHours = (hour >= 9 && hour < 19);
  if (!isBusinessHours) {
    var lastRun = CacheService.getScriptCache().get('lastScheduledUpdate');
    if (lastRun && (Date.now() - parseInt(lastRun, 10)) < 3600000) return;
  }
  Object.keys(DEPT_CONFIG).forEach(function(dk) {
    AppDataCache_warmDept_(dk);
  });
  CacheService.getScriptCache().put('lastScheduledUpdate', String(Date.now()), 7200);
}

function AggregatedCache_setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'AggregatedCache_runScheduledUpdate') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('AggregatedCache_runScheduledUpdate')
    .timeBased().everyMinutes(5).create();
}

function AggregatedCache_getSheet_(deptKey) {
  return getSharedSheet(AGGREGATED_CACHE_SHEET_NAME);
}

function AggregatedCache_getOrCreateSheet_(deptKey) {
  var sheet = getSharedSheet(AGGREGATED_CACHE_SHEET_NAME);
  if (!sheet) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    sheet = ss.insertSheet(AGGREGATED_CACHE_SHEET_NAME);
  }
  return sheet;
}

function AggregatedCache_parseJson_(text, fallback) {
  if (text === '' || text === null || text === undefined) return fallback;
  return JSON.parse(String(text));
}

function AggregatedCache_hasMembers_(deptKey, sheet) {
  var values = sheet.getRange(1, 1, sheet.getLastRow(), 3).getValues();
  var map = AggregatedCache_decodeMap_(deptKey, values);
  return map.hasOwnProperty('members');
}

function AggregatedCache_readValue_(deptKey, sheet, key) {
  if (!sheet || sheet.getLastRow() < 1) return '';
  var values = sheet.getRange(1, 1, sheet.getLastRow(), 3).getValues();
  var map = AggregatedCache_decodeMap_(deptKey, values);
  return map.hasOwnProperty(key) ? String(map[key] || '') : '';
}

function AggregatedCache_findRow_(sheet, key) {
  var found = sheet.createTextFinder(key).matchEntireCell(true).findNext();
  if (!found) return 0;
  return found.getRow();
}

function AggregatedCache_encodeRows_(key, valueJson, cachedAt) {
  var chunkSize = 44000;
  if (String(valueJson).length <= chunkSize) {
    return [[key, valueJson, cachedAt]];
  }

  var rows = [];
  for (var i = 0, index = 0; i < valueJson.length; i += chunkSize, index++) {
    rows.push([key + '__' + index, valueJson.slice(i, i + chunkSize), cachedAt]);
  }
  return rows;
}

function AggregatedCache_decodeMap_(deptKey, values) {
  var prefix = deptKey + ':';
  var map = {};
  var chunkMap = {};

  values.forEach(function(row) {
    var rawKey = String(row[0] || '').trim();
    if (!rawKey) return;
    if (rawKey.indexOf(prefix) !== 0) return;
    var key = rawKey.slice(prefix.length);

    var match = key.match(/^(.*)__([0-9]+)$/);
    if (match) {
      var baseKey = match[1];
      var index = Number(match[2]);
      if (!chunkMap[baseKey]) chunkMap[baseKey] = [];
      chunkMap[baseKey].push({ index: index, value: String(row[1] || '') });
      return;
    }

    map[key] = AggregatedCache_decodeValue_(key, row[1]);
  });

  Object.keys(chunkMap).forEach(function(baseKey) {
    var joined = chunkMap[baseKey]
      .sort(function(a, b) { return a.index - b.index; })
      .map(function(chunk) { return chunk.value; })
      .join('');
    map[baseKey] = AggregatedCache_decodeValue_(baseKey, joined);
  });

  return map;
}

function AggregatedCache_decodeValue_(key, value) {
  if (key === 'lastUpdated' || key === 'cachedAt' || key === 'sfLastUpdated') {
    return String(value || '');
  }
  return AggregatedCache_parseJson_(value, null);
}

function AggregatedCache_getSfLastUpdated_(deptKey) {
  var cached = CacheLayer_read(deptKey, 'sfLastUpdated');
  if (cached) return cached;
  var sheet = getSfDataSheet_(deptKey);
  if (!sheet) return '';
  var lastUpdated = SfDataReader_extractLastUpdated_(sheet.getRange(1, 1));
  CacheLayer_write(deptKey, 'sfLastUpdated', lastUpdated);
  return lastUpdated;
}
