var MRR_DASHBOARD_INITIAL_SNAPSHOT_DATE_LIMIT = 2;
var MRR_DASHBOARD_CACHE_TTL_SECONDS = 300;
var MRR_DASHBOARD_CACHE_PREFIX = 'snapshotLiveData:v12:mrr:';
var MRR_DASHBOARD_CACHE_CHUNK_SIZE = 85000;
var MRR_DASHBOARD_DIVISION_ORDER = ['SS', 'BO', 'CO'];
var MRR_DASHBOARD_ALL_DIVISION = 'COO';
var MRR_DASHBOARD_LIVE_KEY = 'live';

function mrrDashboard_doGet_(e) {
  var tmpl = HtmlService.createTemplateFromFile('mrr-index');
  var initialDivision = e && e.parameter ? MrrDashboard_normalizeDivisionSelection_(e.parameter.division) : '';
  var webAppUrl = '';
  try {
    webAppUrl = ScriptApp.getService().getUrl() || '';
  } catch (err) {
    webAppUrl = '';
  }
  tmpl.mrrInitialDivision = initialDivision || '';
  tmpl.mrrWebAppUrl = webAppUrl || '';
  return tmpl.evaluate()
    .setTitle('MRR進捗ダッシュボード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getMrrDashboardData(division, beforeDate, periodKey) {
  var selection = MrrDashboard_normalizeDivisionSelection_(division) || 'SS';
  var normalizedBeforeDate = MrrDashboard_normalizeSnapshotDate_(beforeDate);
  var normalizedPeriodKey = MrrDashboard_normalizePeriodKey_(periodKey);
  var cacheKey = MRR_DASHBOARD_CACHE_PREFIX + selection + ':snapshots:before:' + (normalizedBeforeDate || 'latest') + ':period:' + (normalizedPeriodKey || 'default');
  var cached = MrrDashboard_cacheGet_(cacheKey);
  if (cached) return cached;

  var result = MrrDashboard_getSnapshotData_(selection, {
    beforeDate: normalizedBeforeDate,
    limit: MRR_DASHBOARD_INITIAL_SNAPSHOT_DATE_LIMIT,
    periodKey: normalizedPeriodKey
  });
  MrrDashboard_cachePut_(cacheKey, result);
  return result;
}

function getMrrDashboardCurrentData(division, periodKey) {
  var selection = MrrDashboard_normalizeDivisionSelection_(division) || 'SS';
  var normalizedPeriodKey = MrrDashboard_normalizePeriodKey_(periodKey);
  var cacheKey = MRR_DASHBOARD_CACHE_PREFIX + selection + ':current:period:' + (normalizedPeriodKey || 'default');
  var cached = MrrDashboard_cacheGet_(cacheKey);
  if (cached) return MrrDashboard_applyFreshCurrentTargets_(cached, selection, normalizedPeriodKey);

  var result = MrrDashboard_getSnapshotData_(selection, {
    currentOnly: true,
    limit: MRR_DASHBOARD_INITIAL_SNAPSHOT_DATE_LIMIT,
    periodKey: normalizedPeriodKey
  });
  result = MrrDashboard_applyFreshCurrentTargets_(result, selection, normalizedPeriodKey);
  if (MrrDashboard_isCurrentResultCacheable_(result)) {
    MrrDashboard_cachePut_(cacheKey, result);
  }
  return result;
}

function MrrDashboard_prewarmCurrentCache() {
  var result = { ok: true, warmed: {}, errors: {} };
  MrrDashboard_getDivisionChoices_().forEach(function(choice) {
    var key = choice && choice.key;
    if (!key) return;
    try {
      var payload = getMrrDashboardCurrentData(key);
      result.warmed[key] = {
        weeks: payload && payload.weeks ? payload.weeks.length : 0,
        diagnostics: payload && payload.diagnostics ? payload.diagnostics.current || [] : []
      };
    } catch (e) {
      result.ok = false;
      result.errors[key] = String(e && e.message ? e.message : e);
    }
  });
  return result;
}

function MrrDashboard_isCurrentResultCacheable_(result) {
  if (!result || !Array.isArray(result.weeks)) return false;
  if (result.weeks.indexOf(MRR_DASHBOARD_LIVE_KEY) === -1) return false;
  var diagnostics = result.diagnostics && result.diagnostics.current;
  if (!Array.isArray(diagnostics)) return true;
  return diagnostics.every(function(item) {
    return item && item.status === 'ok';
  });
}

function getMrrDashboardChoices() {
  return MrrDashboard_getDivisionChoices_();
}

function MrrDashboard_normalizeDivisionSelection_(division) {
  var value = String(division || '').trim().toUpperCase();
  if (value === MRR_DASHBOARD_ALL_DIVISION) return value;
  return MRR_DASHBOARD_DIVISION_ORDER.indexOf(value) !== -1 ? value : '';
}

function MrrDashboard_getDivisionChoices_() {
  return [
    { key: 'SS', label: 'SS', description: 'SS事業部' },
    { key: 'BO', label: 'BO', description: 'BO事業部' },
    { key: 'CO', label: 'CO', description: 'CO事業部' },
    { key: MRR_DASHBOARD_ALL_DIVISION, label: 'COO', description: '全事業部' }
  ];
}

function MrrDashboard_getSelectedDivisionKeys_(selection) {
  var normalized = MrrDashboard_normalizeDivisionSelection_(selection);
  if (normalized === MRR_DASHBOARD_ALL_DIVISION) return MRR_DASHBOARD_DIVISION_ORDER.slice();
  return normalized ? [normalized] : ['SS'];
}

function MrrDashboard_getDivisionLabel_(selection) {
  var normalized = MrrDashboard_normalizeDivisionSelection_(selection) || 'SS';
  if (normalized === MRR_DASHBOARD_ALL_DIVISION) return 'COO (全事業部)';
  return normalized;
}

function MrrDashboard_getTotalLabel_(selection) {
  var normalized = MrrDashboard_normalizeDivisionSelection_(selection) || 'SS';
  if (normalized === MRR_DASHBOARD_ALL_DIVISION) return 'COO全体';
  return normalized + '全体';
}

function MrrDashboard_normalizePeriodKey_(periodKey) {
  return String(periodKey || '').trim();
}

function MrrDashboard_cacheGet_(key) {
  try {
    var cache = CacheService.getScriptCache();
    var raw = cache.get(key);
    if (raw) return JSON.parse(raw);
    return MrrDashboard_cacheReadChunked_(cache, key);
  } catch (e) {
    return null;
  }
}

function MrrDashboard_cachePut_(key, value) {
  try {
    var cache = CacheService.getScriptCache();
    var raw = JSON.stringify(value);
    if (raw.length <= MRR_DASHBOARD_CACHE_CHUNK_SIZE) {
      MrrDashboard_cacheRemoveChunked_(cache, key);
      cache.put(key, raw, MRR_DASHBOARD_CACHE_TTL_SECONDS);
      return;
    }
    cache.remove(key);
    MrrDashboard_cacheWriteChunked_(cache, key, raw);
  } catch (e) {}
}

function MrrDashboard_cacheWriteChunked_(cache, key, raw) {
  var chunks = Math.ceil(raw.length / MRR_DASHBOARD_CACHE_CHUNK_SIZE);
  cache.put(key + ':chunks', String(chunks), MRR_DASHBOARD_CACHE_TTL_SECONDS);
  for (var i = 0; i < chunks; i++) {
    cache.put(
      key + ':chunk:' + i,
      raw.substr(i * MRR_DASHBOARD_CACHE_CHUNK_SIZE, MRR_DASHBOARD_CACHE_CHUNK_SIZE),
      MRR_DASHBOARD_CACHE_TTL_SECONDS
    );
  }
}

function MrrDashboard_cacheReadChunked_(cache, key) {
  var count = parseInt(cache.get(key + ':chunks') || '0', 10);
  if (!count) return null;
  var parts = [];
  for (var i = 0; i < count; i++) {
    var chunk = cache.get(key + ':chunk:' + i);
    if (!chunk) return null;
    parts.push(chunk);
  }
  try {
    return JSON.parse(parts.join(''));
  } catch (e) {
    MrrDashboard_cacheRemoveChunked_(cache, key);
    return null;
  }
}

function MrrDashboard_cacheRemoveChunked_(cache, key) {
  var count = parseInt(cache.get(key + ':chunks') || '0', 10);
  if (!count) return;
  var keys = [key + ':chunks'];
  for (var i = 0; i < count; i++) {
    keys.push(key + ':chunk:' + i);
  }
  try { cache.removeAll(keys); } catch (e) {}
}
