var MRR_DASHBOARD_INITIAL_SNAPSHOT_DATE_LIMIT = 2;
var MRR_DASHBOARD_CACHE_TTL_SECONDS = 300;
var MRR_DASHBOARD_CACHE_PREFIX = 'snapshotLiveData:v6:mrr:';
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

function getMrrDashboardData(division, beforeDate) {
  var selection = MrrDashboard_normalizeDivisionSelection_(division) || 'SS';
  var normalizedBeforeDate = MrrDashboard_normalizeSnapshotDate_(beforeDate);
  var cacheKey = MRR_DASHBOARD_CACHE_PREFIX + selection + ':snapshots:before:' + (normalizedBeforeDate || 'latest');
  var cached = MrrDashboard_cacheGet_(cacheKey);
  if (cached) return cached;

  var result = MrrDashboard_getSnapshotData_(selection, {
    beforeDate: normalizedBeforeDate,
    limit: MRR_DASHBOARD_INITIAL_SNAPSHOT_DATE_LIMIT
  });
  MrrDashboard_cachePut_(cacheKey, result);
  return result;
}

function getMrrDashboardCurrentData(division) {
  var selection = MrrDashboard_normalizeDivisionSelection_(division) || 'SS';
  var cacheKey = MRR_DASHBOARD_CACHE_PREFIX + selection + ':current';
  var cached = MrrDashboard_cacheGet_(cacheKey);
  if (cached) return cached;

  var result = MrrDashboard_getSnapshotData_(selection, {
    currentOnly: true,
    limit: MRR_DASHBOARD_INITIAL_SNAPSHOT_DATE_LIMIT
  });
  MrrDashboard_cachePut_(cacheKey, result);
  return result;
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

function MrrDashboard_cacheGet_(key) {
  try {
    var raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function MrrDashboard_cachePut_(key, value) {
  try {
    var raw = JSON.stringify(value);
    if (raw.length < 90000) {
      CacheService.getScriptCache().put(key, raw, MRR_DASHBOARD_CACHE_TTL_SECONDS);
    }
  } catch (e) {}
}
