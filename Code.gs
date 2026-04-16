function doGet(e) {
  if (e && e.parameter && e.parameter.app === 'mrr') {
    return mrrDashboard_doGet_();
  }
  var deptKey = (e && e.parameter && e.parameter.dept) || null;
  var email = Session.getActiveUser().getEmail();
  var userDefaultDept = UserReader_getUserDefaultDept(email) || null;
  var deptConfigJson = safeJsonForTemplate_(DEPT_CONFIG);

  if (!email.endsWith('@sansan.com')) {
    return HtmlService.createHtmlOutput('<h1>アクセス権限がありません</h1>');
  }

  if (!deptKey || !DEPT_CONFIG[deptKey]) {
    var tmpl = HtmlService.createTemplateFromFile('index');
    tmpl.selectedDept = 'null';
    tmpl.deptConfigJson = deptConfigJson;
    tmpl.userDefaultDept = userDefaultDept || 'null';
    tmpl.userEmail = email;
    tmpl.embeddedInitData = 'null';
    return tmpl.evaluate()
      .setTitle('FCST Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var tmpl = HtmlService.createTemplateFromFile('index');
  tmpl.selectedDept = deptKey;
  tmpl.deptConfigJson = deptConfigJson;
  tmpl.userDefaultDept = userDefaultDept || 'null';
  tmpl.userEmail = email;
  tmpl.embeddedInitData = 'null';
  return tmpl.evaluate()
    .setTitle('FCST Dashboard - ' + DEPT_CONFIG[deptKey].label)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function safeJsonForTemplate_(value) {
  return JSON.stringify(value)
    .replace(/</g, '\\u003c')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
}

function getUserDefaultDept_(email) {
  try {
    return UserReader_getUserDefaultDept(email);
  } catch(e) {
    return null;
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function debug_getInitData_BOCS() {
  var data = AppDataCache_getInitData('BOCS');
  return {
    hasData: !!data,
    topLevelKeys: Object.keys(data || {}),
    membersCount: Array.isArray(data && data.members) ? data.members.length : 0,
    periodOptionsCount: Array.isArray(data && data.periodOptions) ? data.periodOptions.length : 0,
    lastUpdated: data && data.lastUpdated || '',
    sfLastUpdated: data && data.sfLastUpdated || '',
    firstMemberName: data && data.members && data.members[0] ? data.members[0].name : '',
    firstMemberKeys: data && data.members && data.members[0] ? Object.keys(data.members[0]).slice(0, 20) : [],
    firstPeriodOption: data && data.periodOptions && data.periodOptions[0] ? data.periodOptions[0] : null
  };
}

function debug_getClientInitMeta_BOCS() {
  var html = doGet({ parameter: { dept: 'BOCS' } }).getContent();
  var match = html.match(/<script id="gas-embedded-init-data-json" type="application\/json">([\s\S]*?)<\/script>/);
  var deptMatch = html.match(/<script id="gas-dept-config-json" type="application\/json">([\s\S]*?)<\/script>/);
  var raw = match ? match[1] : '';
  var deptRaw = deptMatch ? deptMatch[1] : '';
  var parsed = null;
  var parseError = '';
  try {
    parsed = raw ? JSON.parse(raw) : null;
  } catch (e) {
    parseError = String(e && e.message ? e.message : e);
  }
  return {
    htmlLength: html.length,
    embeddedInitFound: !!match,
    embeddedInitLength: raw.length,
    embeddedInitHead: raw.slice(0, 240),
    embeddedInitTail: raw.slice(Math.max(0, raw.length - 240)),
    deptConfigFound: !!deptMatch,
    deptConfigLength: deptRaw.length,
    parseError: parseError,
    parsedKeys: parsed ? Object.keys(parsed) : [],
    parsedMembersCount: parsed && parsed.members ? parsed.members.length : 0,
    parsedPeriodOptionsCount: parsed && parsed.periodOptions ? parsed.periodOptions.length : 0
  };
}

function debug_getRenderedShellMeta() {
  var html = doGet({ parameter: { dept: 'BOCS' } }).getContent();
  return {
    htmlLength: html.length,
    hasMainContent: html.indexOf('id="main-content"') !== -1,
    hasLoadingOverlay: html.indexOf('id="loading-overlay"') !== -1,
    hasDeptConfigTag: html.indexOf('id="gas-dept-config-json"') !== -1,
    hasEmbeddedInitTag: html.indexOf('id="gas-embedded-init-data-json"') !== -1,
    hasSelectedDept: html.indexOf("var GAS_SELECTED_DEPT = 'BOCS';") !== -1,
    hasInitPage: html.indexOf('function initPage_()') !== -1,
    hasRenderDashboard: html.indexOf('App.renderDashboard = function(data)') !== -1,
    hasBootError: html.indexOf('function renderBootError_(error)') !== -1,
    headSnippet: html.slice(0, 400),
    tailSnippet: html.slice(Math.max(0, html.length - 400))
  };
}

function getCoefficientWebhookUrl_() {
  var props = PropertiesService.getScriptProperties();
  var propUrl = props.getProperty('COEFFICIENT_REFRESH_WEBHOOK_URL');
  if (propUrl) return propUrl;
  if (typeof COEFFICIENT_REFRESH_WEBHOOK_URL_MIGRATION === 'string' && COEFFICIENT_REFRESH_WEBHOOK_URL_MIGRATION) {
    props.setProperty('COEFFICIENT_REFRESH_WEBHOOK_URL', COEFFICIENT_REFRESH_WEBHOOK_URL_MIGRATION);
    return COEFFICIENT_REFRESH_WEBHOOK_URL_MIGRATION;
  }
  return '';
}

function getCoefficientRefreshThrottleInfo_() {
  var props = PropertiesService.getScriptProperties();
  var lastTriggeredAtRaw = props.getProperty('COEFFICIENT_REFRESH_LAST_TRIGGERED_AT');
  var lastTriggeredAt = Number(lastTriggeredAtRaw || 0) || 0;
  var now = Date.now();
  var elapsed = lastTriggeredAt ? (now - lastTriggeredAt) : COEFFICIENT_REFRESH_MIN_INTERVAL_MS;
  var remainingMs = Math.max(0, COEFFICIENT_REFRESH_MIN_INTERVAL_MS - elapsed);
  return {
    lastTriggeredAt: lastTriggeredAt,
    remainingMs: remainingMs,
    canRefresh: remainingMs === 0
  };
}

function formatThrottleWait_(remainingMs) {
  var totalSeconds = Math.ceil((Number(remainingMs) || 0) / 1000);
  var minutes = Math.floor(totalSeconds / 60);
  var seconds = totalSeconds % 60;
  if (minutes <= 0) return seconds + '秒';
  if (seconds === 0) return minutes + '分';
  return minutes + '分' + seconds + '秒';
}

function recordLatestFetchTrace_(kind, stage, detail) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('LATEST_FETCH_TRACE', JSON.stringify({
    kind: kind,
    stage: stage,
    detail: detail || '',
    recordedAt: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')
  }));
}

function getLatestFetchTrace_() {
  var raw = PropertiesService.getScriptProperties().getProperty('LATEST_FETCH_TRACE');
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return { kind: '', stage: 'trace-parse-error', detail: String(e && e.message ? e.message : e), recordedAt: '' };
  }
}

function assertCoefficientRefreshAllowed_() {
  var throttle = getCoefficientRefreshThrottleInfo_();
  if (throttle.canRefresh) return;
  throw new Error('更新は5分以上の間隔をあけてください。あと' + formatThrottleWait_(throttle.remainingMs) + 'で実行できます。');
}

function triggerSfDataRefresh_(options) {
  var opts = options || {};
  if (!opts.skipThrottle) assertCoefficientRefreshAllowed_();
  var url = getCoefficientWebhookUrl_();
  if (!url) throw new Error('Coefficient webhook URL is not configured');

  recordLatestFetchTrace_(opts.kind || 'unknown', 'webhook-start', '');
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    muteHttpExceptions: true,
    followRedirects: true
  });
  var status = response.getResponseCode();
  var body = response.getContentText() || '';
  if (status < 200 || status >= 300) {
    recordLatestFetchTrace_(opts.kind || 'unknown', 'webhook-error', 'HTTP ' + status + ' ' + body.slice(0, 200));
    throw new Error('Coefficient refresh failed: HTTP ' + status + ' ' + body.slice(0, 200));
  }
  PropertiesService.getScriptProperties().setProperty('COEFFICIENT_REFRESH_LAST_TRIGGERED_AT', String(Date.now()));
  recordLatestFetchTrace_(opts.kind || 'unknown', 'webhook-ok', 'HTTP ' + status + ' ' + body.slice(0, 200));
  return {
    ok: true,
    status: status,
    body: body.slice(0, 200)
  };
}

function waitForRefreshedData_(fetcher, getLastUpdated, previousLastUpdated) {
  var startedAt = Date.now();
  var latestResult = fetcher();
  if (latestResult && latestResult.error) return latestResult;
  while (Date.now() - startedAt < COEFFICIENT_REFRESH_TIMEOUT_MS) {
    var latestUpdated = getLastUpdated(latestResult);
    if (latestUpdated && latestUpdated !== previousLastUpdated) return latestResult;
    Utilities.sleep(COEFFICIENT_REFRESH_POLL_INTERVAL_MS);
    latestResult = fetcher();
    if (latestResult && latestResult.error) return latestResult;
  }
  latestResult.refreshTimedOut = true;
  return latestResult;
}

function getTrendData(deptKey, block) {
  try {
    var users = UserReader_getUsers(deptKey);
    var targets = TargetReader_getTargets(deptKey);
    var result = SfDataReader_getAggregated(deptKey, users, targets);
    var fcstState = FcstAdjusted_getState(deptKey);
    var fcstAdj = fcstState.adjusted;
    var periods = FcstPeriods_expandKeys_(result.periodOptions || []);
    result.members.forEach(function(member) {
      periods.forEach(function(p) {
        var key = member.name + '|' + p;
        if (!member[p]) member[p] = {};
        member[p].fcstAdjusted = fcstAdj[key] || { net: 0, newExp: 0, churn: 0 };
      });
    });
    return FcstSnapshot_getTrendData(deptKey, block, result);
  } catch (e) { return { error: e.message }; }
}

function getOppSnapshotData(deptKey, dateStr) { try { return OppListSnapshot_getByDate(deptKey, dateStr); } catch (e) { return { error: e.message }; } }
function saveFcstAdjusted(deptKey, p) { try { return FcstAdjusted_save(deptKey, p); } catch (e) { return { error: e.message }; } }
function saveOppSfValue(deptKey, p) { try { return OppListWriter_saveOppSfValue(deptKey, p); } catch (e) { return { error: e.message }; } }
function saveOppDrafts(deptKey, changes) { try { return OppListWriter_saveDrafts(deptKey, changes); } catch (e) { return { error: e.message }; } }
function saveNote(deptKey, p) { try { return FcstAdjusted_save(deptKey, p); } catch (e) { return { error: e.message }; } }
function saveFcstAdjusted2(deptKey, p) { try { return FcstAdjusted_save(deptKey, p); } catch (e) { return { error: e.message }; } }
function getSnapshotDates(deptKey) { try { return FcstSnapshot_getSnapshotDates(deptKey); } catch (e) { return { error: e.message }; } }
function getSnapshotData(deptKey, dateStr) { try { return FcstSnapshot_getDataByDate(deptKey, dateStr); } catch (e) { return { error: e.message }; } }

function runFcstSnapshotForAllDepts() {
  var results = [];
  Object.keys(DEPT_CONFIG).forEach(function(key) {
    try {
      results.push({ dept: key, result: createSnapshot(key) });
    } catch (e) {
      Logger.log('FcstSnapshot failed for ' + key + ': ' + e.message);
      results.push({ dept: key, error: e.message });
    }
  });
  return { ok: true, results: results };
}

function runOppListSnapshotForAllDepts() {
  Object.keys(DEPT_CONFIG).forEach(function(key) {
    try {
      OppListSnapshot_createWeekly(key);
    } catch(e) {
      Logger.log('OppListSnapshot failed for ' + key + ': ' + e.message);
    }
  });
  return { ok: true };
}

function createOppSnapshot(deptKey) {
  try {
    if (!deptKey) return runOppListSnapshotForAllDepts();
    return OppListSnapshot_createWeekly(deptKey);
  } catch (e) { return { error: e.message }; }
}

function setupOppSnapshotTrigger(deptKey) { try { return OppListSnapshot_setupWeeklyTrigger(); } catch (e) { return { error: e.message }; } }
function setupOppListSnapshotTrigger() { try { return OppListSnapshot_setupWeeklyTrigger(); } catch (e) { return { error: e.message }; } }

function buildInitData_(deptKey) {
  return AppDataCache_refreshInitData(deptKey);
}

function getInitData(deptKey) {
  try {
    var result = AppDataCache_getInitData(deptKey);
    if (!result) {
      return { error: 'データが見つかりません。スプレッドシートのデータを確認してください。' };
    }
    return { data: result };
  } catch (e) {
    return { error: e.message };
  }
}

function refreshFromSpreadsheet(deptKey) {
  try {
    var data = AppDataCache_refreshInitData(deptKey);
    return { data: data };
  } catch (e) {
    return { error: e.message };
  }
}

function refreshSfDataAndGetInitData(deptKey) {
  var before;
  var refresh;
  var result;
  try {
    recordLatestFetchTrace_('fcst', 'start', '');
    Logger.log('[FCST latest] start');
    before = getInitData(deptKey);
    Logger.log('[FCST latest] before lastUpdated=%s', before && before.data ? before.data.lastUpdated : '');
    if (before.error) return { error: '[初期データ読込] ' + before.error };
    var previousLastUpdated = (before.data && before.data.lastUpdated) || '';
    recordLatestFetchTrace_('fcst', 'before-ready', previousLastUpdated || '(empty)');
    refresh = triggerSfDataRefresh_({ kind: 'fcst' });
    Logger.log('[FCST latest] webhook status=%s body=%s', refresh && refresh.status, refresh && refresh.body);
    var startedAt = Date.now();
    var latestLastUpdated = previousLastUpdated;
    while (Date.now() - startedAt < COEFFICIENT_REFRESH_TIMEOUT_MS) {
      latestLastUpdated = AggregatedCache_getSfLastUpdated_(deptKey);
      if (latestLastUpdated && latestLastUpdated !== previousLastUpdated) break;
      Utilities.sleep(COEFFICIENT_REFRESH_POLL_INTERVAL_MS);
    }

    if (latestLastUpdated && latestLastUpdated !== previousLastUpdated) {
      result = { data: AppDataCache_refreshInitData(deptKey) };
    } else {
      result = before;
      result.refreshTimedOut = true;
    }

    Logger.log('[FCST latest] after lastUpdated=%s timedOut=%s', result && result.data ? result.data.lastUpdated : '', !!(result && result.refreshTimedOut));
    if (result && result.error) return { error: '[更新後データ確認] ' + result.error, refresh: refresh };
    result.refresh = refresh;
    result.trace = getLatestFetchTrace_();
    if (result.refreshTimedOut) {
      recordLatestFetchTrace_('fcst', 'timed-out', previousLastUpdated + ' -> ' + ((result.data && result.data.lastUpdated) || '(empty)'));
    } else {
      recordLatestFetchTrace_('fcst', 'done', previousLastUpdated + ' -> ' + ((result.data && result.data.lastUpdated) || '(empty)'));
    }
    return result;
  } catch (e) {
    Logger.log('[FCST latest] error=%s', e && e.message ? e.message : e);
    recordLatestFetchTrace_('fcst', 'error', String(e && e.message ? e.message : e));
    var stage = refresh ? '[Webhook実行待機' : (before ? '[更新前処理' : '[更新開始');
    return { error: stage + ' ' + e.message, refresh: refresh || null, trace: getLatestFetchTrace_() };
  }
}

function getOpportunities(deptKey) {
  try {
    return AppDataCache_getOpportunities(deptKey);
  } catch (e) {
    return { error: e.message };
  }
}

function refreshSfDataAndGetOpportunities(deptKey) {
  var before;
  var refresh;
  var result;
  try {
    recordLatestFetchTrace_('opps', 'start', '');
    Logger.log('[OPPS latest] start');
    before = getOpportunities(deptKey);
    Logger.log('[OPPS latest] before lastUpdated=%s', before ? before.lastUpdated : '');
    if (before.error) return { error: '[案件データ読込] ' + before.error };
    var previousLastUpdated = before.lastUpdated || '';
    recordLatestFetchTrace_('opps', 'before-ready', previousLastUpdated || '(empty)');
    refresh = triggerSfDataRefresh_({ kind: 'opps' });
    Logger.log('[OPPS latest] webhook status=%s body=%s', refresh && refresh.status, refresh && refresh.body);
    result = waitForRefreshedData_(
      function() { return getOpportunities(deptKey); },
      function(payload) { return payload ? payload.lastUpdated : ''; },
      previousLastUpdated
    );
    Logger.log('[OPPS latest] after lastUpdated=%s timedOut=%s', result ? result.lastUpdated : '', !!(result && result.refreshTimedOut));
    if (result && result.error) return { error: '[更新後案件データ確認] ' + result.error, refresh: refresh };
    result.refresh = refresh;
    result.trace = getLatestFetchTrace_();
    if (result.refreshTimedOut) {
      recordLatestFetchTrace_('opps', 'timed-out', previousLastUpdated + ' -> ' + ((result && result.lastUpdated) || '(empty)'));
    } else {
      recordLatestFetchTrace_('opps', 'done', previousLastUpdated + ' -> ' + ((result && result.lastUpdated) || '(empty)'));
    }
    return result;
  } catch (e) {
    Logger.log('[OPPS latest] error=%s', e && e.message ? e.message : e);
    recordLatestFetchTrace_('opps', 'error', String(e && e.message ? e.message : e));
    var stage = refresh ? '[Webhook実行待機' : (before ? '[更新前処理' : '[更新開始');
    return { error: stage + ' ' + e.message, refresh: refresh || null, trace: getLatestFetchTrace_() };
  }
}

function createSnapshot(deptKey) {
  try {
    if (!deptKey) return runFcstSnapshotForAllDepts();
    var masterContext = AssignmentMaster_getContext(deptKey);
    var result = SfDataReader_getAggregated(deptKey, masterContext);
    var fcstState = FcstAdjusted_getState(deptKey);
    var fcstAdj = fcstState.adjusted;
    var notes = fcstState.notes;
    var periods = FcstPeriods_expandKeys_(result.periodOptions || []);
    result.members.forEach(function(member) {
      periods.forEach(function(p) {
        var key = member.name + '|' + p;
        if (!member[p]) member[p] = {};
        member[p].fcstAdjusted = fcstAdj[key] || { net: 0, newExp: 0, churn: 0 };
      });
    });
    return FcstSnapshot_create(deptKey, result.members, notes, periods);
  } catch (e) {
    return { error: e.message };
  }
}

function setupSnapshotTrigger(deptKey) {
  try {
    return FcstSnapshot_setupWeeklyTrigger();
  } catch (e) {
    return { error: e.message };
  }
}

function authorizeCoefficientRefresh(deptKey) {
  var result = triggerSfDataRefresh_({ skipThrottle: true });
  Logger.log(JSON.stringify(result));
  return result;
}
