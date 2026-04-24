function doGet(e) {
  if (e && e.parameter && e.parameter.app === 'mrr') {
    return mrrDashboard_doGet_();
  }
  if (e && e.parameter && e.parameter.diag === '1') {
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>FCST Diagnostic</title></head><body style="font:16px/1.5 sans-serif;padding:24px;">' +
      '<h1>FCST Diagnostic</h1>' +
      '<p>version: 132+</p>' +
      '<p>mode: minimal-html-output</p>' +
      '<p>time: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss') + '</p>' +
      '</body></html>'
    ).setTitle('FCST Diagnostic').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var deptKey = (e && e.parameter && e.parameter.dept) || null;
  var email = Session.getActiveUser().getEmail();
  var deptConfigJson = safeJsonForTemplate_(getDeptConfigMap_());
  var webAppUrl = ScriptApp.getService().getUrl() || 'https://script.google.com/a/macros/sansan.com/s/AKfycbwmFoDiI8auNskiCFqYzOpKQKaS_Tf9-_bAZGDPs_Y/exec';

  if (!email.endsWith('@sansan.com')) {
    return HtmlService.createHtmlOutput('<h1>アクセス権限がありません</h1>');
  }

  if (!deptKey || !isValidDeptKey_(deptKey)) {
    var tmpl = HtmlService.createTemplateFromFile('index');
    tmpl.selectedDept = 'null';
    tmpl.deptConfigJson = deptConfigJson;
    tmpl.userDefaultDept = 'null';
    tmpl.userEmail = email;
    tmpl.webAppUrl = webAppUrl;
    tmpl.embeddedInitData = 'null';
    return tmpl.evaluate()
      .setTitle('FCST Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var initData = null;
  try {
    initData = AppDataCache_getInitData(deptKey);
  } catch(err) {
    initData = null;
  }

  var tmpl = HtmlService.createTemplateFromFile('index');
  tmpl.selectedDept = deptKey;
  tmpl.deptConfigJson = deptConfigJson;
  tmpl.userDefaultDept = 'null';
  tmpl.userEmail = email;
  tmpl.webAppUrl = webAppUrl;
  tmpl.embeddedInitData = 'null';
  return tmpl.evaluate()
    .setTitle('FCST Dashboard - ' + getDeptConfig_(deptKey).label)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function safeJsonForTemplate_(value) {
  return JSON.stringify(value)
    .replace(/</g, '\\u003c')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
}

function getUserDefaultDept_(email) {
  return null;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getClientBundle() {
  var content = HtmlService.createHtmlOutputFromFile('js').getContent();
  content = content.replace(/^\s*<script>\s*/, '');
  content = content.replace(/\s*<\/script>\s*$/, '');
  return content;
}

function debug_getRenderedShellMeta(deptKey) {
  var resolvedDeptKey = deptKey || 'SSEP3';
  var output = doGet({ parameter: { dept: resolvedDeptKey } });
  var html = output.getContent();
  var selectedDeptMatch = html.match(/var GAS_SELECTED_DEPT = '([^']*)';/);
  return {
    requestedDept: resolvedDeptKey,
    htmlLength: html.length,
    selectedDept: selectedDeptMatch ? selectedDeptMatch[1] : null,
    embeddedInitDataIsNull: html.indexOf('id="gas-embedded-init-data-json" type="application/json">null</script>') !== -1,
    usesPathnameRouting: html.indexOf('window.location.pathname') !== -1,
    usesTopTargetLinks: html.indexOf('target="_top"') !== -1,
    lazyLoadsClientBundle: html.indexOf("google.script.run.withSuccessHandler(function(source)") !== -1,
    headPreview: html.slice(0, 500),
    tailPreview: html.slice(Math.max(0, html.length - 500))
  };
}

function debug_getClientInitMeta(deptKey) {
  var resolvedDeptKey = deptKey || 'SSEP3';
  try {
    var initRaw = getInitData(resolvedDeptKey);
    var clientRaw = getClientInitData(resolvedDeptKey);
    var initPayload = initRaw && initRaw.data ? initRaw.data : initRaw;
    var clientPayload = clientRaw && clientRaw.data ? clientRaw.data : clientRaw;
    return {
      deptKey: resolvedDeptKey,
      initError: initRaw && initRaw.error ? initRaw.error : '',
      clientError: clientRaw && clientRaw.error ? clientRaw.error : '',
      initJsonLength: initPayload ? JSON.stringify(initPayload).length : 0,
      clientJsonLength: clientPayload ? JSON.stringify(clientPayload).length : 0,
      memberCount: clientPayload && clientPayload.members ? clientPayload.members.length : 0,
      periodOptionCount: clientPayload && clientPayload.periodOptions ? clientPayload.periodOptions.length : 0,
      firstMemberName: clientPayload && clientPayload.members && clientPayload.members.length ? (clientPayload.members[0].displayName || clientPayload.members[0].name || '') : '',
      sfLastUpdated: clientPayload && clientPayload.sfLastUpdated ? clientPayload.sfLastUpdated : '',
      lastUpdated: clientPayload && clientPayload.lastUpdated ? clientPayload.lastUpdated : ''
    };
  } catch (e) {
    return {
      deptKey: resolvedDeptKey,
      fatalError: String(e && e.message ? e.message : e)
    };
  }
}

function debug_getClientInitMeta_SSEP3() {
  return debug_getClientInitMeta('SSEP3');
}

function debug_getClientInitMeta_BOCS() {
  return debug_getClientInitMeta('BOCS');
}

function rebuildAssignmentMaster() {
  var result = AssignmentMaster_build();
  resetDeptConfigCache_();
  getDeptKeys_().forEach(function(deptKey) {
    invalidateDeptCaches_(deptKey, { opps: true, trend: true, init: true });
  });
  return result;
}

function setupMasterSheets() {
  return MasterSchema_setupSheets();
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
  throw new Error('最新取得は5分間隔です。あと' + formatThrottleWait_(throttle.remainingMs) + 'で再度お試しください。');
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

function getTrendData(deptKey, periodKey) {
  var live = AggregatedCache_read(deptKey) || null;
  return FcstSnapshot_getTrendData(deptKey, periodKey, live);
}

function getTrendWeekDetails(deptKey, periodKey, snapshotKey) {
  return FcstSnapshot_getTrendWeekDetails(deptKey, periodKey, snapshotKey);
}

function runCreateSnapshot_forDate(dateStr, deptKey, force) {
  return FcstSnapshot_runCreateSnapshot_forDate(dateStr, deptKey, force);
}

function runBackfillMissingWeeks(deptKey, weeksBack) {
  return FcstSnapshot_runBackfillMissingWeeks(deptKey, weeksBack);
}

// ---------------------------------------------------------------------------
// Init data / client payloads
// ---------------------------------------------------------------------------

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
    return { error: e && e.message ? e.message : String(e) };
  }
}

function getClientInitData(deptKey) {
  return getInitData(deptKey);
}

function getClientPeriodData(deptKey, periodKey) {
  try {
    var data = buildInitData_(deptKey);
    if (!data) return { error: 'データが見つかりません' };
    var rows = buildRowsForPeriod_(data.members, periodKey);
    return {
      data: {
        periodKey: String(periodKey || ''),
        rows: rows,
        periodOptions: data.periodOptions || [],
        lastUpdated: data.lastUpdated || '',
        sfLastUpdated: data.sfLastUpdated || ''
      }
    };
  } catch (e) {
    return { error: e && e.message ? e.message : String(e) };
  }
}

function getClientSnapshotData(deptKey, dateStr, periodKey) {
  try {
    var snapshot = FcstSnapshot_getDataByDate(deptKey, dateStr);
    if (!snapshot) return { error: 'スナップショットが見つかりません' };
    var resolvedPeriod = String(periodKey || '').trim();
    var periodOptions = snapshot.periodOptions || [];
    if (!resolvedPeriod && periodOptions.length) resolvedPeriod = periodOptions[0].key || '';
    var rows = buildRowsForPeriod_(snapshot.members, resolvedPeriod);
    return {
      date: snapshot.date || dateStr,
      timestampKey: snapshot.timestampKey || '',
      periodKey: resolvedPeriod,
      periodOptions: periodOptions,
      rows: rows
    };
  } catch (e) {
    return { error: e && e.message ? e.message : String(e) };
  }
}

function buildRowsForPeriod_(members, periodKey) {
  var list = Array.isArray(members) ? members : [];
  var key = String(periodKey || '');
  return list.map(function(member) {
    var row = {};
    Object.keys(member || {}).forEach(function(k) { row[k] = member[k]; });
    if (key && member && member[key]) {
      var metric = member[key];
      Object.keys(metric).forEach(function(mk) {
        if (!row.hasOwnProperty(mk)) row[mk] = metric[mk];
      });
    }
    return row;
  });
}

function primeClientPayloadCaches_(deptKey, data) {
  // Intentionally a no-op placeholder: cached payloads are served by getClientInitData via AggregatedCache_read.
  return;
}

// ---------------------------------------------------------------------------
// Snapshots & opportunities
// ---------------------------------------------------------------------------

function getSnapshotDates(deptKey) {
  try { return FcstSnapshot_getSnapshotDates(deptKey); } catch (e) { return { error: e.message }; }
}

function getSnapshotData(deptKey, dateStr) {
  try { return FcstSnapshot_getDataByDate(deptKey, dateStr); } catch (e) { return { error: e.message }; }
}

function saveFcstAdjusted(deptKey, p) {
  try { return FcstAdjusted_save(deptKey, p); } catch (e) { return { error: e.message }; }
}

function saveFcstAdjusted2(deptKey, p) {
  try { return FcstAdjusted_save(deptKey, p); } catch (e) { return { error: e.message }; }
}

function saveNote(deptKey, p) {
  try { return FcstAdjusted_save(deptKey, p); } catch (e) { return { error: e.message }; }
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
    return { error: e && e.message ? e.message : String(e) };
  }
}

function refreshSfDataAndGetInitData(deptKey) {
  var before;
  var refresh;
  var result;
  try {
    recordLatestFetchTrace_('fcst', 'start', '');
    before = getInitData(deptKey);
    if (before.error) return { error: '[初期データ読込] ' + before.error };
    var previousLastUpdated = (before.data && before.data.lastUpdated) || '';
    recordLatestFetchTrace_('fcst', 'before-ready', previousLastUpdated || '(empty)');
    refresh = triggerSfDataRefresh_({ kind: 'fcst' });
    var startedAt = Date.now();
    var latestLastUpdated = previousLastUpdated;
    while (Date.now() - startedAt < COEFFICIENT_REFRESH_TIMEOUT_MS) {
      latestLastUpdated = AggregatedCache_getSfLastUpdated_(deptKey);
      if (latestLastUpdated && latestLastUpdated !== previousLastUpdated) break;
      Utilities.sleep(COEFFICIENT_REFRESH_POLL_INTERVAL_MS);
    }
    if (latestLastUpdated && latestLastUpdated !== previousLastUpdated) {
      result = { data: AggregatedCache_refresh(deptKey) };
    } else {
      result = before;
      result.refreshTimedOut = true;
    }
    if (result && result.error) return { error: '[更新後データ確認] ' + result.error, refresh: refresh };
    result.refresh = refresh;
    result.trace = getLatestFetchTrace_();
    recordLatestFetchTrace_('fcst', result.refreshTimedOut ? 'timed-out' : 'done',
      previousLastUpdated + ' -> ' + ((result.data && result.data.lastUpdated) || '(empty)'));
    return result;
  } catch (e) {
    recordLatestFetchTrace_('fcst', 'error', String(e && e.message ? e.message : e));
    var stage = refresh ? '[Webhook実行待機' : (before ? '[更新前処理' : '[更新開始');
    return { error: stage + ' ' + (e && e.message ? e.message : e), refresh: refresh || null, trace: getLatestFetchTrace_() };
  }
}

function refreshSfDataAndGetOpportunities(deptKey) {
  var before;
  var refresh;
  var result;
  try {
    recordLatestFetchTrace_('opps', 'start', '');
    before = getOpportunities(deptKey);
    if (before.error) return { error: '[案件データ読込] ' + before.error };
    var previousLastUpdated = before.lastUpdated || '';
    recordLatestFetchTrace_('opps', 'before-ready', previousLastUpdated || '(empty)');
    refresh = triggerSfDataRefresh_({ kind: 'opps' });
    result = waitForRefreshedData_(
      function() { return getOpportunities(deptKey); },
      function(payload) { return payload ? payload.lastUpdated : ''; },
      previousLastUpdated
    );
    if (result && result.error) return { error: '[更新後案件データ確認] ' + result.error, refresh: refresh };
    result.refresh = refresh;
    result.trace = getLatestFetchTrace_();
    recordLatestFetchTrace_('opps', result.refreshTimedOut ? 'timed-out' : 'done',
      previousLastUpdated + ' -> ' + ((result && result.lastUpdated) || '(empty)'));
    return result;
  } catch (e) {
    recordLatestFetchTrace_('opps', 'error', String(e && e.message ? e.message : e));
    var stage = refresh ? '[Webhook実行待機' : (before ? '[更新前処理' : '[更新開始');
    return { error: stage + ' ' + (e && e.message ? e.message : e), refresh: refresh || null, trace: getLatestFetchTrace_() };
  }
}

function createSnapshot(deptKey) {
  if (!deptKey) {
    var allFcstResults = getDeptKeys_().map(function(key) {
      var result = createSnapshot(key) || {};
      var entry = { deptKey: key };
      Object.keys(result).forEach(function(prop) { entry[prop] = result[prop]; });
      return entry;
    });
    return {
      ok: !allFcstResults.some(function(entry) { return !!entry.error; }),
      results: allFcstResults
    };
  }

  var fcstResult;
  try {
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
    var live = AggregatedCache_read(deptKey) || AggregatedCache_refresh(deptKey);
    var input = FcstSnapshot_buildSnapshotInputFromLive_(live);
    fcstResult = FcstSnapshot_create(deptKey, input.members, input.notesMap, input.periodKeys);
  } catch (e) {
    fcstResult = { error: e && e.message ? e.message : String(e) };
  }

  SnapshotExecutionLog_record_({
    kind: 'fcst',
    action: 'createSnapshot',
    deptKey: deptKey,
    ok: !fcstResult.error,
    skipped: !!fcstResult.skipped,
    count: Number(fcstResult.count) || 0,
    snapshotDate: String(fcstResult.date || '').trim(),
    snapshotAt: String(fcstResult.snapshotAt || '').trim(),
    captureMode: String(fcstResult.captureMode || '').trim(),
    error: fcstResult.error || ''
  });
  Logger.log('FCST snapshot execution: dept=' + deptKey +
    ' ok=' + (!fcstResult.error) +
    ' skipped=' + (!!fcstResult.skipped) +
    ' count=' + (Number(fcstResult.count) || 0) +
    ' date=' + String(fcstResult.date || '') +
    ' snapshotAt=' + String(fcstResult.snapshotAt || '') +
    ' captureMode=' + String(fcstResult.captureMode || '') +
    (fcstResult.error ? ' error=' + String(fcstResult.error) : ''));
  return fcstResult;
}

function createSnapshotBySheet(sheetKey) {
  try {
    var result = SnapshotManual_runBySheet_(sheetKey, createSnapshot);
    SnapshotManual_logResult_('fcst', result);
    return result;
  } catch (e) {
    var errorResult = { error: e && e.message ? e.message : String(e), sheetKey: SnapshotManual_normalizeSheetKey_(sheetKey) };
    SnapshotManual_logResult_('fcst', errorResult);
    return errorResult;
  }
}

function setupSnapshotTrigger(deptKey) {
  try { return FcstSnapshot_setupWeeklyTrigger(); } catch (e) { return { error: e.message }; }
}

function createOppSnapshot(deptKey) {
  if (!deptKey) {
    var allOppResults = OppSnapshot_getDeptKeys_().map(function(key) {
      var result = createOppSnapshot(key) || {};
      var entry = { deptKey: key };
      Object.keys(result).forEach(function(prop) { entry[prop] = result[prop]; });
      return entry;
    });
    return {
      ok: !allOppResults.some(function(entry) { return !!entry.error; }),
      results: allOppResults
    };
  }

  var oppResult;
  try {
    oppResult = OppListSnapshot_createWeekly(deptKey);
  } catch (e) {
    oppResult = { error: e.message };
  }

  SnapshotExecutionLog_record_({
    kind: 'opp',
    action: 'createOppSnapshot',
    deptKey: deptKey,
    ok: !oppResult.error,
    skipped: !!oppResult.skipped,
    count: Number(oppResult.count) || 0,
    snapshotDate: String(oppResult.date || '').trim(),
    snapshotAt: String(oppResult.snapshotAt || '').trim(),
    error: oppResult.error || ''
  });
  Logger.log('Opp snapshot execution: dept=' + deptKey +
    ' ok=' + (!oppResult.error) +
    ' skipped=' + (!!oppResult.skipped) +
    ' count=' + (Number(oppResult.count) || 0) +
    ' date=' + String(oppResult.date || '') +
    ' snapshotAt=' + String(oppResult.snapshotAt || '') +
    (oppResult.error ? ' error=' + String(oppResult.error) : ''));
  return oppResult;
}

function createOppSnapshotBySheet(sheetKey) {
  try {
    var result = SnapshotManual_runBySheet_(sheetKey, createOppSnapshot, OppSnapshot_getDeptKeysBySheet_);
    SnapshotManual_logResult_('opp', result);
    return result;
  } catch (e) {
    var errorResult = { error: e && e.message ? e.message : String(e), sheetKey: SnapshotManual_normalizeSheetKey_(sheetKey) };
    SnapshotManual_logResult_('opp', errorResult);
    return errorResult;
  }
}

function SnapshotManual_runBySheet_(sheetKey, runner, deptKeyResolver) {
  var normalizedSheetKey = SnapshotManual_normalizeSheetKey_(sheetKey);
  if (!normalizedSheetKey) {
    throw new Error('sheetKey は SS / SSCS / BO / CO のいずれかを指定してください');
  }

  var deptKeys = typeof deptKeyResolver === 'function'
    ? deptKeyResolver(normalizedSheetKey)
    : SnapshotManual_getDeptKeysBySheet_(normalizedSheetKey);
  if (!deptKeys.length) {
    var source = typeof DEPT_CONFIG_CACHE_SOURCE_ === 'string' ? DEPT_CONFIG_CACHE_SOURCE_ : '';
    var totalDeptCount = typeof getDeptKeys_ === 'function' ? getDeptKeys_().length : 0;
    return {
      ok: false,
      error: '対象部門が見つかりません source=' + source + ' total=' + totalDeptCount + ' sheet=' + normalizedSheetKey,
      sheetKey: normalizedSheetKey,
      deptKeys: [],
      results: []
    };
  }

  var results = deptKeys.map(function(deptKey) {
    var result = runner(deptKey) || {};
    var entry = { deptKey: deptKey };
    Object.keys(result).forEach(function(key) {
      entry[key] = result[key];
    });
    if (result && result.error) Logger.log('Snapshot manual run failed: ' + normalizedSheetKey + ' / ' + deptKey + ' / ' + result.error);
    return entry;
  });

  return {
    ok: !results.some(function(entry) { return !!entry.error; }),
    sheetKey: normalizedSheetKey,
    deptKeys: deptKeys,
    results: results
  };
}

function SnapshotManual_logResult_(kind, result) {
  var label = kind === 'opp' ? 'Opp snapshot' : 'FCST snapshot';
  if (!result) {
    Logger.log(label + ' manual run result: empty');
    return;
  }

  if (result.error) {
    Logger.log(label + ' manual run failed: sheet=' + String(result.sheetKey || '') +
      ' error=' + String(result.error));
  }

  Logger.log(label + ' manual run: sheet=' + String(result.sheetKey || '') +
    ' ok=' + (!!result.ok) +
    ' depts=' + (result.deptKeys || []).join(','));

  if (!(result.results || []).length) {
    Logger.log(label + ' manual run details: no result rows');
  }

  (result.results || []).forEach(function(entry) {
    Logger.log(label + ' detail: dept=' + String(entry.deptKey || '') +
      ' ok=' + (!entry.error) +
      ' count=' + (Number(entry.count) || 0) +
      ' date=' + String(entry.date || entry.snapshotAt || '') +
      (entry.error ? ' error=' + String(entry.error) : ''));
  });
}

function SnapshotManual_notifyResult_(kind, result) {
  var ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    return;
  }
  if (!ss || typeof ss.toast !== 'function') return;

  var label = kind === 'opp' ? '案件Snapshot' : 'FCST Snapshot';
  var message = SnapshotManual_buildToastMessage_(result);
  ss.toast(message, label, 10);
}

function SnapshotManual_buildToastMessage_(result) {
  if (!result) return '結果を取得できませんでした';
  if (result.error) return String(result.error);
  if (!result.deptKeys || !result.deptKeys.length) return '対象部門がありません';

  var parts = [];
  parts.push('sheet=' + String(result.sheetKey || ''));

  (result.results || []).forEach(function(entry) {
    var text = String(entry.deptKey || '') + ':';
    if (entry.error) {
      text += 'error';
    } else {
      text += String(Number(entry.count) || 0) + '件';
      if (entry.date || entry.snapshotAt) text += '@' + String(entry.date || entry.snapshotAt);
    }
    parts.push(text);
  });

  return parts.join(' / ');
}

function SnapshotExecutionLog_record_(entry) {
  try {
    var sheet = SnapshotExecutionLog_getOrCreateSheet_();
    if (!sheet) return;

    var now = new Date();
    var deptKey = String(entry && entry.deptKey || '').trim();
    var cfg = deptKey ? getDeptConfig_(deptKey) : null;
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 13).setValues([[
      Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
      String(entry && entry.kind || '').trim(),
      String(entry && entry.action || '').trim(),
      String(entry && entry.captureMode || '').trim(),
      String(entry && entry.requestedSheetKey || '').trim(),
      cfg && cfg.sfSheetKey ? String(cfg.sfSheetKey) : '',
      deptKey,
      entry && entry.ok ? 'TRUE' : 'FALSE',
      entry && entry.skipped ? 'TRUE' : 'FALSE',
      Number(entry && entry.count) || 0,
      String(entry && entry.snapshotDate || '').trim(),
      String(entry && entry.snapshotAt || '').trim(),
      String(entry && entry.error || '').trim()
    ]]);
  } catch (e) {
    Logger.log('SnapshotExecutionLog failed: ' + (e && e.message ? e.message : e));
  }
}

function SnapshotExecutionLog_getOrCreateSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('SnapshotExecutionLog');
  if (!sheet) {
    sheet = ss.insertSheet('SnapshotExecutionLog');
  }
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, 13).setValues([[
      'executed_at',
      'kind',
      'action',
      'capture_mode',
      'requested_sheet_key',
      'sheet_key',
      'dept_key',
      'ok',
      'skipped',
      'count',
      'snapshot_date',
      'snapshot_at',
      'error'
    ]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function SnapshotManual_normalizeSheetKey_(sheetKey) {
  var key = String(sheetKey || '').trim().toUpperCase();
  return ['SS', 'SSCS', 'BO', 'CO'].indexOf(key) !== -1 ? key : '';
}

function SnapshotManual_getDeptKeysBySheet_(sheetKey) {
  return getDeptKeys_().filter(function(deptKey) {
    var cfg = getDeptConfig_(deptKey);
    return !!(cfg && cfg.sfSheetKey === sheetKey);
  });
}

function manualCreateFcstSnapshot_SS() {
  return createSnapshotBySheet('SS');
}

function manualCreateFcstSnapshot_SSCS() {
  return createSnapshotBySheet('SSCS');
}

function manualCreateFcstSnapshot_BO() {
  return createSnapshotBySheet('BO');
}

function manualCreateFcstSnapshot_CO() {
  return createSnapshotBySheet('CO');
}

function manualCreateOppSnapshot_SS() {
  return createOppSnapshotBySheet('SS');
}

function manualCreateOppSnapshot_SSCS() {
  return createOppSnapshotBySheet('SSCS');
}

function manualCreateOppSnapshot_BO() {
  return createOppSnapshotBySheet('BO');
}

function manualCreateOppSnapshot_CO() {
  return createOppSnapshotBySheet('CO');
}

function getOppSnapshotData(deptKey, dateStr) {
  try { return OppListSnapshot_getByDate(deptKey, dateStr); } catch (e) { return { error: e.message }; }
}

function saveOppSfValue(deptKey, p) {
  try { return OppListWriter_saveOppSfValue(deptKey, p); } catch (e) { return { error: e.message }; }
}

function saveOppDrafts(deptKey, changes) {
  try { return OppListWriter_saveDrafts(deptKey, changes); } catch (e) { return { error: e.message }; }
}

function authorizeCoefficientRefresh(deptKey) {
  var result = triggerSfDataRefresh_({ skipThrottle: true });
  Logger.log(JSON.stringify(result));
  return result;
}

// Opportunity snapshots use group_name as deptKey.
function OppSnapshot_getDeptKeys_() {
  var monthKey = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
  var seen = {};

  return OrgMasterReader_getRows().reduce(function(list, row) {
    if (!row) return list;
    if (row.startMonth && row.startMonth > monthKey) return list;
    if (row.endMonth && row.endMonth < monthKey) return list;

    var deptKey = String(row.groupName || '').trim();
    var sheetKey = DeptConfig_resolveSfSheetKey_(row.divisionCode, row.departmentCode);
    if (!deptKey || !sheetKey || seen[deptKey]) return list;

    seen[deptKey] = true;
    list.push(deptKey);
    return list;
  }, []).sort();
}

function OppSnapshot_getDeptKeysBySheet_(sheetKey) {
  var normalizedSheetKey = SnapshotManual_normalizeSheetKey_(sheetKey);
  return OppSnapshot_getDeptKeys_().filter(function(deptKey) {
    var row = OppListReader_getOrgDeptRow_(deptKey);
    if (!row) return false;
    return DeptConfig_resolveSfSheetKey_(row.divisionCode, row.departmentCode) === normalizedSheetKey;
  });
}
