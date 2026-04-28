const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');

function read(name) {
  return fs.readFileSync(path.join(root, name), 'utf8');
}

function assertIncludes(haystack, needle, message) {
  if (!haystack.includes(needle)) {
    throw new Error(message + `: missing "${needle}"`);
  }
}

function assertNotIncludes(haystack, needle, message) {
  if (haystack.includes(needle)) {
    throw new Error(message + `: unexpected "${needle}"`);
  }
}

function assertNoMojibake(name, source) {
  const mojibake = /[隱繧繝騾謐荳譛蜈莠驛螟髮]/;
  const match = source.match(mojibake);
  if (match) {
    throw new Error(`${name}: possible mojibake character "${match[0]}"`);
  }
}

function assertCount(haystack, needle, expected, message) {
  const actual = haystack.split(needle).length - 1;
  if (actual !== expected) {
    throw new Error(`${message}: expected ${expected}, got ${actual} for "${needle}"`);
  }
}

const cacheLayer = read('CacheLayer.gs');
assertIncludes(cacheLayer, "var CACHE_PREFIX = 'fcst:';", 'cache prefix changed');
assertIncludes(cacheLayer, 'var CACHE_TTL_DEFAULT_SECONDS = 300;', 'shared cache default TTL must stay at 5 minutes');
assertIncludes(cacheLayer, 'function CacheLayer_getTtl_(dataKey)', 'shared cache TTL helper missing');
assertIncludes(cacheLayer, 'function CacheLayer_removeChunked_(cache, key)', 'chunked cache invalidation missing');
assertIncludes(cacheLayer, "'initData'", 'initData cache invalidation missing');
assertIncludes(cacheLayer, "'oppList'", 'oppList cache invalidation missing');
assertIncludes(cacheLayer, 'persistToSheet === false', 'ephemeral cache option missing');

const appDataCache = read('AppDataCache.gs');
assertIncludes(appDataCache, "AppDataCache_getData_(deptKey, 'initData')", 'initData shared getter missing');
assertIncludes(appDataCache, "AppDataCache_getData_(deptKey, 'oppList')", 'oppList shared getter missing');
assertIncludes(appDataCache, "CacheLayer_read(deptKey, dataKey, { skipSharedSheet: true })", 'shared cache read path missing');
assertIncludes(appDataCache, 'function AppDataCache_getData_(deptKey, dataKey)', 'shared cache getter missing');
assertIncludes(appDataCache, 'function AppDataCache_warmDept_(deptKey)', 'shared warmup helper missing');

const aggregatedCache = read('AggregatedCache.gs');
[
  "prefix + 'members'",
  "prefix + 'adjusted'",
  "prefix + 'notes'",
  "prefix + 'wowMap'",
  "prefix + 'snapshotDates'",
  "prefix + 'previousSnapshot'",
  "prefix + 'periodOptions'",
  "prefix + 'lastUpdated'",
  "prefix + 'sfLastUpdated'",
  "prefix + 'latestSnapshotData'",
  "prefix + 'cachedAt'"
].forEach((token) => assertIncludes(aggregatedCache, token, 'aggregated cache payload shape changed'));
assertIncludes(aggregatedCache, 'AppDataCache_warmDept_(dk);', 'scheduled warmup must prime shared init/opp caches');
assertIncludes(aggregatedCache, '.timeBased().everyMinutes(5).create();', 'scheduled warmup must run every 5 minutes');

const sfDataReader = read('SfDataReader.gs');
['groupCode', 'totalKind', 'fcstMin', 'fcstMax'].forEach((token) => {
  assertIncludes(sfDataReader, token, 'shared member/metric shape incomplete');
});

const config = read('Config.gs');
assertIncludes(config, 'function isProposalProductsEnabled_(deptKey)', 'proposalProducts helper missing');
[
  "const OPP_WEBAPP_INPUT_SHEET_NAME = '案件_WebApp入力';",
  "const OPP_HISTORY_V2_SHEET_NAME = 'Opp履歴';",
  "const SNAPSHOT_META_SHEET_NAME = 'スナップショットメタ';",
  "const FCST_DECLARATION_CURRENT_SHEET_NAME = 'FCST宣言_現在';"
].forEach((token) => assertIncludes(config, token, 'opp history v2 sheet constants missing'));
assertNotIncludes(config, 'const OPP_LIST_SNAPSHOT_SHEET_NAME', 'legacy opp snapshot sheet constant must be removed');

const code = read('Code.gs');
assertIncludes(code, 'AppDataCache_getInitData', 'Code.gs must use shared init cache');
assertIncludes(code, 'AppDataCache_getOpportunities', 'Code.gs must use shared opp cache');
assertIncludes(code, 'AssignmentMaster_getContext', 'Code.gs must use shared assignment context');
assertIncludes(code, 'function getEmbeddedInitDataForDept_(deptKey)', 'dept embedded init helper missing');
assertIncludes(code, "tmpl.embeddedInitData = getEmbeddedInitDataForDept_(deptKey);", 'dept page must use embedded init helper');
assertIncludes(code, "tmpl.webAppUrl = webAppUrl || 'null';", 'dept page must embed absolute web app url');
assertIncludes(code, "webAppUrl = ScriptApp.getService().getUrl() || '';", 'web app url resolution missing');
assertIncludes(code, 'return json.length <= 800000 ? json : \'null\';', 'embedded init size guard missing');
assertNotIncludes(code, 'function createOppSnapshotV2(deptKey)', 'legacy v2 opp snapshot wrapper must be removed');
assertNotIncludes(code, 'function setupOppSnapshotTriggerV2()', 'legacy v2 opp snapshot trigger wrapper must be removed');

const oppSnapshot = read('OppListSnapshot.gs');
[
  'function OppListSnapshot_createWeekly(deptKey)',
  'function OppListSnapshot_getSnapshotDates(deptKey)',
  'function OppListSnapshot_getByDate(deptKey, dateStr)',
  'function OppListSnapshot_setupWeeklyTrigger()',
  "return OppHistory_createWeekly_(deptKey);",
  "return OppHistory_getSnapshotDates_(deptKey);",
  "return OppHistory_getByDate_(deptKey, dateStr);",
  "trigger.getHandlerFunction() === 'createOppSnapshot'"
].forEach((token) => assertIncludes(oppSnapshot, token, 'canonical opp snapshot API missing'));

const oppHistory = read('OppHistory.gs');
[
  'function OppHistory_createWeekly_(deptKey)',
  'function OppHistory_getSnapshotDates_(deptKey)',
  'function OppHistory_getByDate_(deptKey, dateStr)',
  'function OppHistory_ensureInfrastructure_()',
  'function OppHistory_determineTracking_(snapshotType, currentMap, currentOrder, previousActivePayloadById)',
  'function OppHistory_buildPayload_(deptKey, snapshotType, row, webInput, previousPayload, isNew)',
  'function OppHistory_buildDroppedPayload_(deptKey, snapshotType, previousPayload)',
  'function OppHistory_payloadToLegacyRow_(payload, snapshotDate)',
  'function OppHistory_trimOld_(sheet, snapshotDate)',
  'function OppHistory_formatTimestamp_(date)',
  "Math.ceil(day / 7) === 2 ? 'monthly_dropout' : 'weekly'",
  "payload._meta.status = OPP_HISTORY_STATUS_REMOVED_FROM_P;",
  "'確度': OppHistory_numberOrNull_(row.confidence)",
  "'KeyDeal': webInput ? !!webInput.keyDeal : false",
  "'FCSTコメント': webInput ? String(webInput.fcstComment || '') : ''"
].forEach((token) => assertIncludes(oppHistory, token, 'opp history v2 implementation incomplete'));
assertCount(oppHistory, 'OppHistory_trimOld_(infra.historySheet, snapshotDate)', 1, 'opp history trim must run once per batch');
assertNotIncludes(oppHistory, 'OppListSnapshot_createWeekly_v2', 'legacy v2 snapshot function name must be removed');

const oppListReader = read('OppListReader.gs');
[
  'scheduleDate:',
  'closeDate:',
  'confidence:',
  'function OppListReader_toNullableNumber_(value)'
].forEach((token) => assertIncludes(oppListReader, token, 'opp list reader mapping incomplete for opp history v2'));

const client = read('js.html');
assertNoMojibake('js.html', client);
['isDepartmentTotalMember_', 'isGroupTotalMember_', 'getMemberGroupLabel_'].forEach((token) => {
  assertIncludes(client, token, 'FCST client total handling incomplete');
});
assertIncludes(client, 'TTL: 300000,', 'client cache TTL must stay at 5 minutes');
assertIncludes(client, 'ClientCache.shouldRevalidate(cached)', 'client should only revalidate stale caches');
assertIncludes(client, 'App.refreshInitDataSilently_({ data: _embeddedInitData });', 'embedded init data must refresh in background');
assertIncludes(client, 'App.refreshInitDataSilently_(cached);', 'cached init data must refresh in background');
assertIncludes(client, 'if (isValidInitDataPayload_(_embeddedInitData)) {', 'embedded init data must be validated before render');
assertIncludes(client, 'if (cached && isValidInitDataPayload_(cached.data)) {', 'cached init data must be validated before render');
assertIncludes(client, '!isValidInitDataPayload_(App.state.currentData)', 'silent init refresh must recover broken current state');
assertIncludes(client, 'App.reloadFcstFromSS();', 'invalid init payload must fall back to spreadsheet refresh');
assertIncludes(client, "if (!isValidInitDataPayload_(result && result.data)) {", 'applyFcstLatestData must reject incomplete payloads');
assertIncludes(client, 'function startInitLoadTimeout_()', 'FCST init timeout helper missing');
assertIncludes(client, 'clearInitLoadTimeout_();', 'FCST init timeout clear missing');
assertIncludes(client, 'function navigateToWebApp_(deptKey)', 'web app navigation helper missing');
assertIncludes(client, 'navigateToWebApp_(dept.key);', 'dept card must navigate through web app helper');
assertIncludes(client, 'btn.onclick = function() { navigateToWebApp_(null); };', 'back navigation must use web app helper');
assertIncludes(client, "var nextView = document.getElementById('view-' + view);", 'switchView null guard missing');
assertIncludes(client, 'function getSharedQuarterDefinitions_(periodOptions)', 'shared quarter definitions missing');
assertCount(client, 'function isDepartmentTotalMember_(member)', 1, 'duplicate department total helper');
assertCount(client, 'function isGroupTotalMember_(member)', 1, 'duplicate group total helper');
assertCount(client, 'function getMemberGroupLabel_(member)', 1, 'duplicate group label helper');
assertCount(client, 'function displayMemberName(member)', 1, 'duplicate member display helper');
assertCount(client, 'function renderTopPage_()', 1, 'duplicate top page renderer');

const mrr = read('mrr-index.html');
assertNoMojibake('mrr-index.html', mrr);
['MRR進捗ダッシュボード', '読み込み中...', '全部署', 'getMrrDashboardData();'].forEach((token) => {
  assertIncludes(mrr, token, 'MRR template markers missing');
});

[
  'data-division="SS"',
  'data-division="BO"',
  "switchDivision('BO')",
  '.getMrrDashboardData(division);',
  'labels: D.weeks.map(getWeekLabel_)',
  'Array.isArray(metric.keyDealsData) && metric.keyDealsData.length',
  'function buildMetricDatasets_(rows)',
  'metricDefinitions: Array.isArray(result && result.metricDefinitions)'
].forEach((token) => {
  assertIncludes(mrr, token, 'MRR division/snapshot client path missing');
});
assertCount(
  mrr,
  'metricDefinitions: Array.isArray(result && result.metricDefinitions)',
  2,
  'all active MRR result normalizers must preserve metric definitions'
);

const mrrDashboard = read('MrrDashboard.gs');
[
  'function getMrrDashboardData(division)',
  "if (normalizedDivision === 'BO')",
  'return MrrDashboard_getBoData_();',
  'function MrrDashboard_getSsData_()',
  "totalDeptKey: 'SS'",
  "allLabel: '全事業部'"
].forEach((token) => {
  assertIncludes(mrrDashboard, token, 'MRR dashboard dispatcher incomplete');
});

const mrrSnapshot = read('MrrSnapshotDashboard.gs');
assertNoMojibake('MrrSnapshotDashboard.gs', mrrSnapshot);
assertNotIncludes(mrrSnapshot, 'Object.keys(opp.dates', 'BO MRR week list must come from FCST snapshots only');
assertNotIncludes(mrrSnapshot, 'Fallback_', 'BO MRR runtime must not use ad-hoc fallback naming');
[
  'function MrrDashboard_getBoData_()',
  "totalDeptKey: 'BO'",
  "allLabel: 'BO全体'",
  'function MrrDashboard_readBoFcstSnapshots_(deptKeys)',
  'function MrrDashboard_readBoOppSnapshots_(deptKeys)',
  'getRange(1, 1, sheet.getLastRow(), 4).getValues()',
  'getRange(2, 1, sheet.getLastRow() - 1, OPP_HISTORY_V2_HEADERS.length).getValues()',
  'SharedAppState_isDepartmentTotal_({',
  'function MrrDashboard_addFcstSnapshotPayloadToBucket_(buckets, deptKey, dateStr, period, payload, nameRaw)',
  'function MrrDashboard_selectFcstSnapshotPayload_(bucket)',
  'function MrrDashboard_getFcstSnapshotRowKind_(payload, nameRaw, deptKey)',
  'function MrrDashboard_addFcstPayload_(sum, payload)',
  'var payload = MrrDashboard_selectFcstSnapshotPayload_(buckets[deptKey][dateStr][period]);',
  'function MrrDashboard_getBoMetricDefinitions_()',
  "key: 'fcstAdjusted'",
  "key: 'fcstMax'",
  'keyDealsData: keyDealsData',
  'MrrDashboard_pickSnapshotPeriodKeyFromKeys_',
  'MrrDashboard_formatLegacyDeals_'
].forEach((token) => {
  assertIncludes(mrrSnapshot, token, 'BO MRR snapshot path incomplete');
});

console.log('critical flow checks passed');
