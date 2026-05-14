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

function assertBefore(haystack, before, after, message) {
  const beforeIndex = haystack.indexOf(before);
  const afterIndex = haystack.indexOf(after);
  if (beforeIndex === -1) throw new Error(message + `: missing "${before}"`);
  if (afterIndex === -1) throw new Error(message + `: missing "${after}"`);
  if (beforeIndex > afterIndex) {
    throw new Error(message + `: "${before}" must appear before "${after}"`);
  }
}

function assertCount(haystack, needle, expected, message) {
  const count = haystack.split(needle).length - 1;
  if (count !== expected) {
    throw new Error(message + `: expected ${expected}, found ${count} for "${needle}"`);
  }
}

function assertMatches(haystack, pattern, message) {
  if (!pattern.test(haystack)) {
    throw new Error(message + `: missing pattern ${pattern}`);
  }
}

function getFunctionBody(source, name) {
  const signature = `function ${name}`;
  const start = source.indexOf(signature);
  if (start === -1) {
    throw new Error(`missing function ${name}`);
  }
  const open = source.indexOf('{', start);
  if (open === -1) {
    throw new Error(`missing function body for ${name}`);
  }

  let depth = 0;
  for (let i = open; i < source.length; i++) {
    if (source[i] === '{') depth++;
    if (source[i] === '}') depth--;
    if (depth === 0) return source.slice(open + 1, i);
  }
  throw new Error(`unterminated function body for ${name}`);
}

const cacheLayer = read('CacheLayer.gs');
assertIncludes(cacheLayer, "var CACHE_PREFIX = 'fcst:';", 'cache prefix changed');
assertIncludes(cacheLayer, "'initData'", 'initData cache invalidation missing');
assertIncludes(cacheLayer, "'oppList'", 'oppList cache invalidation missing');
assertIncludes(cacheLayer, 'persistToSheet === false', 'ephemeral cache option missing');
const cacheLayerRemove = getFunctionBody(cacheLayer, 'CacheLayer_remove');
assertIncludes(cacheLayerRemove, "key + ':chunks'", 'CacheLayer_remove must delete chunk count cache');
assertIncludes(cacheLayerRemove, "key + ':chunk:' + i", 'CacheLayer_remove must delete chunk body cache');
assertIncludes(cacheLayerRemove, 'cache.removeAll(keys)', 'CacheLayer_remove must remove base and chunk cache keys');

const appDataCache = read('AppDataCache.gs');
assertIncludes(appDataCache, 'CacheLayer_read(deptKey, cacheLayerKey, { skipSharedSheet: true })', 'initData cache read path missing');
assertIncludes(appDataCache, "'initData'", 'initData cache key literal missing');
assertIncludes(appDataCache, "'initData:fast'", 'initData fast scope cache key literal missing');
assertIncludes(appDataCache, "CacheLayer_read(deptKey, 'oppList', { skipSharedSheet: true })", 'oppList cache read path missing');
assertIncludes(appDataCache, "CacheLayer_write(deptKey, 'oppList', result, { persistToSheet: false })", 'oppList must stay ephemeral');
assertIncludes(
  getFunctionBody(appDataCache, 'AppDataCache_getInitData'),
  'if (cached) return cached;',
  'initData cache hit must not perform extra snapshot/opportunity reads'
);
if (getFunctionBody(appDataCache, 'AppDataCache_attachFcstKeyDeals_').includes('FcstSnapshot_attachSnapshotKeyDealsToData_')) {
  throw new Error('initData attachment must not attach historical snapshot Key Deals during initial load');
}

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
assertIncludes(aggregatedCache, 'FcstSnapshot_getLatestMembers(deptKey, snapshotOpts)', 'aggregated refresh must call FcstSnapshot_getLatestMembers with snapshotOpts (includeKeyDeals: false)');
assertIncludes(aggregatedCache, 'FcstSnapshot_getDataByDate(deptKey, result.snapshotDates[0], snapshotOpts)', 'aggregated refresh must call FcstSnapshot_getDataByDate with snapshotOpts (includeKeyDeals: false)');
assertMatches(aggregatedCache, /snapshotOpts[^;]*includeKeyDeals:\s*false/, 'snapshotOpts must include includeKeyDeals: false to prevent historical Key Deal loading');
if (getFunctionBody(aggregatedCache, 'AggregatedCache_refresh').includes('FcstSnapshot_attachCurrentKeyDealsToData_')) {
  throw new Error('AggregatedCache_refresh must not attach current Key Deals into shared aggregate cache');
}

const sfDataReader = read('SfDataReader.gs');
['groupCode', 'totalKind', 'fcstMin', 'fcstMax'].forEach((token) => {
  assertIncludes(sfDataReader, token, 'shared member/metric shape incomplete');
});

const config = read('Config.gs');
assertIncludes(config, 'function isProposalProductsEnabled_(deptKey)', 'proposalProducts helper missing');

const oppListReader = read('OppListReader.gs');
assertIncludes(
  getFunctionBody(oppListReader, 'OppListReader_toBoolean_'),
  'return value === true;',
  'Opp list Key Deal boolean parsing must accept only boolean true'
);

assertIncludes(
  getFunctionBody(sfDataReader, 'SfDataReader_toBoolean_'),
  'return value === true;',
  'FCST Key Deal boolean parsing must accept only boolean true'
);
const sfAggregated = getFunctionBody(sfDataReader, 'SfDataReader_getAggregated');
assertMatches(
  sfAggregated,
  /if \(isConfirmedPhase\) \{\s*SfDataReader_addBreakdownValue_\(metric\.confirmed,/s,
  'FCST confirmed metric must aggregate only rows where フェーズ_変換 is 確定'
);
if (sfAggregated.includes('metric.keyDeals.push')) {
  throw new Error('FCST aggregation must not build Key Deal payloads from FCST rows');
}

const manifest = read('appsscript.json');
assertIncludes(manifest, '"https://www.googleapis.com/auth/drive"', 'snapshot DB folder writes require Apps Script DriveApp scope');

const snapshotStorage = read('SnapshotStorage.gs');
assertIncludes(snapshotStorage, 'function SnapshotStorage_getDbFolder_', 'snapshot DB source folder resolver missing');
assertIncludes(snapshotStorage, 'DriveApp.getFileById(SPREADSHEET_ID)', 'snapshot DB files must resolve the source spreadsheet folder');
assertIncludes(snapshotStorage, '.moveTo(dbFolder)', 'snapshot DB files must be moved to the source spreadsheet folder');
assertIncludes(getFunctionBody(snapshotStorage, 'SnapshotStorage_getReadSheets_'), 'SnapshotStorage_getReadFileIds_(sheetName)', 'snapshot reads must include registered and active DB files');
const snapshotReadFileIds = getFunctionBody(snapshotStorage, 'SnapshotStorage_getReadFileIds_');
assertIncludes(snapshotReadFileIds, 'SnapshotStorage_getFileIds_(sheetName)', 'snapshot read file ID list must include registered DB files');
assertIncludes(snapshotReadFileIds, 'SnapshotStorage_getActiveFileId_(sheetName)', 'snapshot read file ID list must include active DB file');
assertIncludes(getFunctionBody(snapshotStorage, 'SnapshotStorage_buildWriteResult_'), 'SnapshotStorage_registerFileId_(sheet.getName(), ss.getId())', 'snapshot writes must register the active DB file ID');
if (snapshotStorage.includes('googleapis.com/drive/v3')) {
  throw new Error('SnapshotStorage must avoid Drive API enablement dependency');
}

const mrrDashboard = read('MrrDashboard.gs');
assertIncludes(mrrDashboard, 'CacheLayer_read(MRR_DASHBOARD_CACHE_DEPT, MRR_DASHBOARD_CACHE_KEY, { skipSharedSheet: true })', 'MRR dashboard must use 5-minute cache before snapshot reads');
assertIncludes(mrrDashboard, 'FcstSnapshot_getAllValues_(4)', 'MRR dashboard must read FCST snapshots through snapshot storage helper');
assertIncludes(mrrDashboard, 'OppListSnapshot_getAllValues_(5)', 'MRR dashboard must read Opp snapshots through snapshot storage helper');
assertIncludes(mrrDashboard, "source: 'snapshot+current'", 'MRR dashboard must report snapshot + current source');
assertIncludes(mrrDashboard, "var MRR_DASHBOARD_CACHE_KEY = 'snapshotLiveData:v1';", 'MRR dashboard live payload must use a distinct cache key');
assertIncludes(mrrDashboard, "var MRR_DASHBOARD_LIVE_KEY = 'live';", 'MRR dashboard live key missing');
assertIncludes(mrrDashboard, 'AppDataCache_getInitData(deptKey)', 'MRR dashboard current FCST source must use AppDataCache_getInitData');
assertIncludes(mrrDashboard, 'AppDataCache_getOpportunities(oppDeptKey)', 'MRR dashboard current Key Deal source must use current opp list');
assertIncludes(mrrDashboard, 'periodDeals: {}', 'MRR dashboard must expose period-scoped Key Deals');
assertIncludes(mrrDashboard, 'MrrDashboard_dealMatchesPeriod_(deal, periodKey)', 'MRR dashboard Key Deals must be filtered by selected period');
assertIncludes(mrrDashboard, 'oppDeptKeys: []', 'MRR dashboard must keep explicit Opp group-name mappings');
assertIncludes(mrrDashboard, 'meta.oppDeptKeys.push(row.groupName)', 'MRR dashboard must derive current Opp source keys from OrgMaster groupName');
assertIncludes(mrrDashboard, 'MrrDashboard_buildLiveMetric_(live, member, monthKey)', 'MRR dashboard live metrics must merge current FCST fields explicitly');
assertIncludes(mrrDashboard, 'MrrDashboard_getLiveAdjustedMetric_', 'MRR dashboard live FCST adjusted metric merge missing');
assertIncludes(mrrDashboard, 'FcstPeriods_getQuarterKeyFromMonthKey_(periodKey)', 'MRR dashboard must derive quarter rows from monthly periods');
assertIncludes(mrrDashboard, "MrrDashboard_divisionKey_(cfg)", 'MRR dashboard must map SSCS departments into the SS division');
assertIncludes(mrrDashboard, ".setTitle('MRR進捗ダッシュボード')", 'MRR dashboard server title must be valid UTF-8');
assertIncludes(mrrDashboard, "var MRR_DASHBOARD_TOTAL_KEY = 'total';", 'MRR dashboard total key must be Apps Script RPC-safe');
if (mrrDashboard.includes('__total__')) {
  throw new Error('MRR dashboard response must not include properties ending with "__"');
}
const mrrBuild = getFunctionBody(mrrDashboard, 'MrrDashboard_buildFromSnapshots_');
assertIncludes(mrrBuild, 'FcstSnapshot_parseRowName_(nameRaw)', 'MRR dashboard must recover department key from snapshot row name');
assertIncludes(mrrBuild, 'FcstSnapshot_normalizeMeta_(payload, rowDeptKey, rowInfo.name)', 'MRR dashboard must use normalized snapshot metadata');
assertIncludes(mrrBuild, 'var deptKey = rowDeptKey', 'MRR dashboard must use snapshot row prefix as the canonical department key');
assertIncludes(mrrBuild, 'meta.totalKind !== SHARED_TOTAL_KIND.DEPARTMENT', 'MRR dashboard must identify department totals through normalized metadata');
if (mrrBuild.includes('meta.dept || rowDeptKey') || mrrBuild.includes('metaDeptKey')) {
  throw new Error('MRR dashboard must not fallback between payload __meta.dept and snapshot row prefix');
}
assertIncludes(mrrDashboard, "var MRR_DASHBOARD_DIVISION_ORDER = ['SS', 'BO', 'CO'];", 'MRR dashboard must include CO as a first-class division');
assertIncludes(mrrDashboard, "if (sfSheetKey === 'CO') return 'CO';", 'MRR dashboard must classify CO departments');
assertIncludes(mrrDashboard, 'function MrrDashboard_invalidateCache_', 'MRR dashboard cache invalidation helper missing');
assertIncludes(mrrDashboard, "'snapshotData:v6'", 'MRR dashboard cache invalidation must remove the previous snapshot-only key');
assertIncludes(mrrBuild, 'key: dateKey', 'MRR dashboard must group snapshot weeks by date, not minute');
assertIncludes(mrrBuild, 'division.data[periodKey][dateKey]', 'MRR dashboard data buckets must use date keys');
assertIncludes(mrrBuild, 'if (!FcstPeriods_parseMonthKey_(periodKey)) return;', 'MRR dashboard must ignore persisted quarter snapshot rows');
if (mrrBuild.includes("var timestampKey = Utilities.formatDate(snapshotAt, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm')")) {
  throw new Error('MRR dashboard snapshot grouping must not depend on minute-level timestamps');
}
assertMatches(
  getFunctionBody(mrrDashboard, 'getMrrDashboardData'),
  /(data\.divisions\s*&&\s*data\.divisions\.length|Array\.isArray\(data\.divisions\)\s*&&\s*data\.divisions\.length)/,
  'MRR dashboard must not 5-minute-cache an empty divisions result'
);
const mrrMetricValue = getFunctionBody(mrrDashboard, 'MrrDashboard_metricValue_');
assertIncludes(mrrMetricValue, 'MRR_DASHBOARD_NUMERIC_METRIC_KEYS[metricKey]', 'MRR dashboard must explicitly separate numeric metrics from net-object metrics');
assertIncludes(mrrMetricValue, "typeof value.net === 'number'", 'MRR dashboard must only use numeric net values for net-object metrics');
if (mrrMetricValue.includes('value && value.net') || mrrMetricValue.includes('Number(value.net)')) {
  throw new Error('MRR dashboard must not loosely fallback to arbitrary net values');
}
if (mrrDashboard.includes('MRR_SHEET_ID') || mrrDashboard.includes('SpreadsheetApp.openById(MRR_SHEET_ID)')) {
  throw new Error('MRR dashboard must not read the legacy fixed MRR spreadsheet');
}
if (mrrDashboard.includes('cfg.sfSheetKey || cfg.division')) {
  throw new Error('MRR dashboard must not prioritize sfSheetKey over division; SSCS belongs to SS');
}
if (mrrDashboard.includes('AggregatedCache_read(deptKey) ||') || mrrDashboard.includes('AppDataCache_getInitData(deptKey) ||')) {
  throw new Error('MRR dashboard must not add implicit live-source fallbacks');
}

const fcstSnapshot = read('FcstSnapshot.gs');
const fcstSnapshotCreateAt = getFunctionBody(fcstSnapshot, 'FcstSnapshot_createAt_');
assertIncludes(fcstSnapshot, 'function FcstSnapshot_buildPayloadMeta_', 'FCST snapshot must centralize write metadata shape');
const fcstPayloadMeta = getFunctionBody(fcstSnapshot, 'FcstSnapshot_buildPayloadMeta_');
assertIncludes(fcstSnapshotCreateAt, 'payload.__meta = FcstSnapshot_buildPayloadMeta_(member, options)', 'FCST snapshot writes must use normalized minimal metadata');
assertIncludes(fcstSnapshotCreateAt, 'var periods = FcstSnapshot_filterSnapshotPeriodKeys_(periodKeys || [])', 'FCST snapshot writes must filter persisted periods');
if (fcstSnapshotCreateAt.includes("k === 'keyDeals'")) {
  throw new Error('FCST snapshot payload must not persist keyDeals');
}
assertIncludes(fcstSnapshot, 'function FcstSnapshot_filterSnapshotPeriodKeys_', 'FCST snapshot monthly-only period filter missing');
assertIncludes(
  getFunctionBody(fcstSnapshot, 'FcstSnapshot_filterSnapshotPeriodKeys_'),
  'FcstPeriods_parseMonthKey_(key)',
  'FCST snapshot persisted periods must be monthly keys only'
);
assertIncludes(fcstSnapshot, 'function FcstSnapshot_mergeAdjustedIntoMembers_', 'FCST snapshot must merge adjusted values into snapshot payloads');
assertIncludes(fcstSnapshot, 'function FcstSnapshot_attachSnapshotKeyDealsToData_', 'FCST snapshot reads must attach historical Key Deals from Opp snapshots');
assertIncludes(fcstSnapshot, 'function FcstSnapshot_attachCurrentKeyDealsToData_', 'FCST current reads must attach current Key Deals from current Opp list');
assertIncludes(fcstSnapshot, 'OppListSnapshot_getByDate(oppDeptKey, dateKey)', 'Historical FCST Key Deals must use Opp list snapshots');
assertIncludes(fcstSnapshot, 'AppDataCache_getOpportunities(oppDeptKey)', 'Current FCST Key Deals must use current Opp list');
assertIncludes(
  getFunctionBody(fcstSnapshot, 'FcstSnapshot_getDataByTimestampKey_'),
  'options.includeKeyDeals !== false',
  'FCST snapshot full reads must support skipping historical Key Deal attachment'
);
assertIncludes(
  getFunctionBody(fcstSnapshot, 'FcstSnapshot_buildSnapshotInputFromLive_'),
  'FcstSnapshot_mergeAdjustedIntoMembers_(members, liveData && liveData.fcstAdjusted, periodKeys)',
  'FCST snapshot input must include live adjusted metrics'
);
assertIncludes(fcstSnapshotCreateAt, 'FcstSnapshot_getLatestMetricMap_(deptKey, dateKey)', 'FCST snapshot week-over-week baseline must exclude the same snapshot date');
assertIncludes(fcstSnapshotCreateAt, 'FcstSnapshot_deleteByDate_(deptKey, sheet, dateKey)', 'FCST snapshot writes must replace same dept/date rows before appending');
assertIncludes(fcstSnapshot, 'function FcstSnapshot_deleteByDate_', 'FCST snapshot must have date-level dedupe deletion');
assertIncludes(fcstSnapshot, 'function FcstSnapshot_getLatestDateKey_', 'FCST snapshot must resolve previous snapshots by date');
if (fcstSnapshot.includes('function FcstSnapshot_hasRowAt_')) {
  throw new Error('FCST snapshot creation must not use minute-level duplicate checks');
}
assertIncludes(
  getFunctionBody(fcstSnapshot, 'FcstSnapshot_getTrendData'),
  'snapshotMap[dateKey]',
  'FCST trend data must collapse snapshots by date'
);
assertIncludes(
  getFunctionBody(fcstSnapshot, 'FcstSnapshot_getTrendData'),
  'FcstSnapshot_getSnapshotKeyDealsForPeriod_(deptKey, dateKey, targetPeriod)',
  'FCST trend historical Key Deals must come from Opp snapshots'
);
assertIncludes(
  getFunctionBody(fcstSnapshot, 'FcstSnapshot_getTrendData'),
  'FcstSnapshot_getCurrentKeyDealsForPeriod_(deptKey, livePeriod)',
  'FCST trend current Key Deals must come from current Opp list'
);
assertIncludes(
  getFunctionBody(fcstSnapshot, 'FcstSnapshot_getTrendWeekDetails'),
  "Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd') !== snapshotKey",
  'FCST trend detail lookup must use date keys'
);
if (getFunctionBody(fcstSnapshot, 'FcstSnapshot_getTrendWeekDetails').includes('matchedPayload.keyDeals') ||
    getFunctionBody(fcstSnapshot, 'FcstSnapshot_getTrendWeekDetails').includes('liveMetric && liveMetric.keyDeals')) {
  throw new Error('FCST trend details must not read Key Deals from FCST snapshot/live metric payload');
}
assertIncludes(fcstPayloadMeta, 'meta.totalKind = totalKind', 'FCST snapshot write metadata must include row type');
assertIncludes(fcstPayloadMeta, 'if (group) meta.group = group', 'FCST snapshot write metadata must include historical group when present');
['isTotal:', 'dept:', 'name:', 'captureMode:', 'groupCode:'].forEach((token) => {
  if (fcstPayloadMeta.includes(token)) {
    throw new Error('FCST snapshot write metadata must not persist redundant field: ' + token);
  }
});
assertIncludes(fcstSnapshot, 'function FcstSnapshot_parseRowName_', 'FCST snapshot reads must parse row identity from snapshot row name');
assertIncludes(fcstSnapshot, 'function FcstSnapshot_normalizeMeta_', 'FCST snapshot reads must normalize legacy and minimal metadata');
assertIncludes(fcstSnapshot, 'function FcstSnapshot_isDepartmentTotalRowForDept_', 'FCST snapshot reads must normalize department total rows');
assertIncludes(
  getFunctionBody(fcstSnapshot, 'FcstSnapshot_isDepartmentTotalRowForDept_'),
  'FcstSnapshot_parseRowName_(nameRaw)',
  'FCST snapshot reads must accept existing department total rows with blank meta.dept'
);
if (fcstSnapshot.includes("meta.totalKind !== 'department' || meta.dept !== deptKey")) {
  throw new Error('FCST snapshot reads must not reject existing rows with blank meta.dept');
}

const mrrClient = read('mrr-index.html');
assertIncludes(mrrClient, 'MRR進捗ダッシュボード', 'MRR dashboard title must be valid UTF-8');
assertIncludes(mrrClient, 'getMrrDashboardData()', 'MRR client must load backend dashboard data');
assertIncludes(mrrClient, 'Key Deal はありません', 'MRR client must render empty Key Deal state without parser fallback');
assertIncludes(mrrClient, "return (D && D.totalKey) || 'total';", 'MRR client total fallback key must be Apps Script RPC-safe');
assertIncludes(mrrClient, "w.isLive ? '現在'", 'MRR client must label live/current week as 現在');
assertIncludes(mrrClient, 'height: 560px !important', 'MRR dashboard chart must keep expanded height');
assertIncludes(mrrClient, '--color-primary: #1a73e8;', 'MRR dashboard tone must align to FCST dashboard primary color');
if (mrrClient.includes('__total__')) {
  throw new Error('MRR client must not depend on properties ending with "__"');
}

const guardrails = read('DEVELOPMENT_GUARDRAILS.md');
assertIncludes(guardrails, 'MRR Dashboard Rules', 'MRR dashboard guardrails missing');
assertIncludes(guardrails, 'FCST snapshot rows は月次キーのみ保存する', 'monthly-only snapshot guardrail missing');
assertIncludes(guardrails, 'historical の FCST/MRR Key Deal は `案件リストスナップショット` を正規ソースにする', 'historical Key Deal source guardrail missing');
assertIncludes(guardrails, '`FCSTスナップショット` payload に `keyDeals` を保存しない', 'FCST snapshot Key Deal payload prohibition missing');
assertIncludes(guardrails, 'HTML shell を返す `doGet` では、テンプレートに埋め込まないデータを先読みしない', 'doGet preload guardrail missing');
assertIncludes(guardrails, '初期表示用 `initData` では、過去snapshotのKey Deal付与を行わない', 'initData historical Key Deal guardrail missing');
assertIncludes(guardrails, 'version削除は、上限付近・上限到達・version作成失敗時に限って実施する', 'Apps Script version deletion rule must be limit/failure-only');

const code = read('Code.gs');
assertIncludes(code, 'AppDataCache_getInitData', 'Code.gs must use shared init cache');
assertIncludes(code, 'AppDataCache_getOpportunities', 'Code.gs must use shared opp cache');
assertIncludes(code, 'AssignmentMaster_getContext', 'Code.gs must use shared assignment context');
const doGetBody = getFunctionBody(code, 'doGet');
if (doGetBody.includes('AppDataCache_getInitData(deptKey)')) {
  throw new Error('doGet must not prefetch initData when embeddedInitData is not used');
}
assertIncludes(
  getFunctionBody(code, 'getClientSnapshotData'),
  'FcstSnapshot_getDataByDate(deptKey, dateStr, { includeKeyDeals: false })',
  'client snapshot period loads must avoid full historical Key Deal attachment'
);

const client = read('js.html');
['isDepartmentTotalMember_', 'isGroupTotalMember_', 'getMemberGroupLabel_'].forEach((token) => {
  assertIncludes(client, token, 'FCST client total handling incomplete');
});
assertCount(client, 'var ClientCache = {', 1, 'ClientCache definition count invalid');
assertBefore(client, 'var ClientCache = {', 'ClientCache.get(_selectedDept', 'ClientCache boot dependency is undefined or declared too late');
assertBefore(client, 'var ClientCache = {', "ClientCache.set(_selectedDept, 'initData'", 'ClientCache initData setter dependency is undefined or declared too late');
assertBefore(client, 'var ClientCache = {', "ClientCache.set(_selectedDept, 'oppList'", 'ClientCache oppList setter dependency is undefined or declared too late');
assertIncludes(client, 'set: function(deptKey, dataKey, data, updatedAt)', 'ClientCache set contract missing');
assertIncludes(client, 'get: function(deptKey, dataKey)', 'ClientCache get contract missing');
assertIncludes(client, 'TTL: 5 * 60 * 1000', 'ClientCache TTL must stay aligned to 5-minute display cache rule');

console.log('critical flow checks passed');
