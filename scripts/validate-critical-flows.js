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

const cacheLayer = read('CacheLayer.gs');
assertIncludes(cacheLayer, "var CACHE_PREFIX = 'fcst:';", 'cache prefix changed');
assertIncludes(cacheLayer, "'initData'", 'initData cache invalidation missing');
assertIncludes(cacheLayer, "'oppList'", 'oppList cache invalidation missing');
assertIncludes(cacheLayer, 'persistToSheet === false', 'ephemeral cache option missing');

const appDataCache = read('AppDataCache.gs');
assertIncludes(appDataCache, "CacheLayer_read(deptKey, 'initData', { skipSharedSheet: true })", 'initData cache read path missing');
assertIncludes(appDataCache, "CacheLayer_read(deptKey, 'oppList', { skipSharedSheet: true })", 'oppList cache read path missing');
assertIncludes(appDataCache, "CacheLayer_write(deptKey, 'oppList', result, { persistToSheet: false })", 'oppList must stay ephemeral');

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

const sfDataReader = read('SfDataReader.gs');
['groupCode', 'totalKind', 'fcstMin', 'fcstMax'].forEach((token) => {
  assertIncludes(sfDataReader, token, 'shared member/metric shape incomplete');
});

const config = read('Config.gs');
assertIncludes(config, 'function isProposalProductsEnabled_(deptKey)', 'proposalProducts helper missing');

const manifest = read('appsscript.json');
assertIncludes(manifest, '"https://www.googleapis.com/auth/drive"', 'snapshot DB folder writes require Apps Script DriveApp scope');

const snapshotStorage = read('SnapshotStorage.gs');
assertIncludes(snapshotStorage, 'function SnapshotStorage_getDbFolder_', 'snapshot DB source folder resolver missing');
assertIncludes(snapshotStorage, 'DriveApp.getFileById(SPREADSHEET_ID)', 'snapshot DB files must resolve the source spreadsheet folder');
assertIncludes(snapshotStorage, '.moveTo(dbFolder)', 'snapshot DB files must be moved to the source spreadsheet folder');
if (snapshotStorage.includes('googleapis.com/drive/v3')) {
  throw new Error('SnapshotStorage must avoid Drive API enablement dependency');
}

const mrrDashboard = read('MrrDashboard.gs');
assertIncludes(mrrDashboard, 'CacheLayer_read(MRR_DASHBOARD_CACHE_DEPT, MRR_DASHBOARD_CACHE_KEY, { skipSharedSheet: true })', 'MRR dashboard must use 5-minute cache before snapshot reads');
assertIncludes(mrrDashboard, 'FcstSnapshot_getAllValues_(4)', 'MRR dashboard must read FCST snapshots through snapshot storage helper');
assertIncludes(mrrDashboard, 'OppListSnapshot_getAllValues_(5)', 'MRR dashboard must read Opp snapshots through snapshot storage helper');
assertIncludes(mrrDashboard, "source: 'snapshot'", 'MRR dashboard must report snapshot source');
assertIncludes(mrrDashboard, "MrrDashboard_divisionKey_(cfg)", 'MRR dashboard must map SSCS departments into the SS division');
assertIncludes(mrrDashboard, ".setTitle('MRR進捗ダッシュボード')", 'MRR dashboard server title must be valid UTF-8');
if (mrrDashboard.includes('MRR_SHEET_ID') || mrrDashboard.includes('SpreadsheetApp.openById(MRR_SHEET_ID)')) {
  throw new Error('MRR dashboard must not read the legacy fixed MRR spreadsheet');
}
if (mrrDashboard.includes('cfg.sfSheetKey || cfg.division')) {
  throw new Error('MRR dashboard must not prioritize sfSheetKey over division; SSCS belongs to SS');
}

const mrrClient = read('mrr-index.html');
assertIncludes(mrrClient, 'MRR進捗ダッシュボード', 'MRR dashboard title must be valid UTF-8');
assertIncludes(mrrClient, 'getMrrDashboardData()', 'MRR client must load backend dashboard data');
assertIncludes(mrrClient, 'Key Deal はありません', 'MRR client must render empty Key Deal state without parser fallback');

const code = read('Code.gs');
assertIncludes(code, 'AppDataCache_getInitData', 'Code.gs must use shared init cache');
assertIncludes(code, 'AppDataCache_getOpportunities', 'Code.gs must use shared opp cache');
assertIncludes(code, 'AssignmentMaster_getContext', 'Code.gs must use shared assignment context');

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
