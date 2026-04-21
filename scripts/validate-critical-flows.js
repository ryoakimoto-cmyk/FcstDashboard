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

const code = read('Code.gs');
assertIncludes(code, 'AppDataCache_getInitData', 'Code.gs must use shared init cache');
assertIncludes(code, 'AppDataCache_getOpportunities', 'Code.gs must use shared opp cache');
assertIncludes(code, 'AssignmentMaster_getContext', 'Code.gs must use shared assignment context');
assertIncludes(code, 'function getEmbeddedInitDataForDept_(deptKey)', 'dept embedded init helper missing');
assertIncludes(code, "tmpl.embeddedInitData = getEmbeddedInitDataForDept_(deptKey);", 'dept page must use embedded init helper');
assertIncludes(code, 'return json.length <= 800000 ? json : \'null\';', 'embedded init size guard missing');

const client = read('js.html');
['isDepartmentTotalMember_', 'isGroupTotalMember_', 'getMemberGroupLabel_'].forEach((token) => {
  assertIncludes(client, token, 'FCST client total handling incomplete');
});
assertIncludes(client, 'TTL: 300000,', 'client cache TTL must stay at 5 minutes');
assertIncludes(client, 'ClientCache.shouldRevalidate(cached)', 'client should only revalidate stale caches');
assertIncludes(client, 'App.refreshInitDataSilently_({ data: _embeddedInitData });', 'embedded init data must refresh in background');
assertIncludes(client, 'App.refreshInitDataSilently_(cached);', 'cached init data must refresh in background');
assertIncludes(client, 'function startInitLoadTimeout_()', 'FCST init timeout helper missing');
assertIncludes(client, 'clearInitLoadTimeout_();', 'FCST init timeout clear missing');
assertIncludes(client, 'function getSharedQuarterDefinitions_(periodOptions)', 'shared quarter definitions missing');
assertCount(client, 'function isDepartmentTotalMember_(member)', 1, 'duplicate department total helper');
assertCount(client, 'function isGroupTotalMember_(member)', 1, 'duplicate group total helper');
assertCount(client, 'function getMemberGroupLabel_(member)', 1, 'duplicate group label helper');
assertCount(client, 'function displayMemberName(member)', 1, 'duplicate member display helper');
assertCount(client, 'function renderTopPage_()', 1, 'duplicate top page renderer');

const mrr = read('mrr-index.html');
['MRR進捗ダッシュボード', '読込中...', '全部署', 'getMrrDashboardData();'].forEach((token) => {
  assertIncludes(mrr, token, 'MRR template markers missing');
});

console.log('critical flow checks passed');
