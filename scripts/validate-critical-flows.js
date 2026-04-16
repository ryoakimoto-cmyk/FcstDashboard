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

const client = read('js.html');
['isDepartmentTotalMember_', 'isGroupTotalMember_', 'getMemberGroupLabel_'].forEach((token) => {
  assertIncludes(client, token, 'FCST client total handling incomplete');
});
assertIncludes(client, 'TTL: 300000,', 'client cache TTL must stay at 5 minutes');
assertIncludes(client, 'ClientCache.shouldRevalidate(cached)', 'client should only revalidate stale caches');
assertIncludes(client, 'function getSharedQuarterDefinitions_(periodOptions)', 'shared quarter definitions missing');

console.log('critical flow checks passed');
