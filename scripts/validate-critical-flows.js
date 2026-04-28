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
assertIncludes(manifest, '"https://www.googleapis.com/auth/drive.file"', 'snapshot DB folder writes must use drive.file scope');
if (manifest.includes('"https://www.googleapis.com/auth/drive"')) {
  throw new Error('snapshot DB folder writes must not require full Drive scope');
}

const snapshotStorage = read('SnapshotStorage.gs');
assertIncludes(snapshotStorage, 'function SnapshotStorage_createSpreadsheetInDbFolder_', 'snapshot DB folder creation path missing');
assertIncludes(snapshotStorage, 'https://www.googleapis.com/drive/v3/files?supportsAllDrives=true', 'snapshot DB file creation must target Drive API folder path');
assertIncludes(snapshotStorage, 'parents: [folderId]', 'snapshot DB files must be created directly under configured folder');
if (snapshotStorage.includes('DriveApp.')) {
  throw new Error('SnapshotStorage must avoid DriveApp full-drive authorization dependency');
}

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
