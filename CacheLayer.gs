var CACHE_PREFIX = 'fcst:';
var CACHE_TTL_DEFAULT_SECONDS = 300;
var CACHE_TTL_SECONDS_BY_KEY = {
  initData: 300,
  oppList: 300,
  members: 300,
  fcstData: 300,
  targets: 300,
  adjustments: 300,
  snapshots: 300,
  notes: 300,
  sfLastUpdated: 300
};

function CacheLayer_read(deptKey, dataKey, options) {
  var opts = options || {};
  var cache = CacheService.getScriptCache();
  var key = CacheLayer_buildKey_(deptKey, dataKey);
  var ttlSeconds = CacheLayer_getTtl_(dataKey);
  var hit = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit); } catch(e) {}
  }
  var chunked = CacheLayer_readChunked_(cache, key);
  if (chunked) return chunked;
  if (opts.skipSharedSheet) return null;
  var ssData = AggregatedCache_read(deptKey);
  if (ssData && ssData[dataKey] !== undefined) {
    try {
      var json = JSON.stringify(ssData[dataKey]);
      if (json.length <= 100000) {
        cache.put(key, json, ttlSeconds);
      } else {
        CacheLayer_writeChunked_(cache, key, json, ttlSeconds);
      }
    } catch(e) {}
    return ssData[dataKey];
  }
  return null;
}

function CacheLayer_write(deptKey, dataKey, value, options) {
  var opts = options || {};
  var cache = CacheService.getScriptCache();
  var key = CacheLayer_buildKey_(deptKey, dataKey);
  var ttlSeconds = CacheLayer_getTtl_(dataKey);
  try {
    var json = JSON.stringify(value);
    if (json.length <= 100000) {
      cache.put(key, json, ttlSeconds);
    } else {
      CacheLayer_writeChunked_(cache, key, json, ttlSeconds);
    }
  } catch(e) {}
  if (opts.persistToSheet === false) return;
  try { AggregatedCache_writeKey(deptKey, dataKey, value); } catch(e) {}
}

function CacheLayer_buildKey_(deptKey, dataKey) {
  return CACHE_PREFIX + deptKey + ':' + dataKey;
}

function CacheLayer_getTtl_(dataKey) {
  return CACHE_TTL_SECONDS_BY_KEY[dataKey] || CACHE_TTL_DEFAULT_SECONDS;
}

function CacheLayer_writeChunked_(cache, baseKey, json, ttlSeconds) {
  var CHUNK = 90000;
  var chunks = Math.ceil(json.length / CHUNK);
  var entries = {};
  for (var i = 0; i < chunks; i++) {
    entries[baseKey + ':chunk:' + i] = json.substr(i * CHUNK, CHUNK);
  }
  entries[baseKey + ':chunks'] = String(chunks);
  try { cache.putAll(entries, ttlSeconds || CACHE_TTL_DEFAULT_SECONDS); } catch(e) {}
}

function CacheLayer_readChunked_(cache, key) {
  var n = cache.get(key + ':chunks');
  if (!n) return null;
  var parts = [];
  for (var i = 0; i < parseInt(n); i++) {
    var chunk = cache.get(key + ':chunk:' + i);
    if (!chunk) return null;
    parts.push(chunk);
  }
  try { return JSON.parse(parts.join('')); } catch(e) { return null; }
}

function CacheLayer_removeChunked_(cache, key) {
  var count = parseInt(cache.get(key + ':chunks') || '0', 10);
  var keys = [key + ':chunks'];
  for (var i = 0; i < count; i++) {
    keys.push(key + ':chunk:' + i);
  }
  try { cache.removeAll(keys); } catch (e) {}
}

function CacheLayer_invalidate(deptKey) {
  var cache = CacheService.getScriptCache();
  var keys = ['initData','members','fcstData','oppList','targets',
    'adjustments','snapshots','notes','sfLastUpdated'];
  var cacheKeys = keys.map(function(k) {
    return CacheLayer_buildKey_(deptKey, k);
  });
  cacheKeys.forEach(function(key) {
    CacheLayer_removeChunked_(cache, key);
  });
  cache.removeAll(cacheKeys);
}
