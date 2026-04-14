var CACHE_TTL = 3600;
var CACHE_PREFIX = 'fcst:';

function CacheLayer_read(deptKey, dataKey, options) {
  var opts = options || {};
  var cache = CacheService.getScriptCache();
  var key = CacheLayer_buildKey_(deptKey, dataKey);
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
        cache.put(key, json, CACHE_TTL);
      } else {
        CacheLayer_writeChunked_(cache, key, json);
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
  try {
    var json = JSON.stringify(value);
    if (json.length <= 100000) {
      cache.put(key, json, CACHE_TTL);
    } else {
      CacheLayer_writeChunked_(cache, key, json);
    }
  } catch(e) {}
  if (opts.persistToSheet === false) return;
  try { AggregatedCache_writeKey(deptKey, dataKey, value); } catch(e) {}
}

function CacheLayer_buildKey_(deptKey, dataKey) {
  return CACHE_PREFIX + deptKey + ':' + dataKey;
}

function CacheLayer_writeChunked_(cache, baseKey, json) {
  var CHUNK = 90000;
  var chunks = Math.ceil(json.length / CHUNK);
  var entries = {};
  for (var i = 0; i < chunks; i++) {
    entries[baseKey + ':chunk:' + i] = json.substr(i * CHUNK, CHUNK);
  }
  entries[baseKey + ':chunks'] = String(chunks);
  try { cache.putAll(entries, CACHE_TTL); } catch(e) {}
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

function CacheLayer_invalidate(deptKey) {
  var cache = CacheService.getScriptCache();
  var keys = ['initData','members','fcstData','oppList','targets',
    'adjustments','snapshots','notes','sfLastUpdated'];
  cache.removeAll(keys.map(function(k) {
    return CacheLayer_buildKey_(deptKey, k);
  }));
}
