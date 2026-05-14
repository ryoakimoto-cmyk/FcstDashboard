const FCST_FISCAL_BASE_YEAR = 2025;
const FCST_FISCAL_BASE_MONTH = 6;

function FcstPeriods_isSupportedDate_(date) {
  if (!(date instanceof Date) || isNaN(date)) return false;
  var base = new Date(FCST_FISCAL_BASE_YEAR, FCST_FISCAL_BASE_MONTH - 1, 1);
  var target = new Date(date.getFullYear(), date.getMonth(), 1);
  return target >= base;
}

// Returns month keys around `now`: `before` months before and `after` months
// after, plus the current month. Months earlier than the fiscal base are
// dropped (no data for them anyway).
function FcstPeriods_getActiveMonthsAroundNow_(now, before, after) {
  var base = (now instanceof Date && !isNaN(now)) ? now : new Date();
  var b = Math.max(0, Number(before) || 0);
  var a = Math.max(0, Number(after) || 0);
  var year = base.getFullYear();
  var month = base.getMonth() + 1; // 1-12
  var months = [];
  for (var offset = -b; offset <= a; offset++) {
    var monthKey = FcstPeriods_addMonthsToMonthKey_(year, month, offset);
    if (!FcstPeriods_getQuarterKeyFromMonthKey_(monthKey)) continue; // before fiscal base
    months.push(monthKey);
  }
  return months;
}

// Quarter keys that contain any of the given month keys.
function FcstPeriods_getQuarterKeysForMonths_(monthKeys) {
  var seen = {};
  var result = [];
  (monthKeys || []).forEach(function(monthKey) {
    var qKey = FcstPeriods_getQuarterKeyFromMonthKey_(monthKey);
    if (!qKey || seen[qKey]) return;
    seen[qKey] = true;
    result.push(qKey);
  });
  return result;
}

function FcstPeriods_formatMonthKey_(date) {
  return date.getFullYear() + '-' + String(date.getMonth() + 1).padStart(2, '0');
}

function FcstPeriods_buildDefinitionsFromMonthKeys_(monthKeys) {
  var quarterMap = {};
  (monthKeys || []).forEach(function(monthKey) {
    if (!FcstPeriods_parseMonthKey_(monthKey)) return;
    var quarterKey = FcstPeriods_getQuarterKeyFromMonthKey_(monthKey);
    if (!quarterKey || quarterMap[quarterKey]) return;
    quarterMap[quarterKey] = FcstPeriods_getQuarterDefinitionByKey_(quarterKey);
  });

  return Object.keys(quarterMap)
    .sort(function(a, b) { return a < b ? -1 : a > b ? 1 : 0; })
    .map(function(key) { return quarterMap[key]; });
}

function FcstPeriods_getQuarterDefinitionByKey_(quarterKey) {
  var match = String(quarterKey || '').match(/^(\d+)Q([1-4])$/);
  if (!match) return null;
  var fiscal = Number(match[1]);
  var quarter = Number(match[2]);
  var baseOffset = ((fiscal - 19) * 12) + ((quarter - 1) * 3);
  var months = [];
  for (var i = 0; i < 3; i++) {
    months.push(FcstPeriods_addMonthsToMonthKey_(FCST_FISCAL_BASE_YEAR, FCST_FISCAL_BASE_MONTH, baseOffset + i));
  }
  return {
    key: quarterKey,
    label: fiscal + '\u671f' + quarter + 'Q',
    months: months
  };
}

function FcstPeriods_getQuarterKeyFromMonthKey_(monthKey) {
  var parsed = FcstPeriods_parseMonthKey_(monthKey);
  if (!parsed) return '';
  var offset = (parsed.year - FCST_FISCAL_BASE_YEAR) * 12 + (parsed.month - FCST_FISCAL_BASE_MONTH);
  if (offset < 0) return '';
  var fiscal = 19 + Math.floor(offset / 12);
  var quarter = Math.floor((offset % 12) / 3) + 1;
  return fiscal + 'Q' + quarter;
}

function FcstPeriods_parseMonthKey_(monthKey) {
  var match = String(monthKey || '').match(/^(\d{4})-(\d{2})$/);
  if (!match) return null;
  return { year: Number(match[1]), month: Number(match[2]) };
}

function FcstPeriods_addMonthsToMonthKey_(year, month, offset) {
  var total = (year * 12 + (month - 1)) + offset;
  var y = Math.floor(total / 12);
  var m = (total % 12) + 1;
  return y + '-' + String(m).padStart(2, '0');
}

function FcstPeriods_expandKeys_(periodOptions) {
  var result = [];
  (periodOptions || []).forEach(function(option) {
    if (!option || !option.key) return;
    result.push(option.key);
    (option.months || []).forEach(function(monthKey) {
      result.push(monthKey);
    });
  });
  return result;
}
