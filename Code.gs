function doGet(e) {
  const email = Session.getActiveUser().getEmail();
  if (!email.endsWith('@sansan.com')) {
    return HtmlService.createHtmlOutput('<h1>アクセス権限がありません</h1>');
  }
  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('FCST Dashboard - BOAM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function diagnoseInitData() {
  const result = getInitData();
  if (result.error) { Logger.log('ERROR: ' + result.error); return; }
  const d = result.data;
  Logger.log('lastUpdated: ' + d.lastUpdated);
  Logger.log('members count: ' + d.members.length);
  d.members.slice(0, 5).forEach(function(m) {
    Logger.log(m.name + ' [' + m.group + '] isTotal=' + m.isTotal + ' Q.fcstCommit=' + m.Q.fcstCommit + ' Q.target=' + m.Q.target);
  });
}

function diagnoseSourceSchemas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ['SFデータ更新', '目標', 'SFユーザー'].forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { Logger.log(name + ': シートが見つかりません'); return; }
    const nCols = sheet.getLastColumn();
    const rows = sheet.getRange(1, 1, Math.min(4, sheet.getLastRow()), nCols).getValues();
    Logger.log('=== ' + name + ' (' + nCols + '列, ' + sheet.getLastRow() + '行) ===');
    rows.forEach(function(row, i) {
      Logger.log('row' + (i + 1) + ': ' + JSON.stringify(row.map(function(v) {
        return v instanceof Date ? v.toISOString().slice(0, 10) : v;
      })));
    });
  });
}

function getSnapshotList() { try { return FcstReader_getSnapshotList(); } catch (e) { return { error: e.message }; } }
function getFcstData(snapIdx) { try { return FcstReader_getFcstData(snapIdx); } catch (e) { return { error: e.message }; } }
// 初期ロード用: スナップショット一覧 + 最新データを1回のサーバーコールで返す
function getInitData() {
  try {
    var users = UserReader_getUsers();
    var targets = TargetReader_getTargets();
    var result = SfDataReader_getAggregated(users, targets);
    var fcstState = FcstAdjusted_getState();
    result.notes = fcstState.notes;
    result.fcstAdjusted = fcstState.adjusted;
    result.weekOverWeekMap = FcstSnapshot_getWeekOverWeek();
    result.snapshotDates = FcstSnapshot_getSnapshotDates();
    result.previousSnapshot = FcstSnapshot_getLatestMembers();
    return { data: result };
  } catch (e) { return { error: e.message }; }
}
function getTrendData(block) { try { return FcstReader_getTrendData(block); } catch (e) { return { error: e.message }; } }
function getOpportunities() {
  try {
    var result = OppListReader_getLiveRows();
    result.snapshotDates = OppListSnapshot_getSnapshotDates();
    result.previousRows = result.snapshotDates.length
      ? OppListSnapshot_getByDate(result.snapshotDates[0]).rows
      : [];
    return result;
  } catch (e) { return { error: e.message }; }
}
function getOppSnapshotData(dateStr) { try { return OppListSnapshot_getByDate(dateStr); } catch (e) { return { error: e.message }; } }
function getSummaryData() { try { return SummaryReader_getSummaryData(); } catch (e) { return { error: e.message }; } }
function saveFcstAdjusted(p) { try { return FcstWriter_saveFcstAdjusted(p); } catch (e) { return { error: e.message }; } }
function saveOppSfValue(p) { try { return OppListWriter_saveDrafts([p]); } catch (e) { return { error: e.message }; } }
function saveOppDrafts(changes) { try { return OppListWriter_saveDrafts(changes); } catch (e) { return { error: e.message }; } }
function saveNote(p) { try { return FcstAdjusted_save(p); } catch (e) { return { error: e.message }; } }
function saveFcstAdjusted2(p) { try { return FcstAdjusted_save(p); } catch (e) { return { error: e.message }; } }
function getSnapshotDates() { try { return FcstSnapshot_getSnapshotDates(); } catch (e) { return { error: e.message }; } }
function getSnapshotData(dateStr) { try { return FcstSnapshot_getDataByDate(dateStr); } catch (e) { return { error: e.message }; } }
function createOppSnapshot() { try { return OppListSnapshot_createWeekly(); } catch (e) { return { error: e.message }; } }
function setupOppSnapshotTrigger() { try { return OppListSnapshot_setupWeeklyTrigger(); } catch (e) { return { error: e.message }; } }
function createSnapshot() {
  try {
    var users = UserReader_getUsers();
    var targets = TargetReader_getTargets();
    var result = SfDataReader_getAggregated(users, targets);
    var fcstState = FcstAdjusted_getState();
    var fcstAdj = fcstState.adjusted;
    var notes = fcstState.notes;
    var periods = ['Q', 'M5', 'M6', 'M7'];
    result.members.forEach(function(member) {
      periods.forEach(function(p) {
        var key = member.name + '|' + p;
        if (!member[p]) member[p] = {};
        member[p].fcstAdjusted = fcstAdj[key] || { net: 0, newExp: 0, churn: 0 };
      });
    });
    return FcstSnapshot_create(result.members, notes);
  } catch (e) { return { error: e.message }; }
}
