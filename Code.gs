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
    var notes = NotesSheet_getNotes();
    result.notes = notes;
    return { data: result };
  } catch (e) { return { error: e.message }; }
}
function getTrendData(block) { try { return FcstReader_getTrendData(block); } catch (e) { return { error: e.message }; } }
function getOpportunities() { try { return OppReader_getOpportunities(); } catch (e) { return { error: e.message }; } }
function getSummaryData() { try { return SummaryReader_getSummaryData(); } catch (e) { return { error: e.message }; } }
function saveFcstAdjusted(p) { try { return FcstWriter_saveFcstAdjusted(p); } catch (e) { return { error: e.message }; } }
function saveOppSfValue(p) { try { return FcstWriter_saveOppSfValue(p); } catch (e) { return { error: e.message }; } }
function saveNote(p) { try { return NotesSheet_saveNote(p); } catch (e) { return { error: e.message }; } }
