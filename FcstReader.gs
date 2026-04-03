function FcstReader_getSnapshotList() {
  const sheet = FcstReader_getFcstSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < SNAPSHOT_START_ROW) return [];

  const numRows = lastRow - SNAPSHOT_START_ROW + 1;
  // 日付はcol A（index 0）に入っている
  const values = sheet.getRange(SNAPSHOT_START_ROW, 1, numRows, 1).getValues();
  const snapshots = [];

  for (let offset = 0; offset < values.length; offset += ROWS_PER_SNAPSHOT) {
    const value = values[offset] ? values[offset][0] : '';
    if (!FcstReader_isSnapshotValue_(value)) continue;
    snapshots.push({
      index: snapshots.length,
      date: FcstReader_formatDate_(value),
      row: SNAPSHOT_START_ROW + offset,
    });
  }
  return snapshots;
}

function FcstReader_getFcstData(snapIdx) {
  const snapshots = FcstReader_getSnapshotList();
  if (!snapshots.length) throw new Error('スナップショットが見つかりません');
  if (snapIdx < 0 || snapIdx >= snapshots.length) throw new Error('存在しないスナップショットです');

  const sheet = FcstReader_getFcstSheet_();
  const startRow = SNAPSHOT_START_ROW + snapIdx * ROWS_PER_SNAPSHOT;
  const values = sheet.getRange(startRow, 1, ROWS_PER_SNAPSHOT, BLOCK_COL_STARTS.M7 + 34).getValues();
  const members = [];

  values.forEach(function(row) {
    const label = String(row[BLOCK_COL_STARTS.Q - 1] || '').trim();
    if (!label) return;

    const member = {
      name: label,
      isTotal: label.indexOf('全体') !== -1 || label.indexOf('合計') !== -1,
      Q: FcstReader_extractBlock_(row, BLOCK_COL_STARTS.Q),
      M5: FcstReader_extractBlock_(row, BLOCK_COL_STARTS.M5),
      M6: FcstReader_extractBlock_(row, BLOCK_COL_STARTS.M6),
      M7: FcstReader_extractBlock_(row, BLOCK_COL_STARTS.M7),
    };
    members.push(member);
  });

  return {
    date: snapshots[snapIdx].date,
    members: members,
  };
}

function FcstReader_getTrendData(block) {
  const targetBlock = BLOCK_COL_STARTS[block] ? block : 'Q';
  const snapshots = FcstReader_getSnapshotList().slice(0, 6);
  if (!snapshots.length) return { weeks: [], dates: [], series: { target: [], fcstAdjusted: [], fcstCommit: [], received: [] } };

  const sheet = FcstReader_getFcstSheet_();
  const lastSnapIdx = snapshots[snapshots.length - 1].index;
  const totalRows = (lastSnapIdx + 1) * ROWS_PER_SNAPSHOT;
  const colCount = BLOCK_COL_STARTS.M7 + 34;

  // 6スナップショット分を1回のgetValues()で一括取得
  const allValues = sheet.getRange(SNAPSHOT_START_ROW, 1, totalRows, colCount).getValues();

  const data = { weeks: [], dates: [], series: { target: [], fcstAdjusted: [], fcstCommit: [], received: [] } };

  snapshots.reverse().forEach(function(snapshot) {
    const offset = snapshot.index * ROWS_PER_SNAPSHOT;
    // 全体行を探す
    for (let r = 0; r < ROWS_PER_SNAPSHOT; r++) {
      const row = allValues[offset + r];
      if (!row) continue;
      const label = String(row[BLOCK_COL_STARTS.Q - 1] || '').trim();
      if (label.indexOf('全体') === -1 && label.indexOf('合計') === -1) continue;

      const blockData = FcstReader_extractBlock_(row, BLOCK_COL_STARTS[targetBlock]);
      data.dates.push(snapshot.date);
      data.weeks.push('W' + Utilities.formatDate(new Date(snapshot.date), Session.getScriptTimeZone(), 'ww'));
      data.series.target.push(FcstReader_toNumber_(blockData.target.net));
      data.series.fcstAdjusted.push(FcstReader_toNumber_(blockData.fcstAdjusted.net));
      data.series.fcstCommit.push(FcstReader_toNumber_(blockData.fcstCommit.net));
      data.series.received.push(FcstReader_toNumber_(blockData.received.net));
      break;
    }
  });

  return data;
}

// GASエディタから直接実行して動作確認できる診断関数
function FcstReader_diagnose() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  Logger.log('=== シート一覧 ===');
  sheets.forEach(function(s) { Logger.log(s.getName()); });

  const fcstSheet = sheets.find(function(s) { return s.getName() === FCST_SHEET_NAME; });
  if (!fcstSheet) { Logger.log('FCSTシートが見つかりません。FCST_SHEET_NAMEを確認してください。'); return; }
  Logger.log('FCSTシート: ' + fcstSheet.getName());

  // row4のA〜E列を確認
  const sample = fcstSheet.getRange(SNAPSHOT_START_ROW, 1, 1, 5).getValues()[0];
  Logger.log('row' + SNAPSHOT_START_ROW + ' A〜E: ' + JSON.stringify(sample));
  Logger.log('col A の型: ' + Object.prototype.toString.call(sample[0]));
  Logger.log('isSnapshotValue: ' + FcstReader_isSnapshotValue_(sample[0]));

  const list = FcstReader_getSnapshotList();
  Logger.log('スナップショット数: ' + list.length);
  if (list.length) Logger.log('最新: ' + JSON.stringify(list[0]));

  // 全データ行のうち「空でない行」を探して最初の50件を表示（列はA〜M = 13列）
  Logger.log('=== 非空行ダンプ (row1〜200, cols A〜M) ===');
  const totalRows = Math.min(fcstSheet.getLastRow(), 200);
  const allRows = fcstSheet.getRange(1, 1, totalRows, 13).getValues();
  let found = 0;
  allRows.forEach(function(row, i) {
    const rowNum = i + 1;
    const hasData = row.some(function(v) { return v !== '' && v !== 0 && v !== false && v !== null; });
    if (!hasData) return;
    if (found >= 50) return;
    found++;
    const preview = row.map(function(v) {
      if (v instanceof Date) return v.toISOString().slice(0,10);
      return JSON.stringify(v);
    }).join(', ');
    Logger.log('row' + rowNum + ': [' + preview + ']');
  });
  Logger.log('非空行合計: ' + found + '件 (最大50件表示)');
}

function FcstReader_getFcstSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(FCST_SHEET_NAME);
  if (!sheet) throw new Error('FCSTシートが見つかりません');
  return sheet;
}

function FcstReader_extractBlock_(row, startCol) {
  return {
    target: FcstReader_triplet_(row, startCol, BLOCK_OFFSETS.target),
    monthStartCommit: FcstReader_toNumber_(row[startCol + BLOCK_OFFSETS.monthStartCommit - 1]),
    fcstAdjusted: FcstReader_triplet_(row, startCol, BLOCK_OFFSETS.fcstAdjusted),
    fcstCommit: FcstReader_triplet_(row, startCol, BLOCK_OFFSETS.fcstCommit),
    received: FcstReader_triplet_(row, startCol, BLOCK_OFFSETS.received),
    debtMgmt: FcstReader_triplet_(row, startCol, BLOCK_OFFSETS.debtMgmt),
    debtMgmtLite: FcstReader_triplet_(row, startCol, BLOCK_OFFSETS.debtMgmtLite),
    expense: FcstReader_triplet_(row, startCol, BLOCK_OFFSETS.expense),
    expectedMrr: FcstReader_toNumber_(row[startCol + BLOCK_OFFSETS.expectedMrr - 1]),
    keyDeals: row[startCol + BLOCK_OFFSETS.keyDeals - 1] || '',
    targetDiff: FcstReader_triplet_(row, startCol, BLOCK_OFFSETS.targetDiff),
    monthStartDiff: FcstReader_toNumber_(row[startCol + BLOCK_OFFSETS.monthStartDiff - 1]),
    weekOverWeek: FcstReader_triplet_(row, startCol, BLOCK_OFFSETS.weekOverWeek),
    notes: row[startCol + BLOCK_OFFSETS.notes - 1] || '',
  };
}

function FcstReader_triplet_(row, startCol, offsets) {
  return {
    net: FcstReader_toNumber_(row[startCol + offsets.net - 1]),
    newExp: FcstReader_toNumber_(row[startCol + offsets.newExp - 1]),
    churn: FcstReader_toNumber_(row[startCol + offsets.churn - 1]),
  };
}

function FcstReader_toNumber_(value) {
  return typeof value === 'number' ? value : value === '' || value === null ? 0 : Number(value) || 0;
}

function FcstReader_isSnapshotValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return true;
  if (typeof value !== 'string') return false;
  return /\d{4}[\/-]\d{1,2}[\/-]\d{1,2}/.test(value.trim());
}

function FcstReader_formatDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  const text = String(value).trim();
  const date = new Date(text);
  if (!isNaN(date)) return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd');
  return text;
}
