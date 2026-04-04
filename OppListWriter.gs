function OppListWriter_saveDrafts(changes) {
  var list = Array.isArray(changes) ? changes : [];
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var exportRes = OppListWriter_upsertExportWaiting_(ss, list);
  var proposalRes = OppListWriter_upsertProposalExportWaiting_(ss, list);

  return {
    success: true,
    exportWaiting: exportRes,
    proposalExportWaiting: proposalRes
  };
}

function OppListWriter_upsertExportWaiting_(ss, changes) {
  var sheet = ss.getSheetByName('Export待機');
  if (!sheet) throw new Error('Export待機シートが見つかりません');

  var startRow = 2;
  var colCount = 8;
  var values = sheet.getLastRow() >= startRow
    ? sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, colCount).getValues()
    : [];
  var rowMap = {};

  values.forEach(function(row, idx) {
    var oppId = String(row[0] || '').trim();
    if (oppId) rowMap[oppId] = idx;
  });

  var newRows = [];
  var updatedIndexes = {};
  var updatedCount = 0;

  changes.forEach(function(change) {
    var oppId = String(change.oppId || '').trim();
    if (!oppId) return;

    var fields = change.fields || {};
    var row = rowMap.hasOwnProperty(oppId)
      ? (values[rowMap[oppId]] || Array(colCount).fill('')).slice()
      : Array(colCount).fill('');

    row[0] = oppId;
    row[1] = change.keyDeal === true;
    if (fields.hasOwnProperty('fcstCommit')) row[2] = OppListWriter_toNumberOrBlank_(fields.fcstCommit);
    if (fields.hasOwnProperty('fcstMin')) row[3] = OppListWriter_toNumberOrBlank_(fields.fcstMin);
    if (fields.hasOwnProperty('fcstMax')) row[4] = OppListWriter_toNumberOrBlank_(fields.fcstMax);
    if (fields.hasOwnProperty('comment')) row[5] = String(fields.comment || '');
    row[7] = '-';

    if (rowMap.hasOwnProperty(oppId)) {
      values[rowMap[oppId]] = row;
      updatedIndexes[rowMap[oppId]] = true;
      updatedCount++;
    } else {
      newRows.push(row);
    }
  });

  Object.keys(updatedIndexes).forEach(function(idxText) {
    var idx = Number(idxText);
    sheet.getRange(startRow + idx, 1, 1, colCount).setValues([values[idx]]);
  });

  if (newRows.length) {
    var appendRow = Math.max(sheet.getLastRow() + 1, startRow);
    sheet.getRange(appendRow, 1, newRows.length, colCount).setValues(newRows);
  }

  return { newCount: newRows.length, updatedCount: updatedCount };
}

function OppListWriter_upsertProposalExportWaiting_(ss, changes) {
  var sheet = ss.getSheetByName('Export待機_提案商品');
  if (!sheet) throw new Error('Export待機_提案商品シートが見つかりません');

  var startRow = 2;
  var colCount = 6;
  var values = sheet.getLastRow() >= startRow
    ? sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, colCount).getValues()
    : [];
  var rowMap = {};

  values.forEach(function(row, idx) {
    var proposalId = String(row[0] || '').trim();
    if (proposalId) rowMap[proposalId] = idx;
  });

  var modules = [
    { field: 'received', proposalKey: 'received', moduleName: 'AP' },
    { field: 'debtMgmt', proposalKey: 'debtMgmt', moduleName: 'AR' },
    { field: 'debtMgmtLite', proposalKey: 'debtMgmtLite', moduleName: '債権管理 Lite' },
    { field: 'expense', proposalKey: 'expense', moduleName: 'EX' }
  ];

  var newRows = [];
  var updatedIndexes = {};
  var updatedCount = 0;

  changes.forEach(function(change) {
    var fields = change.fields || {};
    var proposalIds = change.proposalProductIds || {};

    modules.forEach(function(module) {
      if (!fields.hasOwnProperty(module.field)) return;

      var proposalId = String(proposalIds[module.proposalKey] || '').trim();
      if (!proposalId) return;

      var row = rowMap.hasOwnProperty(proposalId)
        ? (values[rowMap[proposalId]] || Array(colCount).fill('')).slice()
        : Array(colCount).fill('');

      row[0] = proposalId;
      row[1] = module.moduleName;
      row[2] = OppListWriter_toNumberOrBlank_(fields[module.field]);
      row[3] = String(change.oppId || '');
      row[5] = '-';

      if (rowMap.hasOwnProperty(proposalId)) {
        values[rowMap[proposalId]] = row;
        updatedIndexes[rowMap[proposalId]] = true;
        updatedCount++;
      } else {
        newRows.push(row);
      }
    });
  });

  Object.keys(updatedIndexes).forEach(function(idxText) {
    var idx = Number(idxText);
    sheet.getRange(startRow + idx, 1, 1, colCount).setValues([values[idx]]);
  });

  if (newRows.length) {
    var appendRow = Math.max(sheet.getLastRow() + 1, startRow);
    sheet.getRange(appendRow, 1, newRows.length, colCount).setValues(newRows);
  }

  return { newCount: newRows.length, updatedCount: updatedCount };
}

function OppListWriter_toNumberOrBlank_(value) {
  if (value === '' || value === null || value === undefined) return '';
  return Number(value) || 0;
}
