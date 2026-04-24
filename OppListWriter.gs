function OppListWriter_saveDrafts(deptKey, changes) {
  var list = OppListWriter_normalizeChanges_(changes);
  if (!list.length) {
    return {
      success: true,
      exportWaiting: { newCount: 0, updatedCount: 0 },
      proposalExportWaiting: { newCount: 0, updatedCount: 0 }
    };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var exportRes = OppListWriter_upsertExportWaiting_(list);
    var proposalRes = isProposalProductsEnabled_(deptKey)
      ? OppListWriter_upsertProposalExportWaiting_(list)
      : { newCount: 0, updatedCount: 0 };

    return {
      success: true,
      exportWaiting: exportRes,
      proposalExportWaiting: proposalRes
    };
  } finally {
    lock.releaseLock();
  }
}

function OppListWriter_saveOppSfValue(deptKey, p) {
  var oppId = String(p && p.oppId || '').trim();
  var fieldName = String(p && p.fieldName || '').trim();
  if (!oppId || !fieldName) return { error: 'oppId and fieldName are required' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSharedSheet(EXPORT_WAITING_SHEET_NAME);
    if (!sheet) return { error: EXPORT_WAITING_SHEET_NAME + ' sheet not found' };

    var now = new Date();
    var lastColumn = Math.max(sheet.getLastColumn(), 6);
    var row = Array(lastColumn).fill('');
    row[0] = now;
    row[1] = oppId;
    row[2] = deptKey || '';
    row[3] = fieldName;
    row[4] = p && p.value !== undefined ? p.value : '';
    row[5] = 'sf_field_update';
    sheet.appendRow(row);
    return { status: 'ok' };
  } finally {
    lock.releaseLock();
  }
}

function OppListWriter_upsertExportWaiting_(changes) {
  var sheet = getSharedSheet(EXPORT_WAITING_SHEET_NAME);
  if (!sheet) return { newCount: 0, updatedCount: 0, error: EXPORT_WAITING_SHEET_NAME + ' sheet not found' };

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
    if (!OppListWriter_hasExportWaitingChanges_(change)) return;

    var oppId = String(change.oppId || '').trim();
    if (!oppId) return;

    var fields = change.fields || {};
    var rowIndex = rowMap.hasOwnProperty(oppId) ? rowMap[oppId] : -1;
    var row = rowIndex >= 0
      ? (values[rowIndex] || Array(colCount).fill('')).slice()
      : Array(colCount).fill('');
    var before = row.slice();

    row[0] = oppId;
    if (Object.prototype.hasOwnProperty.call(change, 'keyDeal')) row[1] = change.keyDeal === true;
    if (fields.hasOwnProperty('fcstCommit')) row[2] = OppListWriter_toNumberOrBlank_(fields.fcstCommit);
    if (fields.hasOwnProperty('fcstMin')) row[3] = OppListWriter_toNumberOrBlank_(fields.fcstMin);
    if (fields.hasOwnProperty('fcstMax')) row[4] = OppListWriter_toNumberOrBlank_(fields.fcstMax);
    if (fields.hasOwnProperty('comment')) row[5] = String(fields.comment || '');
    row[7] = '-';

    if (rowIndex >= 0) {
      if (OppListWriter_rowsEqual_(before, row)) return;
      values[rowIndex] = row;
      updatedIndexes[rowIndex] = true;
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

function OppListWriter_upsertProposalExportWaiting_(changes) {
  var sheet = getSharedSheet(EXPORT_WAITING_PROPOSAL_PRODUCTS_SHEET_NAME);
  if (!sheet) return { newCount: 0, updatedCount: 0, error: EXPORT_WAITING_PROPOSAL_PRODUCTS_SHEET_NAME + ' sheet not found' };

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
    { field: 'debtMgmtLite', proposalKey: 'debtMgmtLite', moduleName: '債権管理Lite' },
    { field: 'expense', proposalKey: 'expense', moduleName: 'EX' }
  ];

  var newRows = [];
  var updatedIndexes = {};
  var updatedCount = 0;

  changes.forEach(function(change) {
    if (!OppListWriter_hasProposalExportChanges_(change)) return;

    var fields = change.fields || {};
    var proposalIds = change.proposalProductIds || {};

    modules.forEach(function(module) {
      if (!fields.hasOwnProperty(module.field)) return;

      var proposalId = String(proposalIds[module.proposalKey] || '').trim();
      if (!proposalId) return;

      var rowIndex = rowMap.hasOwnProperty(proposalId) ? rowMap[proposalId] : -1;
      var row = rowIndex >= 0
        ? (values[rowIndex] || Array(colCount).fill('')).slice()
        : Array(colCount).fill('');
      var before = row.slice();

      row[0] = proposalId;
      row[1] = module.moduleName;
      row[2] = OppListWriter_toNumberOrBlank_(fields[module.field]);
      row[3] = String(change.oppId || '');
      row[5] = '-';

      if (rowIndex >= 0) {
        if (OppListWriter_rowsEqual_(before, row)) return;
        values[rowIndex] = row;
        updatedIndexes[rowIndex] = true;
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

function OppListWriter_normalizeChanges_(changes) {
  var mergedByOppId = {};
  var merged = [];

  (Array.isArray(changes) ? changes : []).forEach(function(change) {
    var oppId = String(change && change.oppId || '').trim();
    if (!oppId) return;

    if (!mergedByOppId[oppId]) {
      mergedByOppId[oppId] = {
        oppId: oppId,
        proposalProductIds: {},
        fields: {}
      };
      merged.push(mergedByOppId[oppId]);
    }

    var target = mergedByOppId[oppId];
    if (Object.prototype.hasOwnProperty.call(change, 'keyDeal')) {
      target.keyDeal = change.keyDeal === true;
    }

    var proposalIds = change && change.proposalProductIds || {};
    Object.keys(proposalIds).forEach(function(key) {
      if (!Object.prototype.hasOwnProperty.call(proposalIds, key)) return;
      var value = String(proposalIds[key] || '').trim();
      if (value) target.proposalProductIds[key] = value;
    });

    var fields = change && change.fields || {};
    Object.keys(fields).forEach(function(key) {
      if (!Object.prototype.hasOwnProperty.call(fields, key)) return;
      target.fields[key] = fields[key];
    });
  });

  return merged;
}

function OppListWriter_hasExportWaitingChanges_(change) {
  if (change && Object.prototype.hasOwnProperty.call(change, 'keyDeal')) return true;
  return OppListWriter_hasAnyOwnProperty_(change && change.fields, ['fcstCommit', 'fcstMin', 'fcstMax', 'comment']);
}

function OppListWriter_hasProposalExportChanges_(change) {
  return OppListWriter_hasAnyOwnProperty_(change && change.fields, ['received', 'debtMgmt', 'debtMgmtLite', 'expense']);
}

function OppListWriter_hasAnyOwnProperty_(obj, keys) {
  var source = obj || {};
  for (var i = 0; i < keys.length; i++) {
    if (Object.prototype.hasOwnProperty.call(source, keys[i])) return true;
  }
  return false;
}

function OppListWriter_rowsEqual_(left, right) {
  var a = Array.isArray(left) ? left : [];
  var b = Array.isArray(right) ? right : [];
  if (a.length !== b.length) return false;
  for (var i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}

function saveDiffOppList(deptKey, dirtyRows, userEmail, timestamp) {
  if (!dirtyRows || dirtyRows.length === 0) return { status: 'no_changes' };
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    appendDiffToExportWaiting_(ss, dirtyRows, userEmail, timestamp);
    appendToChangeLog_(ss, dirtyRows, userEmail, timestamp);
    return { status: 'ok', savedRows: dirtyRows.length };
  } finally {
    lock.releaseLock();
  }
}

function appendDiffToExportWaiting_(ss, dirtyRows, userEmail, timestamp) {
  var sheet = getSharedSheet(EXPORT_WAITING_SHEET_NAME);
  if (!sheet) return;
  var rows = [];
  dirtyRows.forEach(function(dr) {
    Object.keys(dr.changes).forEach(function(col) {
      var ch = dr.changes[col];
      rows.push([timestamp, userEmail, dr.rowKey, col, ch.old, ch.new]);
    });
  });
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  }
}

function appendToChangeLog_(ss, dirtyRows, userEmail, timestamp) {
  var sheet = getSharedSheet(CHANGE_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CHANGE_LOG_SHEET_NAME);
    sheet.appendRow(['timestamp', 'user_email', 'row_key', 'column', 'old_value', 'new_value']);
  }
  var rows = [];
  dirtyRows.forEach(function(dr) {
    Object.keys(dr.changes).forEach(function(col) {
      var ch = dr.changes[col];
      rows.push([timestamp, userEmail, dr.rowKey, col, ch.old, ch.new]);
    });
  });
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  }
}
