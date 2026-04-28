const SNAPSHOT_STORAGE_CELL_LIMIT = 9900000;
const SNAPSHOT_STORAGE_PROP_PREFIX = 'snapshotStorage:';
const SNAPSHOT_STORAGE_CLEANUP_ARM_PROP = SNAPSHOT_STORAGE_PROP_PREFIX + 'freshStartCleanupArmedUntil';
const SNAPSHOT_STORAGE_INDEX_HEADERS = ['created_at', 'sheet_name', 'file_id', 'file_url', 'active', 'cell_count', 'row_count'];

function SnapshotStorage_getReadSheets_(sheetName, headers) {
  var sheets = [];

  SnapshotStorage_getFileIds_(sheetName).forEach(function(fileId) {
    try {
      var ss = SpreadsheetApp.openById(fileId);
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        SnapshotStorage_ensureSheetShape_(sheet, headers || [], false);
        sheets.push(sheet);
      }
    } catch (e) {
      Logger.log('SnapshotStorage read file skipped: ' + fileId + ' / ' + (e && e.message ? e.message : e));
    }
  });

  return sheets;
}

function SnapshotStorage_getAllValues_(sheetName, headers, columnCount) {
  var values = [];
  SnapshotStorage_getReadSheets_(sheetName, headers).forEach(function(sheet) {
    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return;
    values = values.concat(sheet.getRange(1, 1, lastRow, columnCount).getValues());
  });
  return values;
}

function SnapshotStorage_appendRows_(sheetName, headers, rows) {
  var appendRows = rows || [];
  if (!appendRows.length) {
    return {
      sheet: null,
      sheetName: sheetName,
      fileId: '',
      fileUrl: '',
      rolledOver: false
    };
  }

  var sheet = SnapshotStorage_getWriteSheet_(sheetName, headers, appendRows.length);
  try {
    SnapshotStorage_ensureAppendRows_(sheet, appendRows.length);
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, headers.length).setValues(appendRows);
    return SnapshotStorage_buildWriteResult_(sheet, false);
  } catch (e) {
    if (!SnapshotStorage_isCellLimitError_(e)) throw e;
    sheet = SnapshotStorage_createFile_(sheetName, headers);
    SnapshotStorage_ensureAppendRows_(sheet, appendRows.length);
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, headers.length).setValues(appendRows);
    return SnapshotStorage_buildWriteResult_(sheet, true);
  }
}

function SnapshotStorage_getWriteSheet_(sheetName, headers, appendRowCount) {
  var sheet = SnapshotStorage_getActiveSheet_(sheetName, headers);
  if (!sheet || SnapshotStorage_wouldExceedCellLimit_(sheet, headers.length, appendRowCount || 0)) {
    sheet = SnapshotStorage_createFile_(sheetName, headers);
  }
  SnapshotStorage_ensureSheetShape_(sheet, headers, true);
  return sheet;
}

function SnapshotStorage_getActiveSheet_(sheetName, headers) {
  var activeFileId = SnapshotStorage_getActiveFileId_(sheetName);
  if (activeFileId) {
    try {
      var ss = SpreadsheetApp.openById(activeFileId);
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        SnapshotStorage_ensureSheetShape_(sheet, headers, true);
        return sheet;
      }
    } catch (e) {
      Logger.log('SnapshotStorage active file unavailable: ' + activeFileId + ' / ' + (e && e.message ? e.message : e));
    }
  }

  return SnapshotStorage_createFile_(sheetName, headers);
}

function SnapshotStorage_createFile_(sheetName, headers) {
  var timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmmss');
  var ss = SpreadsheetApp.create('FcstDashboard DB - ' + sheetName + ' - ' + timestamp);
  var sheet = ss.getSheets()[0];
  sheet.setName(sheetName);
  SnapshotStorage_ensureSheetShape_(sheet, headers, true);
  SnapshotStorage_moveFileToDbFolder_(ss);

  var fileId = ss.getId();
  var ids = SnapshotStorage_getFileIds_(sheetName);
  if (ids.indexOf(fileId) === -1) ids.push(fileId);
  SnapshotStorage_setFileIds_(sheetName, ids);
  SnapshotStorage_setActiveFileId_(sheetName, fileId);
  SnapshotStorage_recordIndex_(sheetName, ss, sheet, true);
  Logger.log('SnapshotStorage created file: sheet=' + sheetName + ' fileId=' + fileId + ' url=' + ss.getUrl());
  return sheet;
}

function SnapshotStorage_moveFileToDbFolder_(spreadsheet) {
  var folderId = String(typeof SNAPSHOT_DB_FOLDER_ID !== 'undefined' ? SNAPSHOT_DB_FOLDER_ID : '').trim();
  if (!folderId) return;

  var fileId = spreadsheet.getId();
  try {
    DriveApp.getFileById(fileId).moveTo(DriveApp.getFolderById(folderId));
  } catch (e) {
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (ignored) {}
    throw new Error('Snapshot DB file could not be moved to configured folder: ' + folderId + ' / ' + (e && e.message ? e.message : e));
  }
}

function SnapshotStorage_ensureSheetShape_(sheet, headers, allowWrite) {
  var columnCount = headers.length;
  if (!columnCount) return;

  var maxColumns = sheet.getMaxColumns();
  if (maxColumns < columnCount && allowWrite) {
    sheet.insertColumnsAfter(maxColumns, columnCount - maxColumns);
  }
  if (sheet.getMaxColumns() > columnCount && allowWrite) {
    sheet.deleteColumns(columnCount + 1, sheet.getMaxColumns() - columnCount);
  }

  if (allowWrite && sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, columnCount).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function SnapshotStorage_ensureAppendRows_(sheet, appendRowCount) {
  var requiredLastRow = sheet.getLastRow() + appendRowCount;
  if (requiredLastRow <= sheet.getMaxRows()) return;
  sheet.insertRowsAfter(sheet.getMaxRows(), requiredLastRow - sheet.getMaxRows());
}

function SnapshotStorage_wouldExceedCellLimit_(sheet, columnCount, appendRowCount) {
  var ss = sheet.getParent();
  var currentCells = SnapshotStorage_countCells_(ss);
  var requiredLastRow = sheet.getLastRow() + appendRowCount;
  var additionalRows = Math.max(0, requiredLastRow - sheet.getMaxRows());
  var additionalCells = additionalRows * Math.max(columnCount, sheet.getMaxColumns());
  return currentCells + additionalCells > SNAPSHOT_STORAGE_CELL_LIMIT;
}

function SnapshotStorage_countCells_(spreadsheet) {
  return spreadsheet.getSheets().reduce(function(total, sheet) {
    return total + sheet.getMaxRows() * sheet.getMaxColumns();
  }, 0);
}

function SnapshotStorage_buildWriteResult_(sheet, rolledOver) {
  var ss = sheet.getParent();
  SnapshotStorage_recordIndex_(sheet.getName(), ss, sheet, true);
  return {
    sheet: sheet,
    sheetName: sheet.getName(),
    fileId: ss.getId(),
    fileUrl: ss.getUrl(),
    rolledOver: !!rolledOver
  };
}

function SnapshotStorage_recordSheet_(sheetName, sheet) {
  if (!sheet) return;
  SnapshotStorage_recordIndex_(sheetName, sheet.getParent(), sheet, true);
}

function SnapshotStorage_getFreshStartPlan() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var result = {
    mainSpreadsheetId: SPREADSHEET_ID,
    mainSpreadsheetUrl: ss.getUrl(),
    mainSheetsToDelete: [],
    dbFilesToTrash: [],
    unregisteredDbFileCandidates: [],
    unregisteredDbFileSearchSkipped: true,
    scriptPropertiesToClear: []
  };
  var seenFileIds = {};

  SnapshotStorage_getFreshStartMainSheetNames_().forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    result.mainSheetsToDelete.push({
      sheetName: sheetName,
      rows: sheet.getLastRow(),
      columns: sheet.getLastColumn()
    });
  });

  var props = PropertiesService.getScriptProperties().getProperties();
  Object.keys(props).sort().forEach(function(key) {
    if (key.indexOf(SNAPSHOT_STORAGE_PROP_PREFIX) !== 0) return;
    result.scriptPropertiesToClear.push(key);
  });

  SnapshotStorage_getDbOwnedSheetNames_().forEach(function(sheetName) {
    var activeFileId = SnapshotStorage_getActiveFileId_(sheetName);
    SnapshotStorage_getFileIds_(sheetName).forEach(function(fileId) {
      SnapshotStorage_addFreshStartDbFile_(result, sheetName, fileId, fileId === activeFileId, true, seenFileIds);
    });
  });

  return result;
}

function manualLogSnapshotStorageFreshStartPlan() {
  var plan = SnapshotStorage_getFreshStartPlan();
  Logger.log(JSON.stringify(plan, null, 2));
  return plan;
}

function manualArmSnapshotStorageFreshStartCleanup() {
  var expiresAt = Date.now() + 10 * 60 * 1000;
  PropertiesService.getScriptProperties().setProperty(SNAPSHOT_STORAGE_CLEANUP_ARM_PROP, String(expiresAt));
  var plan = SnapshotStorage_getFreshStartPlan();
  plan.armedUntil = Utilities.formatDate(new Date(expiresAt), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  Logger.log(JSON.stringify(plan, null, 2));
  return plan;
}

function manualCleanupSnapshotStorageForFreshStart() {
  var result = SnapshotStorage_cleanupForFreshStart();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function SnapshotStorage_cleanupForFreshStart() {
  SnapshotStorage_assertFreshStartCleanupArmed_();
  var plan = SnapshotStorage_getFreshStartPlan();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var deletedMainSheets = [];
  var trashedDbFiles = [];
  var clearErrors = [];

  SnapshotStorage_getFreshStartMainSheetNames_().forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    if (ss.getSheets().length <= 1) {
      clearErrors.push('Skipped deleting the last sheet: ' + sheetName);
      return;
    }
    ss.deleteSheet(sheet);
    deletedMainSheets.push(sheetName);
  });

  var seenFiles = {};
  plan.dbFilesToTrash.forEach(function(entry) {
    var fileId = String(entry.fileId || '').trim();
    if (!fileId || seenFiles[fileId]) return;
    seenFiles[fileId] = true;
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
      trashedDbFiles.push(fileId);
    } catch (e) {
      clearErrors.push('Failed to trash DB file ' + fileId + ': ' + (e && e.message ? e.message : e));
    }
  });

  var props = PropertiesService.getScriptProperties();
  plan.scriptPropertiesToClear.forEach(function(key) {
    props.deleteProperty(key);
  });
  props.deleteProperty(SNAPSHOT_STORAGE_CLEANUP_ARM_PROP);

  return {
    ok: clearErrors.length === 0,
    deletedMainSheets: deletedMainSheets,
    trashedDbFiles: trashedDbFiles,
    clearedProperties: plan.scriptPropertiesToClear,
    errors: clearErrors
  };
}

function SnapshotStorage_getDbOwnedSheetNames_() {
  return [
    FCST_SNAPSHOT_SHEET_NAME,
    OPP_LIST_SNAPSHOT_SHEET_NAME,
    FCST_ADJUSTED_SHEET_NAME
  ];
}

function SnapshotStorage_getFreshStartMainSheetNames_() {
  return SnapshotStorage_getDbOwnedSheetNames_().concat([
    SNAPSHOT_DB_INDEX_SHEET_NAME
  ]);
}

function SnapshotStorage_assertFreshStartCleanupArmed_() {
  var raw = PropertiesService.getScriptProperties().getProperty(SNAPSHOT_STORAGE_CLEANUP_ARM_PROP);
  var expiresAt = Number(raw) || 0;
  if (!expiresAt || Date.now() > expiresAt) {
    throw new Error('Run manualArmSnapshotStorageFreshStartCleanup first. Cleanup stays armed for 10 minutes.');
  }
}

function SnapshotStorage_addFreshStartDbFile_(result, sheetName, fileId, active, registered, seenFileIds) {
  var id = String(fileId || '').trim();
  if (!id || seenFileIds[id]) return;
  seenFileIds[id] = true;

  var entry = {
    sheetName: sheetName || '',
    fileId: id,
    active: !!active,
    registered: !!registered,
    fileName: '',
    fileUrl: '',
    rows: 0,
    error: ''
  };

  try {
    var db = SpreadsheetApp.openById(id);
    var sheet = SnapshotStorage_findDbOwnedSheet_(db, sheetName);
    entry.fileName = db.getName();
    entry.fileUrl = db.getUrl();
    entry.sheetName = sheet ? sheet.getName() : entry.sheetName;
    entry.rows = sheet ? sheet.getLastRow() : 0;
  } catch (e) {
    entry.error = e && e.message ? e.message : String(e);
  }

  result.dbFilesToTrash.push(entry);
}

function SnapshotStorage_findDbOwnedSheet_(spreadsheet, preferredSheetName) {
  var preferred = String(preferredSheetName || '').trim();
  if (preferred) {
    var sheet = spreadsheet.getSheetByName(preferred);
    if (sheet) return sheet;
  }

  var sheetNames = SnapshotStorage_getDbOwnedSheetNames_();
  for (var i = 0; i < sheetNames.length; i++) {
    var candidate = spreadsheet.getSheetByName(sheetNames[i]);
    if (candidate) return candidate;
  }
  return null;
}

function SnapshotStorage_recordIndex_(sheetName, spreadsheet, sheet, active) {
  try {
    var indexSheet = SnapshotStorage_getIndexSheet_();
    var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    var fileId = spreadsheet.getId();
    var values = indexSheet.getLastRow() > 1
      ? indexSheet.getRange(2, 1, indexSheet.getLastRow() - 1, SNAPSHOT_STORAGE_INDEX_HEADERS.length).getValues()
      : [];
    var rowIndex = -1;
    values.forEach(function(row, idx) {
      if (String(row[1] || '') === sheetName && String(row[2] || '') !== fileId && active) {
        indexSheet.getRange(idx + 2, 5).setValue('FALSE');
      }
      if (String(row[2] || '') === fileId && String(row[1] || '') === sheetName) rowIndex = idx + 2;
    });

    var rowValues = [
      now,
      sheetName,
      fileId,
      spreadsheet.getUrl(),
      active ? 'TRUE' : 'FALSE',
      SnapshotStorage_countCells_(spreadsheet),
      sheet.getLastRow()
    ];
    if (rowIndex > 0) {
      indexSheet.getRange(rowIndex, 1, 1, SNAPSHOT_STORAGE_INDEX_HEADERS.length).setValues([rowValues]);
    } else {
      indexSheet.getRange(indexSheet.getLastRow() + 1, 1, 1, SNAPSHOT_STORAGE_INDEX_HEADERS.length).setValues([rowValues]);
    }
  } catch (e) {
    Logger.log('SnapshotStorage index failed: ' + (e && e.message ? e.message : e));
  }
}

function SnapshotStorage_getIndexSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SNAPSHOT_DB_INDEX_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SNAPSHOT_DB_INDEX_SHEET_NAME);
  if (sheet.getMaxColumns() < SNAPSHOT_STORAGE_INDEX_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), SNAPSHOT_STORAGE_INDEX_HEADERS.length - sheet.getMaxColumns());
  }
  if (sheet.getMaxColumns() > SNAPSHOT_STORAGE_INDEX_HEADERS.length) {
    sheet.deleteColumns(SNAPSHOT_STORAGE_INDEX_HEADERS.length + 1, sheet.getMaxColumns() - SNAPSHOT_STORAGE_INDEX_HEADERS.length);
  }
  var minRows = Math.max(sheet.getLastRow(), 100);
  if (sheet.getMaxRows() > minRows) {
    sheet.deleteRows(minRows + 1, sheet.getMaxRows() - minRows);
  }
  var shouldWriteHeader = sheet.getLastRow() < 1;
  if (!shouldWriteHeader) {
    var existing = sheet.getRange(1, 1, 1, SNAPSHOT_STORAGE_INDEX_HEADERS.length).getValues()[0];
    shouldWriteHeader = SNAPSHOT_STORAGE_INDEX_HEADERS.some(function(header, idx) {
      return String(existing[idx] || '') !== header;
    });
  }
  if (shouldWriteHeader) {
    sheet.getRange(1, 1, 1, SNAPSHOT_STORAGE_INDEX_HEADERS.length).setValues([SNAPSHOT_STORAGE_INDEX_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function SnapshotStorage_isCellLimitError_(error) {
  var message = String(error && error.message ? error.message : error);
  return message.indexOf('10000000') !== -1 || /cell/i.test(message);
}

function SnapshotStorage_getFileIds_(sheetName) {
  var raw = PropertiesService.getScriptProperties().getProperty(SnapshotStorage_propKey_(sheetName, 'fileIds'));
  if (!raw) return [];
  try {
    var parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.filter(function(id) { return !!String(id || '').trim(); }) : [];
  } catch (e) {
    return [];
  }
}

function SnapshotStorage_setFileIds_(sheetName, fileIds) {
  PropertiesService.getScriptProperties().setProperty(
    SnapshotStorage_propKey_(sheetName, 'fileIds'),
    JSON.stringify(fileIds || [])
  );
}

function SnapshotStorage_getActiveFileId_(sheetName) {
  return String(PropertiesService.getScriptProperties().getProperty(SnapshotStorage_propKey_(sheetName, 'activeFileId')) || '').trim();
}

function SnapshotStorage_setActiveFileId_(sheetName, fileId) {
  PropertiesService.getScriptProperties().setProperty(SnapshotStorage_propKey_(sheetName, 'activeFileId'), String(fileId || '').trim());
}

function SnapshotStorage_propKey_(sheetName, suffix) {
  return SNAPSHOT_STORAGE_PROP_PREFIX + Utilities.base64EncodeWebSafe(String(sheetName || '')) + ':' + suffix;
}
