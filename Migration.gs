function migrateToSharedArchitecture() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  LEGACY_DISPLAY_SHEET_NAMES.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      ss.deleteSheet(sheet);
      Logger.log('Deleted display sheet: ' + name);
    }
  });

  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    for (var i = 0; i < LEGACY_DEPT_SHEET_PREFIXES.length; i++) {
      if (name.indexOf(LEGACY_DEPT_SHEET_PREFIXES[i]) !== 0) continue;
      ss.deleteSheet(sheet);
      Logger.log('Deleted old-prefix sheet: ' + name);
      break;
    }
  });

  if (!ss.getSheetByName(CHANGE_LOG_SHEET_NAME)) {
    var changeLog = ss.insertSheet(CHANGE_LOG_SHEET_NAME);
    changeLog.getRange(1, 1, 1, 6).setValues([
      ['timestamp', 'user_email', 'row_key', 'column', 'old_value', 'new_value']
    ]);
    Logger.log('Created: ' + CHANGE_LOG_SHEET_NAME);
  }

  MasterSchema_setupSheets();
  Logger.log('Created/updated master sheets: ' + [ORG_MASTER_SHEET_NAME, TARGET_MASTER_SHEET_NAME, ASSIGNMENT_MASTER_SHEET_NAME].join(', '));

  Logger.log('Migration complete.');
}

function createSfDataSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var existing = ss.getSheetByName(LEGACY_SF_DATA_SHEET_NAME);
  if (existing) {
    existing.setName(SF_DATA_SHEET_BO);
    Logger.log('Renamed: ' + LEGACY_SF_DATA_SHEET_NAME + ' → ' + SF_DATA_SHEET_BO);
  }

  [SF_DATA_SHEET_SS, SF_DATA_SHEET_SSCS, SF_DATA_SHEET_CO].forEach(function(name) {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
      Logger.log('Created: ' + name);
    }
  });

  Logger.log(LEGACY_SF_DATA_SHEET_NAME + ' 4-sheet split complete.');
  Logger.log('NEXT: Configure Coefficient sync for 4 sheets (当期19期+来期20期)');
}
