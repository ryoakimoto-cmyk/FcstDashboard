const SPREADSHEET_ID = '1j_0TNBRhhfLLcr9DBr5mPLitiFRXYtsqohfQ-fxSPGg';
const CURRENT_FISCAL_YEAR = 19;

const FCST_ADJUSTED_SHEET_NAME = 'FCST調整';
const EXPORT_WAITING_SHEET_NAME = 'Export待機';
const AGGREGATED_CACHE_SHEET_NAME = '集計キャッシュ';
const DEPT_MASTER_SHEET_NAME = '部門マスタ';
const MONTHLY_TARGET_MASTER_SHEET_NAME = '月次目標マスタ';
const CHANGE_LOG_SHEET_NAME = '変更ログ';
const TARGET_SHEET_NAME = '目標';
const FCST_SNAPSHOT_SHEET_NAME = 'FCSTスナップショット';
const OPP_WEBAPP_INPUT_SHEET_NAME = '案件_WebApp入力';
const OPP_HISTORY_V2_SHEET_NAME = 'Opp履歴';
const SNAPSHOT_META_SHEET_NAME = 'スナップショットメタ';
const FCST_DECLARATION_CURRENT_SHEET_NAME = 'FCST宣言_現在';
const EXPORT_WAITING_PROPOSAL_PRODUCTS_SHEET_NAME = 'Export待機_提案商品';
const SF_USER_SHEET_NAME = 'SFユーザー';
const SF_DATA_SHEET_BO = 'SFデータ更新_BO';
const SF_DATA_SHEET_SS = 'SFデータ更新_SS';
const SF_DATA_SHEET_SSCS = 'SFデータ更新_SSCS';
const SF_DATA_SHEET_CO = 'SFデータ更新_CO';
const LEGACY_SF_DATA_SHEET_NAME = 'SFデータ更新';
const LEGACY_DISPLAY_SHEET_NAMES = ['FCST', '案件リスト', '案件リスト(最新)', '集計'];
const LEGACY_DEPT_SHEET_PREFIXES = [
  'BOAM_', 'BO_EAST1_', 'BO_EAST2_', 'BO_WEST_', 'BO_PARTNER_',
  'EP1_', 'EP2_', 'EP3_', 'SMB1_', 'SMB2_', 'SMB3_', 'SMB4_', 'SMB5_', 'SMBCS_'
];

const SSCS_COLUMN_MAP = {
  '案件CS担当部署': '担当部署',
  'CS_Key Deal フラグ': 'Key Deal フラグ',
  'CS_FCST(コミット)(換算値)': 'FCST(コミット)(換算値)',
  'CS_FCST(MIN)(換算値)': 'FCST(MIN)(換算値)',
  'CS_FCST(MAX)(換算値)': 'FCST(MAX)(換算値)',
  'CS_FCSTコメント': 'FCSTコメント'
};

const SSCS_CONFLICT_COLUMNS = [
  'FCST(コミット)(換算値)',
  'FCST(MIN)(換算値)',
  'FCST(MAX)(換算値)',
  'FCSTコメント'
];

const DEPT_CONFIG = {
  BOCS:    { label: 'BOCS',   division: 'BO', sfSheet: SF_DATA_SHEET_BO, features: { oppList: true, adjustment: true, snapshot: true, chart: true, proposalProducts: true } },
  BOE1:    { label: 'BO東1',  division: 'BO', sfSheet: SF_DATA_SHEET_BO, features: { oppList: true, adjustment: true, snapshot: true, chart: true, proposalProducts: true } },
  BOE2:    { label: 'BO東2',  division: 'BO', sfSheet: SF_DATA_SHEET_BO, features: { oppList: true, adjustment: true, snapshot: true, chart: true, proposalProducts: true } },
  BOW:     { label: 'BO西',   division: 'BO', sfSheet: SF_DATA_SHEET_BO, features: { oppList: true, adjustment: true, snapshot: true, chart: true, proposalProducts: true } },
  BOPA:    { label: 'BOPA',   division: 'BO', sfSheet: SF_DATA_SHEET_BO, features: { oppList: true, adjustment: true, snapshot: true, chart: true, proposalProducts: true } },
  SSEP1:   { label: 'EP1',    division: 'SS', sfSheet: SF_DATA_SHEET_SS, features: { oppList: true, adjustment: true, snapshot: true, chart: true, proposalProducts: false } },
  SSEP2:   { label: 'EP2',    division: 'SS', sfSheet: SF_DATA_SHEET_SS, features: { oppList: true, adjustment: true, snapshot: true, chart: true, proposalProducts: false } },
  SSEP3:   { label: 'EP3',    division: 'SS', sfSheet: SF_DATA_SHEET_SS, features: { oppList: true, adjustment: true, snapshot: true, chart: true, proposalProducts: false } },
  SSSMB1:  { label: 'SMB1',   division: 'SS', sfSheet: SF_DATA_SHEET_SS, features: { oppList: true, adjustment: true, snapshot: true, chart: true, proposalProducts: false } },
  SSSMB2:  { label: 'SMB2',   division: 'SS', sfSheet: SF_DATA_SHEET_SS, features: { oppList: true, adjustment: true, snapshot: true, chart: true, proposalProducts: false } },
  SSSMB3:  { label: 'SMB3',   division: 'SS', sfSheet: SF_DATA_SHEET_SS, features: { oppList: true, adjustment: true, snapshot: true, chart: true, proposalProducts: false } },
  SSSMB4:  { label: 'SMB4',   division: 'SS', sfSheet: SF_DATA_SHEET_SS, features: { oppList: true, adjustment: true, snapshot: true, chart: true, proposalProducts: false } },
  SSSMB5:  { label: 'SMB5',   division: 'SS', sfSheet: SF_DATA_SHEET_SS, features: { oppList: true, adjustment: true, snapshot: true, chart: true, proposalProducts: false } },
  SSSMBCS: { label: 'SMBCS',  division: 'SS', sfSheet: SF_DATA_SHEET_SSCS, features: { oppList: true, adjustment: true, snapshot: true, chart: true, proposalProducts: false } },
  CO:      { label: 'CO',     division: 'CO', sfSheet: SF_DATA_SHEET_CO, features: { oppList: true, adjustment: true, snapshot: true, chart: true, proposalProducts: false } }
};

function getDeptConfig_(deptKey) {
  return deptKey && DEPT_CONFIG[deptKey] ? DEPT_CONFIG[deptKey] : null;
}

function isFeatureEnabled_(deptKey, featureKey) {
  var cfg = getDeptConfig_(deptKey);
  if (!cfg || !cfg.features) return false;
  return cfg.features[featureKey] !== false;
}

function isProposalProductsEnabled_(deptKey) {
  return isFeatureEnabled_(deptKey, 'proposalProducts');
}

function getSharedSheet(sheetName) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
}

function getSfDataSheet_(deptKey) {
  var cfg = DEPT_CONFIG[deptKey];
  if (!cfg || !cfg.sfSheet) return null;
  return getSharedSheet(cfg.sfSheet);
}

function normalizeSSCSHeaders_(headers) {
  return (headers || []).map(function(header) {
    var name = String(header || '').trim();
    if (SSCS_COLUMN_MAP[name]) return SSCS_COLUMN_MAP[name];
    if (SSCS_CONFLICT_COLUMNS.indexOf(name) !== -1) return '_orig_' + name;
    return name;
  });
}

const COEFFICIENT_REFRESH_WEBHOOK_URL_MIGRATION = 'https://app.coefficient.io/api/webhook/sheet_1j_0TNBRhhfLLcr9DBr5mPLitiFRXYtsqohfQ-fxSPGg_131df86d-242f-4820-be9e-083fa1610ed8';
const COEFFICIENT_REFRESH_MIN_INTERVAL_MS = 5 * 60 * 1000;
const COEFFICIENT_REFRESH_POLL_INTERVAL_MS = 5000;
const COEFFICIENT_REFRESH_TIMEOUT_MS = 60000;

const BLOCK_OFFSETS = {
  target: { net: 1, newExp: 2, churn: 3 },
  monthStartCommit: 4,
  fcstAdjusted: { net: 5, newExp: 6, churn: 7 },
  fcstCommit: { net: 8, newExp: 9, churn: 10 },
  received: { net: 11, newExp: 12, churn: 13 },
  debtMgmt: { net: 14, newExp: 15, churn: 16 },
  debtMgmtLite: { net: 17, newExp: 18, churn: 19 },
  expense: { net: 20, newExp: 21, churn: 22 },
  expectedMrr: 23,
  keyDeals: 24,
  targetDiff: { net: 25, newExp: 26, churn: 27 },
  monthStartDiff: 28,
  weekOverWeek: { net: 29, newExp: 30, churn: 31 },
  notes: 32
};

const BLOCK_COL_STARTS = { Q: 4, M5: 38, M6: 72, M7: 106 };
const ROWS_PER_SNAPSHOT = 32;
const SNAPSHOT_START_ROW = 4;
const OPP_HEADER_ROW = 8;
const FCST_TARGET_YEAR = 2026;
const FCST_TARGET_MONTHS = [5, 6, 7];
