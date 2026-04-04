const SPREADSHEET_ID = '1RJkk_KIh7orvzP4RmW5UdegrJnik0U4PwrFzrMZJfEk';

// 完全一致で検索（部分一致だと FCST_正規化 等に誤マッチするため）
const FCST_SHEET_NAME    = 'FCST';
const OPP_SHEET_NAME     = '案件リスト (最新)';
const SUMMARY_SHEET_NAME = 'サマリ';

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
  notes: 32,
};

const BLOCK_COL_STARTS = { Q: 4, M5: 38, M6: 72, M7: 106 };
const ROWS_PER_SNAPSHOT = 32;
const SNAPSHOT_START_ROW = 4;
const OPP_HEADER_ROW = 8;

const SF_DATA_SHEET_NAME = 'SFデータ更新';
const SF_USER_SHEET_NAME = 'SFユーザー';
const TARGET_SHEET_NAME = '目標';
const FCST_TARGET_YEAR = 2026;
const FCST_TARGET_MONTHS = [5, 6, 7];
const FCST_ADJUSTED_SHEET_NAME = 'FCST\u8abf\u6574';
const FCST_SNAPSHOT_SHEET_NAME = 'FCST\u30b9\u30ca\u30c3\u30d7\u30b7\u30e7\u30c3\u30c8';
