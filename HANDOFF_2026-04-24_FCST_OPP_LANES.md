# Handoff for FCST lane / Opp lane

Date: 2026-04-24
Repo: `C:\Users\RyoAkimoto\Documents\FcstDashboard`
Status: local changes only. `git push` / `clasp push` not executed for the changes below.

## Scope
This handoff covers only the recent backend changes around:

- opportunity snapshot payload mapping
- FCST aggregate source-column mapping
- new FCST metrics for confirmed MRR proposal-product amounts

Frontend display work is intentionally left to `FCST lane` and `Opp lane`.

## What Changed

### 1. Opp snapshot backend mapping was updated
Files:

- [OppListReader.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/OppListReader.gs:31)
- [OppListSnapshot.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/OppListSnapshot.gs:19)

Current live/snapshot mapping for opp rows:

- `oppId` <- `案件 ID`
- `dealName` <- `案件名`
- `dept` <- `担当部署`
- `subOwner` <- `ユーザー`
- `completedMonth` <- `完了予定月`
- `type` <- `種別`
- `phase` <- `フェーズ_変換`
- `forecast` <- `フォーキャスト_変換`
- `scheduleOrCloseDate` <- `予定日 / 確定日`
- `allocationPercent` <- `パーセント (%)`
- `mrr` <- `月額(換算値)`
- `initialCost` <- `初期費用額(換算値)`
- `keyDeal` <- `KeyDeal_最新`
- `fcstCommit` <- `FCST(コミット)(換算値)`
- `fcstMin` <- `FCST(MIN)(換算値)`
- `fcstMax` <- `FCST(MAX)(換算値)`
- `received` <- `月額_受領`
- `debtMgmt` <- `月額_債権管理`
- `debtMgmtLite` <- `月額_債権管理 Lite`
- `expense` <- `月額_経費`
- `fcstCommitReceived` <- `FCST(コミット)_受領`
- `fcstCommitDebtMgmt` <- `FCST(コミット)_債権管理`
- `fcstCommitDebtMgmtLite` <- `FCST(コミット)_債権管理 Lite`
- `fcstCommitExpense` <- `FCST(コミット)_経費`
- `fcstComment` <- `FCSTコメント_最新`
- `firstMeetingDate` <- `初回営業日 / ｺﾝﾀｸﾄ日`
- `summary` <- `案件概要`
- `optionFeatures` <- `オプション機能`
- `initiative` <- `施策`

Snapshot behavior:

- `payload_json` always drops `proposalProductIds`
- `snapshotDate` is still added at snapshot-write time

Reference:

- [OppListReader.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/OppListReader.gs:54)
- [OppListSnapshot.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/OppListSnapshot.gs:21)

### 2. FCST aggregate source columns were updated
File:

- [SfDataReader.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/SfDataReader.gs:18)

Current aggregate mapping:

- period <- `完了予定月`
- owner <- `ユーザー`
- type bucket <- `種別（バケット）`
- `fcstCommit` <- `FCST(コミット)_分割後`
- `fcstMin` <- `FCST(MIN)_分割後`
- `fcstMax` <- `FCST(MAX)_分割後`
- `confirmed` <- `金額（LK＋新ソリューション）(換算値)`
- `expectedMrr` <- `期待MRR_分割後`
- `received` <- `FCST(コミット)_受領_分割後`
- `debtMgmt` <- `FCST(コミット)_債権管理_分割後`
- `debtMgmtLite` <- `FCST(コミット)_債権管理 Lite_分割後`
- `expense` <- `FCST(コミット)_経費_分割後`
- `keyDeals[].company` <- `取引先名`
- `keyDeals[].monthlyMrr` <- `金額（LK＋新ソリューション）(換算値)`
- `keyDeals[].phase` <- `フェーズ_変換`
- `keyDeals[].fcst` <- `FCST(コミット)_分割後`
- `keyDeals[].oppId` <- `案件 ID`

### 3. New confirmed-MRR proposal-product metrics were added
Files:

- [Config.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/Config.gs:45)
- [SfDataReader.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/SfDataReader.gs:59)
- [AggregatedCache.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/AggregatedCache.gs:117)
- [FcstSnapshot.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/FcstSnapshot.gs:35)

New metric keys:

- `confirmedReceived`
- `confirmedDebtMgmt`
- `confirmedDebtMgmtLite`
- `confirmedExpense`

Aggregation rule:

- source columns are `月額_受領`, `月額_債権管理`, `月額_債権管理 Lite`, `月額_経費`
- only rows where `フェーズ_変換 === '確定'` are included

Implementation details:

- metric objects now contain these four keys
- group totals / department totals also sum them
- finalized aggregate payload includes them
- FCST snapshot payload includes them
- week-over-week snapshot diff includes them
- proposal-products-disabled departments strip these keys from cache output as well

References:

- [SfDataReader.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/SfDataReader.gs:57)
- [SfDataReader.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/SfDataReader.gs:218)
- [SfDataReader.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/SfDataReader.gs:330)
- [SfDataReader.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/SfDataReader.gs:378)
- [FcstSnapshot.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/FcstSnapshot.gs:32)
- [FcstSnapshot.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/FcstSnapshot.gs:739)

## FCST lane TODO

Backend is ready, but the frontend does not expose the new confirmed proposal-product metrics yet.

Likely touchpoints:

- [js.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/js.html:91)
  - `proposalProductFields_()`
  - likely introduce a second helper or merged helper for `confirmedReceived` / `confirmedDebtMgmt` / `confirmedDebtMgmtLite` / `confirmedExpense`
- [js.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/js.html:116)
  - `getFcstMetricKeys_()`
- [js.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/js.html:357)
  - `expandedMetrics` default state
- [js.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/js.html:965)
  - `breakdownCols` in FCST table rendering
- [js.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/js.html:3087)
  - summary cards / summary fallback if these should appear there
- [js.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/js.html:1583)
  - trend/chart labels only if you want these metrics chartable

Suggested label mapping:

- `confirmedReceived` -> `確定_受領`
- `confirmedDebtMgmt` -> `確定_債権管理`
- `confirmedDebtMgmtLite` -> `確定_債権管理 Lite`
- `confirmedExpense` -> `確定_経費`

Open product decision for FCST lane:

- whether these four metrics should be shown as regular FCST breakdown columns
- whether they should appear in summary cards
- whether they should appear in trend charts

## Opp lane TODO

Opp snapshot backend is updated, but UI follow-up is still open if the new payload fields should be visible.

Potential touchpoints:

- [index.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/index.html:105)
  - opp table headers / diff headers if the new `fcstCommit*` fields should be shown
- [js.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/js.html:2095)
  - opp base row / diff handling
- [js.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/js.html:2133)
  - opp field label map

New opp fields available from backend:

- `fcstCommitReceived`
- `fcstCommitDebtMgmt`
- `fcstCommitDebtMgmtLite`
- `fcstCommitExpense`

Snapshot note for Opp lane:

- `proposalProductIds` is no longer present in `payload_json`
- if any snapshot comparison logic assumes it exists, remove that assumption

## Validation

Ran:

```bash
npm run validate:syntax
```

Result: passed.

## Worktree Notes

Current local modifications in this worktree include:

- backend changes from this task:
  - `AggregatedCache.gs`
  - `Config.gs`
  - `FcstSnapshot.gs`
  - `OppListReader.gs`
  - `OppListSnapshot.gs`
  - `SfDataReader.gs`
- pre-existing dirty files also remain:
  - `index.html`
  - `js.html`
  - `scripts/validate-gas-api.js`
  - `_clasp_tmp/`

Do not assume the pre-existing dirty files belong to this backend task.
