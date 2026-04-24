# Shared lane handoff

Date: 2026-04-24
Repo: `C:\Users\RyoAkimoto\Documents\FcstDashboard`
Purpose: hand over from the current lane to `shared lane`

## Current state

The recent snapshot / FCST / opp work is now split across backend and frontend changes in the main worktree.

Backend changes are present in:

- `AggregatedCache.gs`
- `Config.gs`
- `FcstSnapshot.gs`
- `OppListReader.gs`
- `OppListSnapshot.gs`
- `SfDataReader.gs`

Frontend changes from `FCST lane` and `Opp lane` are present in:

- `index.html`
- `js.html`

Validation:

- `npm run validate:syntax` passed on 2026-04-24

Not done:

- `git push`
- `clasp push`
- manual GAS verification after the latest frontend changes

## Backend changes already implemented

### Opp snapshot mapping

Implemented in:

- [OppListReader.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/OppListReader.gs:31)
- [OppListSnapshot.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/OppListSnapshot.gs:21)

Key points:

- opp snapshot fields now use the new source headers such as `案件 ID`, `フェーズ_変換`, `フォーキャスト_変換`, `パーセント (%)`, `月額(換算値)`, `初期費用額(換算値)`
- new opp fields added:
  - `fcstCommitReceived`
  - `fcstCommitDebtMgmt`
  - `fcstCommitDebtMgmtLite`
  - `fcstCommitExpense`
- `proposalProductIds` is removed from snapshot `payload_json`

### FCST aggregate source remap

Implemented in:

- [SfDataReader.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/SfDataReader.gs:18)

Current aggregate input mapping:

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

### New confirmed proposal-product metrics

Implemented in:

- [Config.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/Config.gs:45)
- [SfDataReader.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/SfDataReader.gs:59)
- [AggregatedCache.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/AggregatedCache.gs:117)
- [FcstSnapshot.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/FcstSnapshot.gs:35)

New metric keys:

- `confirmedReceived`
- `confirmedDebtMgmt`
- `confirmedDebtMgmtLite`
- `confirmedExpense`

Rule:

- source columns are `月額_受領`, `月額_債権管理`, `月額_債権管理 Lite`, `月額_経費`
- only rows where `フェーズ_変換 === '確定'` are aggregated into these keys

These keys are already included in:

- aggregate metric objects
- group / department totals
- aggregated cache payload
- FCST snapshot payload
- FCST snapshot week-over-week diff

## Frontend changes now present

### FCST lane result

Evidence in main worktree:

- [js.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/js.html:96)
- [js.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/js.html:114)
- [js.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/js.html:397)

Implemented summary:

- FCST UI now has frontend support for:
  - `confirmedReceived`
  - `confirmedDebtMgmt`
  - `confirmedDebtMgmtLite`
  - `confirmedExpense`
- labels used:
  - `確定受領`
  - `確定債権管理`
  - `確定債権管理 Lite`
  - `確定経費`
- FCST table / metric key arrays / state defaults / summary fallback / summary cards / trend table-detail side were updated

Open note from FCST lane:

- trend chart series were not made default-visible to avoid overcrowding

### Opp lane result

Evidence in main worktree:

- [index.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/index.html:112)
- [js.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/js.html:100)
- [js.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/js.html:2197)

Implemented summary:

- opp UI now has frontend support for:
  - `fcstCommitReceived`
  - `fcstCommitDebtMgmt`
  - `fcstCommitDebtMgmtLite`
  - `fcstCommitExpense`
- opp table columns were added
- opp diff handling was extended
- these new opp fields are display-only; save flow was intentionally not expanded

Open note from Opp lane:

- live save flow still uses `proposalProductIds` for edit/export behavior
- snapshot payload no longer carries `proposalProductIds`, and no UI dependency was found there

## Files to review first

If `shared lane` resumes from here, start with these:

1. [SfDataReader.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/SfDataReader.gs:18)
2. [OppListReader.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/OppListReader.gs:31)
3. [OppListSnapshot.gs](C:/Users/RyoAkimoto/Documents/FcstDashboard/OppListSnapshot.gs:21)
4. [js.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/js.html:91)
5. [index.html](C:/Users/RyoAkimoto/Documents/FcstDashboard/index.html:105)

## Recommended next steps for shared lane

1. Review `git diff` for the files listed above and confirm the UI wording/layout is acceptable.
2. Manually verify in GAS / Web UI:
   - FCST screen shows the new `確定*` metrics only for proposal-products-enabled departments.
   - opp screen shows the new `FCST(コミット)_*` columns and diff behavior correctly.
   - snapshot views still load without `proposalProductIds`.
3. If accepted, perform `git push` and `clasp push`.

## Related handoff

There is also a narrower lane-oriented memo here:

- [HANDOFF_2026-04-24_FCST_OPP_LANES.md](C:/Users/RyoAkimoto/Documents/FcstDashboard/HANDOFF_2026-04-24_FCST_OPP_LANES.md)

This shared-lane memo supersedes it for general continuation.

## Worktree caveat

Current dirty entries also include:

- `scripts/validate-gas-api.js`
- `_clasp_tmp/`

Treat those as unrelated unless you confirm otherwise.
