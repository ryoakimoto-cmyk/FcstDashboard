# Opp History Tracking

Last updated: 2026-04-24

## Current canonical path

- Public entrypoints:
  - `OppListSnapshot_createWeekly(deptKey?)`
  - `OppListSnapshot_getSnapshotDates(deptKey)`
  - `OppListSnapshot_getByDate(deptKey, dateStr)`
  - `OppListSnapshot_setupWeeklyTrigger()`
- Storage:
  - `Opp履歴`
  - `案件_WebApp入力`
  - `スナップショットメタ`
  - `FCST宣言_現在` (header only)
- Internal implementation:
  - `OppHistory.gs`

## Notes

- Legacy `案件リストスナップショット` write-path has been removed from the canonical flow.
- Existing UI-facing snapshot readers still receive the legacy row shape via an adapter so current Opp screens do not break.
- Weekly batch trigger remains Monday 03:00 Asia/Tokyo through `createOppSnapshot`.
- `monthly_dropout` rows are stored in `Opp履歴` with `_status = removed_from_p`.
