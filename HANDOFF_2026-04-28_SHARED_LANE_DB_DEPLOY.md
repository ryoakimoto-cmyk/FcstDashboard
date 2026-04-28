# Shared lane handoff: snapshot DB deploy

Date: 2026-04-28
Repo: `C:\Users\RyoAkimoto\Documents\FcstDashboard`

## Deployment status

The snapshot DB-file separation work is deployed to the production Apps Script web app.

- Git commit: `a671418 Move snapshot storage to DB files`
- Apps Script version: `191`
- Deployment ID: `AKfycbxTk-eqvU_ObxIyGlOgzl_IglVID3bfQTfW_Wf2xlH_S0Ws52uTweXl4Gmfv1L0ANE3`
- Production URL: `https://script.google.com/macros/s/AKfycbxTk-eqvU_ObxIyGlOgzl_IglVID3bfQTfW_Wf2xlH_S0Ws52uTweXl4Gmfv1L0ANE3/exec`
- Description: `Move snapshot storage to DB files`

Verification commands already run:

- `cmd /c npm run validate:syntax`
- `cmd /c npm run validate:critical`
- `cmd /c clasp push -f`
- `cmd /c clasp version "Move snapshot storage to DB files"`
- `cmd /c clasp deploy -i AKfycbxTk-eqvU_ObxIyGlOgzl_IglVID3bfQTfW_Wf2xlH_S0Ws52uTweXl4Gmfv1L0ANE3 -V 191 -d "Move snapshot storage to DB files"`
- `cmd /c clasp deployments`

## Storage split

Main spreadsheet remains the metadata / input / operations file.

Main file stays responsible for:

- `SFデータ更新_BO`
- `SFデータ更新_SS`
- `SFデータ更新_SSCS`
- `SFデータ更新_CO`
- `組織マスタ`
- `所属マスタ`
- `目標マスタ`
- `SFユーザー`
- `Export待機`
- `Export待機_提案商品`
- `変更ログ`
- `集計キャッシュ`
- `AppDataCache`
- `SnapshotExecutionLog`
- `SnapshotDBIndex`

DB files now store app-owned growth/state data:

- `FCSTスナップショット`
- `案件リストスナップショット`
- `FCST調整`

## Implementation notes

`SnapshotStorage.gs` is the new storage layer. It creates DB spreadsheets with names like `FcstDashboard DB - <sheet name> - <timestamp>`, stores the active file ID in script properties, and rolls over to a new DB file before hitting the configured cell limit or after a cell-limit write error.

Reads are DB-only for the DB-owned sheets above. The app does not read legacy sheets from the main spreadsheet for `FCSTスナップショット`, `案件リストスナップショット`, or `FCST調整`. This project has not started production operation for this storage split, so old compatibility paths should be cleaned up instead of preserved.

Writes now go to DB files for the three DB-owned sheets above. `SnapshotExecutionLog` stays in the main file and now includes:

- `storage_file_id`
- `storage_file_url`
- `storage_sheet_name`
- `storage_rolled_over`

`SnapshotDBIndex` is also created in the main file when possible. If the main file is already too close to the 10M cell limit and index creation fails, DB writes still continue; the write location is still visible from `SnapshotExecutionLog`.

Fresh-start cleanup helpers are available:

- `SnapshotStorage_getFreshStartPlan()`: lists old main-spreadsheet DB-owned sheets, registered DB files, and script properties that would be cleared.
- `manualArmSnapshotStorageFreshStartCleanup()`: arms cleanup for 10 minutes and returns the current cleanup plan.
- `manualCleanupSnapshotStorageForFreshStart()`: after arming, deletes the old main-spreadsheet DB-owned sheets/index, trashes registered DB files, and clears snapshot storage script properties.

`SnapshotExecutionLog` is not deleted by cleanup because it is an operations log in the main spreadsheet. Unregistered DB-looking files found by Drive name search are listed as candidates only and are not automatically trashed.

## Manual verification

Use Apps Script editor manual execution. Do not use `clasp run` in this project.

Recommended smoke test:

1. Run `manualCreateFcstSnapshot_BO`.
2. Confirm `BOPA` and `BOW` no longer fail with the 10,000,000 cell limit error.
3. Confirm `SnapshotExecutionLog` has `storage_file_id`, `storage_file_url`, `storage_sheet_name`, and `storage_rolled_over`.
4. Open the `storage_file_url` and confirm `FCSTスナップショット` rows were written there.
5. Run `manualCreateOppSnapshot_BO` after FCST verification and check the same storage columns for `案件リストスナップショット`.

## Local caveats

There are unrelated local working-tree changes that were intentionally not included in the deploy commit:

- `.claspignore`
- `.claude/settings.local.json`
- `scripts/validate-gas-api.js`
- `.claude/hooks/`
- `_clasp_tmp/`

The `.claspignore` local change excludes `.claude/**`; it was used during `clasp push` so Claude hook files were not pushed to Apps Script.
