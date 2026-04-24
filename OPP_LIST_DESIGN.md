# Opp List Design

This file was refreshed on 2026-04-24 to avoid mixing the removed legacy snapshot path with the current implementation.

## Current snapshot design

- Canonical storage is `Opp履歴`
- Web App input merge source is `案件_WebApp入力`
- Batch run metadata is `スナップショットメタ`
- Internal implementation lives in `OppHistory.gs`
- Public entrypoints stay in `OppListSnapshot.gs`

## Compatibility note

- The old `案件リストスナップショット` write path is no longer part of the runtime design.
- `OppListSnapshot_getByDate()` still returns the legacy row shape so the current Opp UI can continue to read snapshot rows without a simultaneous UI rewrite.

## Reference

- See [OPP_HISTORY_TRACKING.md](/C:/Users/RyoAkimoto/.codex/worktrees/7f00/FcstDashboard/OPP_HISTORY_TRACKING.md) for the canonical tracking summary.
