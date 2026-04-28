# Development Guardrails

Last updated: 2026-04-28

## Fallback 判定の追加ルール

- 安直な fallback 判定、推定判定、緩和判定を無断で追加しない。
- 既存データが複数形に見える場合でも、先に実データ、既存 contract、想定される影響範囲を確認する。
- fallback が必要に見える場合は、実装前に次を明示する。
  - どの実データが現行 contract と合っていないか
  - fallback なしで何が壊れるか
  - fallback によって FCST / Opp / MRR / cache / snapshot のどこへ影響するか
  - payload shape、cache key、sheet schema を変えないことの確認
  - 追加後の検証方法と rollback 方針
- ユーザーが明示承認した場合、または仕様書に明記されている場合だけ runtime fallback を入れる。
- 一時調査用の fallback や debug 分岐を本番コードに残さない。
- fallback を入れた場合は、commit message または最終報告で「なぜ必要だったか」「どの範囲に効くか」を明記する。

## Snapshot / cache の変更ルール

- snapshot の読み取り条件を変更する前に、実際の保存形式と生成側の code path を確認する。
- cache key と payload shape は安易に変更しない。
- 表示不具合の修正であっても、サーバー側の contract を緩める場合は影響範囲を先に明示する。
- `npm run validate:syntax` と `npm run validate:critical` を通してから deploy する。

## Apps Script deploy 前の共有資材確認

- `clasp push -f` は Apps Script 側のファイル一覧を現在の checkout で置き換える。古い lane branch から push すると、別 branch で追加済みの共有資材を本番から消すリスクがある。
- deploy 前に、現在 branch が本番運用中の共有基盤を含んでいるか確認する。特に snapshot / cache / MRR / FCST を触る場合は `SnapshotStorage.gs`、storage-aware reader、5 分 cache 経路、URL routing の有無を確認する。
- branch が `master` または production deploy の基準 commit より古い場合、先に必要な共有基盤を取り込む。古い branch のまま画面修正だけを push/deploy しない。
- snapshot 保存先が DB ファイル化されている場合、表示側は main spreadsheet 直読みではなく storage-aware reader を使う。FCST snapshot は `FcstSnapshot_getAllValues_()`、DB 共通読み込みは `SnapshotStorage_getAllValues_()` を優先する。
- MRR / FCST / Opp の表示が `データなし` になった時は、UI修正を繰り返す前に「生成側の保存先」と「表示側の読み込み先」が一致しているかを最優先で確認する。
- deploy 前 validation には、重要共有ファイルの存在確認と、DB保存済み snapshot を main sheet 直読みしていないことの検査を含める。

## MRR ダッシュボード読み込みルール

- SS / BO / CO / COO は同一の FCST スナップショット集計ロジックを使う。事業部ごとの専用ロジックや旧 SS 専用シートへの分岐を追加しない。
- MRR 画面は未指定 URL では事業部選択画面を出し、選択後は `?app=mrr&division=SS|BO|CO|COO` の専用 URL にする。
- 初期読み込みは選択された各事業部につき直近 2 snapshot_date までに制限する。COO は SS / BO / CO それぞれ最大 2 日付まで。
- それより古い履歴はユーザー操作時だけ追加取得する。初期ロードで全履歴を読まない。
- Key Deal は初期 payload に混ぜない。表示中の日付・部署で必要になった時だけ Opp スナップショットから取得する。
- 5 分キャッシュを前提にする。ただし cache key は MRR 専用に分け、FCST / Opp 本体の cache key や payload shape を変更しない。
