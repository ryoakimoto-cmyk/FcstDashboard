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

## MRR ダッシュボード読み込みルール

- SS / BO / CO / COO は同一の FCST スナップショット集計ロジックを使う。事業部ごとの専用ロジックや旧 SS 専用シートへの分岐を追加しない。
- MRR 画面は未指定 URL では事業部選択画面を出し、選択後は `?app=mrr&division=SS|BO|CO|COO` の専用 URL にする。
- 初期読み込みは選択された各事業部につき直近 2 snapshot_date までに制限する。COO は SS / BO / CO それぞれ最大 2 日付まで。
- それより古い履歴はユーザー操作時だけ追加取得する。初期ロードで全履歴を読まない。
- Key Deal は初期 payload に混ぜない。表示中の日付・部署で必要になった時だけ Opp スナップショットから取得する。
- 5 分キャッシュを前提にする。ただし cache key は MRR 専用に分け、FCST / Opp 本体の cache key や payload shape を変更しない。
