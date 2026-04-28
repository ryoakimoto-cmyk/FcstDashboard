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
