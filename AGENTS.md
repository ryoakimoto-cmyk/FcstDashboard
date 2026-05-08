# FcstDashboard Codex Rules

このファイルは、FcstDashboard の全Codexチャットで最初に読む共通ルールです。

## Workspace Model

今後の基本運用は次の通りです。

- 相談・調査・方針整理: `C:\Users\RyoAkimoto\Documents\FcstDashboard`
- 実装: 改修テーマごとの専用workspace
- 旧常設laneは使わない
  - `FcstDashboard-shared`
  - `FcstDashboard-fcst`
  - `FcstDashboard-opps`

テーマworkspaceでは、そのテーマに必要なファイルをまとめて扱ってよい。`shared / fcst / opp` の固定laneへ作業を戻さない。

## Handoff Rules

テーマ間で作業を渡すHandoffは、必ず共通Handoff置き場に作成する。

```text
C:\Users\RyoAkimoto\Documents\FcstDashboard-handoffs
```

- Handoffはテーマworkspace内だけに置かない
- 別テーマへ渡す最終報告では、Handoffファイルの絶対パスを必ず書く
- Handoffには、source workspace、branch、commit hash、変更ファイル、取り込み先、deploy可否、残リスクを明記する
- テーマworkspace内で一時的にHandoffを作った場合も、完了前に共通Handoff置き場へコピーまたは作成し直す
- 別テーマのHandoffを参照する場合は、workspace内検索ではなく、まず共通Handoff置き場の絶対パスを確認する
- Codex同士へ渡す実装指示・調査依頼・レビュー依頼・Handoffは、原則として英語で書く
- 相談用正本で作成するHandoffや、各テーマworkspaceで作成するHandoffも英語で書く
- ユーザーへの通常報告・途中経過・最終報告は日本語で書く
- 別Codexが読む成果物は英語で、実装対象・禁止事項・検証条件・完了条件を明確に書く

## User-First Design Rule

FcstDashboard の設計・実装・レビューでは、開発者都合ではなくユーザー価値を最優先する。

ユーザーがダッシュボードに求める基本価値は次の2つ。

1. 欲しい情報が瞬時に見られること
2. 情報が最新だと判断できること

そのため、次を完了条件に含める。

- `validate pass`、`HTTP 200`、`deploy済み` だけで完了扱いにしない
- 最初にユーザーが見るべき主要情報を定義し、それが初期表示で最優先に描画されることを確認する
- 重い履歴、詳細、Key Deal、補助チャートは、主要情報の初期表示をブロックしない
- 最終更新時刻、何分前のデータか、snapshot基準日、cache/trigger/summary状態を画面上で判断できるようにする
- stale、missing、更新中、部分失敗を正常データに混ぜず、ユーザー向け状態と管理者向け詳細を分けて表示する
- 「手動で関数を実行すれば見える」は通常運用として不可。初回認可・初回移行以外は自動復旧・自動更新を原則とする
- cache missing や source missing で画面全体を落とさない。主要情報が出せない場合でも、原因と復旧経路が分かる状態表示にする
- ダッシュボード改修では、主要情報表示までの体感時間、最新性表示、stale時の見え方を最終報告に含める
- 実装方針がユーザー価値と衝突する場合は、先に設計を見直す。開発者にとって都合のよいfallbackや空表示で済ませない

## Start-of-Task Checklist

作業開始時に必ず確認する。

1. 現在のフォルダ、ブランチ、未コミット差分
2. `DEVELOPMENT_GUARDRAILS.md`
3. ユーザーが渡したHandoff
4. 既存差分がある場合は、上書きせず内容を確認

非自明な調査・実装・検証では、複数エージェントを並行利用して高速化する。

## Multi-Agent Acceleration

調査・開発・レビューは、原則として複数エージェントを使って最短完了を狙う。

### Research

- 独立して調べられる論点は複数エージェントに分ける
- 例: server path、client path、snapshot/cache path、validator/deploy path
- 親エージェントは即座に非重複の調査を進め、サブエージェントの完了待ちだけで止まらない
- 調査結果は、実装方針・触るファイル・触らないファイル・検証方法まで整理する

### Development

- 実装がファイル単位・責務単位で分けられる場合は、workerエージェントに分担させる
- 分担時は、各workerの担当ファイルと禁止範囲を明示する
- 同じファイルを複数workerに編集させない
- 親エージェントは統合、競合解消、最終検証に責任を持つ

### Review

- 実装後は、可能な限りレビュー専用エージェントを使う
- レビュー観点は、バグ、回帰、source of truth違反、payload/cache key破壊、検証漏れを優先する
- レビュー指摘はファイルと行に紐づける
- 親エージェントはレビュー結果を鵜呑みにせず、必要な修正と検証を完了させる

### Agent Lifecycle

- 役割が完了したエージェントは、結果を回収したら速やかにcloseする
- 完了済み・不要になったエージェントを開いたまま放置しない
- 並行エージェント上限に達しないよう、親エージェントが起動中エージェント数を管理する
- 追加調査や再レビューが必要な場合は、既存エージェントを再利用できるか確認してから新規spawnする

### When Not To Spawn

- 1ファイルの小修正、単純な文面修正、明らかに数分で終わる確認では無理にspawnしない
- それ以外の非自明な調査・開発・レビューでは、ユーザーが明示しなくても原則spawnする
- 例外的にspawnしない場合は、理由を作業開始時に明示する

## Search and Shell Rules

このWindows環境では、FcstDashboard作業で `rg` / `rg.exe` を使わない。最初からPowerShell検索を使う。

- ファイル検索: `Get-ChildItem`
- 文字列検索: `Select-String`
- `rg` を試して失敗してから切り替える、という手順は禁止
- ユーザーへの作業ログで「rgが拒否されたので切り替えた」と報告しない
- Apps Script CLIは必要に応じて `cmd /c clasp ...`

## Implementation Rules

- 相談用正本では原則コード実装しない
- 実装はテーマworkspaceで行う
- 同じテーマの変更は1つのworkspaceに集約する
- 旧laneや退役済みフォルダを編集しない
- cache key、payload shape、source of truthを変える場合は影響範囲を明示する
- FCST snapshot payload の `keyDeals` を互換fallbackとして復活させない
- MRR/FCSTのhistorical Key Dealは案件リストスナップショットを正とする
- current Key Dealは現在案件リストを正とする

## Validation and Deploy

実装タスクでは、可能な限り次を実行してから完了報告する。

```powershell
cmd /c npm run validate:syntax
cmd /c npm run validate:critical
git diff --check
```

ユーザーが本番反映、push、deploy、リリースまで求めている場合は、検証後にApps Scriptの通常手順まで進める。

```powershell
cmd /c clasp push -f
cmd /c clasp version "..."
cmd /c clasp deploy -i <deploymentId> -V <version> -d "..."
cmd /c clasp deployments
```

deploy後は、versionと影響範囲を報告する。

ユーザーがdeployを求めていない場合は、検証完了で止めて、deploy未実施と明記する。

## Clasp Authentication Rule

`clasp` が `invalid_grant`、`invalid_rapt`、認証期限切れ、ログイン要求、権限エラーで失敗した場合は、回避策で長時間粘らない。

その場合は、認証用のWebページをユーザーが開ける状態にする。

優先手順:

1. `cmd /c clasp login` を実行してブラウザ認証を起動する
2. ブラウザが自動で開かない場合は、ターミナルに出た認証URLをユーザーへ提示する
3. ユーザーの認証完了後に、失敗した `clasp` コマンドから再開する

`clasp logout`、認証情報削除、別アカウントへの切替は、ユーザーが明示的に求めた場合だけ行う。

認証待ちの間は、原因調査や別手段の試行で時間を使わず、ユーザー認証を待つ。

## Apps Script Runtime Diagnostics Rule

このプロジェクトでは `clasp run` を使わない。

`clasp run` は実行/API権限や認証条件で失敗しやすく、この環境の診断経路として信頼しない。毎回試して失敗するのは時間の無駄なので禁止する。

実データ確認・バックフィル・手動診断が必要な場合は、次のいずれかを使う。

- Apps Scriptエディタ上で対象関数を手動実行する
- Web Appの実画面で動作確認する
- Apps Scriptの実行ログを確認する
- 必要なら、手動実行専用の診断関数を追加して validate -> push -> version -> deploy の通常経路で反映する

`clasp` CLIで使ってよい代表例は、`push`、`version`、`deploy`、`deployments`、`status`、`login`。`run` は使わない。

## Autonomous Completion Rule

ユーザーが「deployまで」「本番反映まで」「最後まで」進めてよいと言っているタスクでは、途中で確認待ちにしない。

その場合は、Handoff、`AGENTS.md`、`DEVELOPMENT_GUARDRAILS.md`、既存コードの契約に従って、実装・検証・push・version作成・production deploy・deployment確認まで完了させる。

作業場所は必ずそのテーマworkspaceに固定する。detached HEADや退役済みフォルダでは編集・検証・deployしない。

deploy込みタスクで必ず実行するもの:

```powershell
cmd /c npm run validate:syntax
cmd /c npm run validate:critical
git diff --check
cmd /c clasp push -f
cmd /c clasp version "..."
cmd /c clasp deployments
```

production deployment は、`cmd /c clasp deployments` で既存IDを確認して更新する。新規deploymentを作る必要がある場合だけ、理由を明示する。

最終報告には必ず次を含める。

- cwd
- branch
- `git status --short`
- 変更ファイル一覧
- 実行した検証と結果
- 作成したversion番号
- 更新したdeployment ID
- deploy先
- 残リスク

## Git Hygiene

- ユーザーや別作業の差分を勝手に戻さない
- 作業前後に `git status --short` を確認する
- 不要な旧worktreeを復活させない
- テーマ完了後は、統合・deploy・削除方針を報告する

## Retired Worktrees

旧lane差分の退避先:

```text
C:\Users\RyoAkimoto\Documents\FcstDashboard-retired-worktrees\20260507-183046
```

必要な場合は参照だけにし、そこでは編集しない。
