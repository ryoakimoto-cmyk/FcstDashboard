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

### When Not To Spawn

- 1ファイルの小修正、単純な文面修正、明らかに数分で終わる確認では無理にspawnしない
- それ以外の非自明な調査・開発・レビューでは、ユーザーが明示しなくても原則spawnする
- 例外的にspawnしない場合は、理由を作業開始時に明示する

## Search and Shell Rules

このWindows環境では、FcstDashboard作業で `rg` を使わない。

- ファイル検索: `Get-ChildItem`
- 文字列検索: `Select-String`
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
