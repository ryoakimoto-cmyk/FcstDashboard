# FcstDashboard Development Guardrails

このファイルは作業開始前に必ず読む。特に shared / FCST / Opp lane をまたぐ作業では、ここにある制約を先に確認してから調査・編集する。

## Local Search Rules

- この Windows 環境では `rg` / `rg.exe` が `Access is denied` で失敗することがあるため、FcstDashboard 作業では `rg` を使わない。
- ファイル検索は PowerShell の `Get-ChildItem` を使う。
- 文字列検索は PowerShell の `Select-String` を使う。
- 複数ファイルを読む場合も、`rg` に戻さず `Get-ChildItem` と `Select-String` を組み合わせる。

例:

```powershell
Get-ChildItem -Recurse -Filter *.gs
Select-String -Path *.gs,*.html -Pattern 'SnapshotStorage_'
Get-ChildItem -Recurse -Include *.gs,*.html | Select-String -Pattern 'SnapshotStorage_'
```

## Parallel Agent Rules

- 調査・実装・検証を分離できる作業では、複数 agent を並行利用して最短完了を目指す。
- ただし、同じファイルを複数 agent に編集させない。編集範囲が衝突する場合は、親 agent が実装し、子 agent は調査・レビューに限定する。
- agent に渡す指示にも、このファイルの制約、特に `rg` 禁止とフォールバック禁止を明記する。

## Fallback Rules

- 集計・スナップショット・保存先判定に関するフォールバックは、ユーザー承認なしに追加しない。
- やむを得ずフォールバックが必要な場合は、影響範囲・発火条件・正規経路へ戻す条件を明示してから実装する。

## Read Performance Rules

- 読み込み対象が多岐にわたる場合でも、正規データソースを先に絞り、不要な legacy / fallback / 全件探索を通常経路に入れない。
- Web App 表示で使う重い読み込み・集計は、原則 5 分キャッシュを使ってユーザー表示負荷を下げる。
- キャッシュキーは部署・期間・表示モード・データ種別を明示し、異なる条件のデータを混ぜない。
- cache key / payload shape は安易に変えない。変更が必要な場合は影響範囲を先に明示する。
- キャッシュ miss 時だけ正規ソースを読みに行く。キャッシュ hit 時に Spreadsheet / Drive の全探索を併用しない。
- スナップショット・集計系では、読み込み範囲を日付・部署・期間で可能な限り先に絞る。
- FCST と Opp で共通化できるキャッシュ層・日付処理・部署フィルタは shared helper に寄せる。ただし payload shape が違うものを無理に同一化しない。

## Existing Feature Protection Rules

- 追加改修で既存機能を壊さないことを最優先する。特に FCST boot / Opp boot / 部署選択 / snapshot / save の初期表示経路は毎回保護対象として扱う。
- client 側の global helper を参照する場合は、同じ bundle 内で参照より前に定義されていることを確認する。
- boot に必要な client global helper は `scripts/validate-critical-flows.js` の検証対象に追加し、未定義・宣言順序ミスを deploy 前に検出する。
- `js.html` を編集した場合は、`npm run validate:syntax` と `npm run validate:critical` を必ず実行する。
- deploy 後は `clasp deployments` で production deployment の version を確認し、ユーザーに version と影響範囲を報告する。
- エラー修正時に安直な fallback で握りつぶさない。原因となる欠落・契約不一致を正規経路で直し、再発防止の検証を追加する。

## Apps Script Version Rules

- `clasp version` で新規versionを作成する前に、Apps Scriptのプロジェクト履歴で最古の不要versionを1つ削除する。
- version上限に到達してから気づく運用にしない。`clasp versions` で件数を確認し、上限付近または上限到達時は先に削除してからversion作成する。
- `clasp` にはversion削除コマンドがないため、削除はApps Scriptエディタのプロジェクト履歴画面で行う。
- 削除してよいのはproduction deploymentや直近rollback候補に使っていない古いversionだけ。削除前後に `clasp deployments` でproduction deploymentのversionを確認する。

## Apps Script Runtime Diagnostics Rules

- Do not use `clasp run` in this project. It is blocked by execution/API permissions and is not a reliable diagnostic path.
- For runtime data checks, use explicit manual functions in the Apps Script editor, deployed web app behavior, logs, or direct Drive/Sheets inspection with existing OAuth permissions.
- If a runtime diagnostic function is missing, add a bounded manual diagnostic function and deploy it through the normal validate -> push -> version -> deploy flow instead of trying `clasp run`.

## Snapshot DB Folder Rules

- Snapshot DB ファイルは元ファイル `SPREADSHEET_ID` と同じ親フォルダへ作成する。
- Google Drive API の GCP 有効化が必要な実装にはしない。`googleapis.com/drive/v3` 直叩きは禁止。
- フォルダ移動は Apps Script 組み込みの `DriveApp` を使う。初回のみ Apps Script の Drive 権限承認が必要。
- snapshot payload、cache key、既存シート schema を権限回避目的で変更しない。
