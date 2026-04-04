# 案件リスト Web App 設計メモ

## 目的

FCST 運用のうち `案件リスト (最新)` 相当を Web App に置き換える。

今回の対象:

- 読取元を `SFデータ更新` シートへ統一
- Web App 上で案件リストを表示
- `Key Deal` と各 `SF更新用` をインライン編集
- 保存時に `Export待機` / `Export待機_提案商品` へ転記
- 毎週月曜 03:00 にスナップショットを保存
- Web App 上で週次スナップショットを左右遷移

## 現状把握

### 既存 GAS 側

`ForecastCommon` に既存運用ロジックがある。

- `08_OppListRefresher.js`
  - `SFデータ更新` を元に `案件リスト (最新)` を生成
- `07_ExportWaitingSync.js`
  - `(最新)` シートの編集内容を `Export待機` へ転記
- `14_ProposalProductsTransferService.js`
  - `(最新)` シートの `受領 / 債権管理 / 債権管理 Lite / 経費` を `Export待機_提案商品` へ転記
- `06_ArchiveService.js`
  - `案件リスト (最新)` を週次アーカイブ

### 現在の Web App 側

- `OppReader.gs`
  - 読取元はまだ `案件リスト (最新)`
- `FcstWriter.gs`
  - 一部の `SF更新用` セルへ即時保存している
- `js.html`
  - 案件リスト表示とインライン編集の土台あり
  - ただし「保存ボタンでまとめて転記」ではなく即時保存寄り
- `FcstSnapshot.gs`
  - FCST 集計用スナップショットは実装済み
  - 案件リスト用スナップショットは未実装

## 方針

案件リストは `(最新)` シート依存をやめ、Web App 専用の読取モデルに切り替える。

重要な制約:

- `C:\Users\RyoAkimoto\Box\ryo.akimoto\Forecast\GAS` 配下の既存 `ForecastCommon` / `Export待機` 系 GAS は改修しない
- 既存 GAS は仕様参照のみとし、実装は `C:\Users\RyoAkimoto\Documents\FcstDashboard` 側で完結させる
- Web App の表示時参照は `SFデータ更新` のみとする
- `Export待機` / `Export待機_提案商品` は保存先キューとしてのみ扱う

責務を 4 つに分ける。

1. `SFデータ更新` から表示用案件データを生成
2. Web App で編集差分を保持
3. 保存時に `Export待機` / `Export待機_提案商品` へ反映
4. 週次で案件一覧スナップショットを保存・参照

## 表示マッピング

| 案件リスト | SFデータ更新 |
|---|---|
| 完了予定月 | 完了予定月 |
| 担当部署 | 担当部署 |
| 種別 | 種別 |
| 案件副担当 | ユーザー |
| フェーズ | フェーズ_変換 |
| フォーキャスト | フォーキャスト_変換 |
| 予定日 / 確定日 | 予定日 / 確定日 |
| 案件名 | 案件名 |
| 計上割合 (%) | パーセント (%) |
| MRR | 月額(換算値) |
| 初期費用 | 初期費用額(換算値) |
| Key Deal | KeyDeal_最新 |
| FCST (コミット) | FCST(コミット)_最新 |
| FCST (MIN) | FCST(MIN)_最新 |
| FCST (MAX) | FCST(MAX)_最新 |
| 受領 | FCST(コミット)_受領 |
| 債権管理 | FCST(コミット)_債権管理 |
| 債権管理 Lite | FCST(コミット)_債権管理 Lite |
| 経費 | FCST(コミット)_経費 |
| FCSTコメント | FCSTコメント_最新 |
| 初回商談日 | 初回営業日 / ｺﾝﾀｸﾄ日 |
| 案件概要 | 案件概要 |
| オプション機能 | オプション機能 |
| 施策 | 施策 |

## データモデル

```js
{
  oppId: '',
  completedMonth: '',
  dept: '',
  type: '',
  subOwner: '',
  phase: '',
  forecast: '',
  scheduleOrCloseDate: '',
  dealName: '',
  allocationPercent: 0,
  mrr: 0,
  initialCost: 0,
  keyDeal: false,
  fcstCommit: { sfValue: 0, draftValue: 0 },
  fcstMin: { sfValue: 0, draftValue: 0 },
  fcstMax: { sfValue: 0, draftValue: 0 },
  received: { sfValue: 0, draftValue: 0 },
  debtMgmt: { sfValue: 0, draftValue: 0 },
  debtMgmtLite: { sfValue: 0, draftValue: 0 },
  expense: { sfValue: 0, draftValue: 0 },
  proposalProductIds: {
    received: '',
    debtMgmt: '',
    debtMgmtLite: '',
    expense: ''
  },
  fcstComment: '',
  firstMeetingDate: '',
  summary: '',
  optionFeatures: '',
  initiative: '',
  snapshotDate: null
}
```

補足:

- `sfValue`: `SFデータ更新` 上の最新値
- `draftValue`: Web App 上で編集する値
- 保存対象は `keyDeal` と各 `draftValue`
- 行キーは `oppId`
- `proposalProductIds.*` は `Export待機_提案商品` upsert に必要

表示時の初期値ルール:

- `sfValue` は `SFデータ更新` の値
- `draftValue` も初期表示は `SFデータ更新` の値
- 未処理キュー値の逆参照はしない

## シート責務

### 1. `SFデータ更新`

唯一の表示元。

- 最新表示は常にここから取得
- `最新取得` ボタン押下時もここを再読込

### 2. `Export待機`

Salesforce 書き戻し待ちのキュー。

既存 `ForecastCommon/07_ExportWaitingSync.js` と責務を合わせる。

列構成:

- A: 案件 ID
- B: Key Deal
- C: FCST(コミット)
- D: FCST(MIN)
- E: FCST(MAX)
- F: FCSTコメント
- G: 予備
- H: Status

保存対象:

- `Key Deal`
- `FCST(コミット)`
- `FCST(MIN)`
- `FCST(MAX)`
- `FCSTコメント`

参照用途には使わない。

### 3. `Export待機_提案商品`

`受領 / 債権管理 / 債権管理 Lite / 経費` の保存先。

既存 `ForecastCommon/14_ProposalProductsTransferService.js` で確認した内容:

- `受領` -> `提案商品ID_受領`
- `債権管理` -> `提案商品ID_債権管理`
- `債権管理 Lite` -> `提案商品ID_債権管理 Lite`
- `経費` -> `提案商品ID_経費`

列構成:

- A: 提案商品ID
- B: モジュール名
- C: 金額
- D: 案件ID
- E: 予備
- F: Status

upsert 仕様:

- `提案商品ID_*` が既存行にあれば update
- なければ新規 append
- 更新判定は C 列の価格差分
- 更新時は F 列に `-` を設定

参照用途には使わない。

### 4. 新設 `案件リストスナップショット`

案件リスト用の週次保存先として新設。

推奨構造:

| A | B | C | D | E |
|---|---|---|---|---|
| snapshot_at | snapshot_date | opp_id | dept | payload_json |

`payload_json` に 1 行分の案件データを保存する。

理由:

- 表示項目追加に強い
- シート列増減の影響を受けにくい
- 週次遷移表示を組みやすい

## サーバー構成

### 読取系

- `OppListReader_getLiveRows()`
  - `SFデータ更新` から最新案件一覧を返す
- `OppListSnapshot_getSnapshotDates()`
  - 保存済み週次一覧を返す
- `OppListSnapshot_getByDate(dateStr)`
  - 指定週の案件一覧を返す

### 保存系

- `OppListWriter_saveDrafts(changes)`
  - `Key Deal / FCST系 / FCSTコメント` を `Export待機` に upsert
  - `受領 / 債権管理 / 債権管理 Lite / 経費` を `Export待機_提案商品` に upsert
- `OppListSnapshot_createWeekly()`
  - 月曜 03:00 時点の live rows をスナップショット保存
- `OppListSnapshot_setupWeeklyTrigger()`
  - 週次トリガー設定

## クライアント構成

### 画面

案件リストタブを 2 モードにする。

- 最新
  - `SFデータ更新` ベース
  - 編集可
- スナップショット
  - 左右遷移で週切替
  - 編集不可

### 操作

- `最新取得`
  - `SFデータ更新` を再取得して再描画
- セル編集
  - ローカル state のみ更新
- `保存`
  - 差分のみサーバー送信

### テーブル UI

- `thead th { position: sticky; top: 0; }`
- 横スクロール用 wrapper と縦スクロール領域を分ける
- 2 行ヘッダで親見出しと `SF上の値 / SF更新用` を表示
- `Key Deal` は checkbox
- 金額 `SF更新用` は inline editor

## ヘッダ仕様

1 行目:

- 完了予定月
- 担当部署
- 種別
- 案件副担当
- フェーズ
- フォーキャスト
- 予定日 / 確定日
- 案件名
- 計上割合(%)
- MRR
- 初期費用
- Key Deal
- FCST(コミット)
- FCST(MIN)
- FCST(MAX)
- 受領
- 債権管理
- 債権管理 Lite
- 経費
- FCSTコメント
- 初回商談日
- 案件概要
- オプション機能
- 施策

2 行目:

- 金額系ブロックのみ `SF上の値` / `SF更新用`

## 編集仕様

### 編集可

- `Key Deal`
- `FCST(コミット)` の `SF更新用`
- `FCST(MIN)` の `SF更新用`
- `FCST(MAX)` の `SF更新用`
- `受領` の `SF更新用`
- `債権管理` の `SF更新用`
- `債権管理 Lite` の `SF更新用`
- `経費` の `SF更新用`
- `FCSTコメント`

### 編集不可

- `SF上の値`
- `SFデータ更新` 由来の属性列
- スナップショット表示中の全項目

### 保存単位

```js
[
  {
    oppId: '006xxxxxxxxxxxx',
    keyDeal: true,
    proposalProductIds: {
      received: 'a1Bxxxxxxxxxxxx',
      debtMgmt: 'a1Bxxxxxxxxxxxx',
      debtMgmtLite: 'a1Bxxxxxxxxxxxx',
      expense: 'a1Bxxxxxxxxxxxx'
    },
    fields: {
      fcstCommit: 1200000,
      fcstMin: 1000000,
      fcstMax: 1500000,
      received: 200000,
      debtMgmt: 300000,
      debtMgmtLite: 0,
      expense: 50000,
      comment: '...'
    }
  }
]
```

## 保存フロー

1. 画面表示時に live rows を取得
2. ユーザーが `Key Deal` / `SF更新用` を編集
3. フロントで差分管理
4. 保存ボタンクリック
5. `OppListWriter_saveDrafts(changes)` 実行
6. `Export待機` に `oppId` 単位で upsert
7. `Export待機_提案商品` に `提案商品ID_*` 単位で upsert
8. 成功後 live rows を再取得

再取得元は `SFデータ更新` のみ。

返却値イメージ:

```js
{
  exportWaiting: { newCount: 1, updatedCount: 2 },
  exportProposal: { newCount: 0, updatedCount: 3 }
}
```

## 週次スナップショット

### 保存タイミング

- 毎週月曜 03:00 Asia/Tokyo

### 保存内容

- その時点の live rows 全件
- `SF上の値`
- `draftValue` 相当
- `keyDeal`
- コメント

### 表示

- 「今週」は live 表示
- 過去週は snapshot 表示
- 左右ボタンで週移動

## 実装ステップ案

### Step 1

読取元を `案件リスト (最新)` から `SFデータ更新` へ切り替える。

- `OppReader.gs` は置換ではなく新規 `OppListReader.gs` に分離
- まずは表示専用で live rows を返す

### Step 2

フロントの案件リストを 2 行ヘッダへ置換する。

- sticky header
- checkbox
- inline editor
- 保存バーを「クリア」中心から「保存」中心へ変更

### Step 3

差分保存を実装する。

- 既存の `saveOppSfValue` 即時保存は廃止
- `saveOppDrafts` 一括保存へ変更
- `Export待機` upsert 実装
- `Export待機_提案商品` upsert 実装

### Step 4

案件リスト用スナップショットを追加する。

- 保存関数
- 取得関数
- 週次トリガー
- 左右遷移 UI

## 既知の確認事項

以下はまだ確認が必要。

1. `提案商品ID_受領 / 債権管理 / 債権管理 Lite / 経費` の読取元
   - `SFデータ更新` に存在する前提で実装済み
2. `FCSTコメント` は `Export待機` F 列で確定か

## 推奨判断

- 表示元は常に `SFデータ更新`
- 編集初期値は常に `SF上の値`
- `Key Deal / FCST系 / FCSTコメント` は `Export待機` upsert
- `受領 / 債権管理 / 債権管理 Lite / 経費` は `Export待機_提案商品` upsert
- 案件リスト用スナップショットは新設シートへ JSON 保存
- 週次 snapshot では編集不可
