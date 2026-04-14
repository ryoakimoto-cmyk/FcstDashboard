# FcstDashboard コンテキスト引き継ぎファイル

作成日: 2026-04-09

---

## プロジェクト概要

**目的**: FCSTシート（営業フォーキャスト）をGAS Webアプリとして再構築・UI洗練
**プロジェクトディレクトリ**: `C:\Users\RyoAkimoto\Documents\FcstDashboard\`
**GAS Script ID**: `1gl7-Q2lUPHZjGz029LYKmr-DW68Ai5hd2lp1fhYgDQCd5fPLrqG9wBV-`
**スプレッドシートID**: `1j_0TNBRhhfLLcr9DBr5mPLitiFRXYtsqohfQ-fxSPGg`

---

## ファイル構成

```
FcstDashboard/
├── Code.gs              # doGet / getInitData / createSnapshot 等のエントリポイント
├── Config.gs            # 定数定義（シート名・カラムオフセット等）
├── FcstAdjusted.gs      # 調整予測値の保存・読込
├── FcstSnapshot.gs      # 週次スナップショット管理
├── FcstPeriods.gs       # 期間定義ユーティリティ
├── FcstReader.gs        # FCSTシート読込
├── FcstWriter.gs        # FCSTシート書込
├── MonthlyMasterReader.gs
├── OppListReader.gs / OppListSnapshot.gs / OppListWriter.gs
├── OppReader.gs
├── SfDataReader.gs      # Salesforceデータ集計
├── SummaryReader.gs
├── TargetMigration.gs / TargetReader.gs / UserReader.gs
├── index.html           # メインHTMLテンプレート
├── css.html             # スタイルシート（<?!= include('css') ?>でinclude）
├── js.html              # フロントエンドJS（<?!= include('js') ?>でinclude）
├── appsscript.json      # webapp設定済み
└── .clasp.json
```

---

## アーキテクチャ・設計決定事項

### データフロー
1. Salesforce → Coefficient → スプレッドシート（SFデータシート）自動更新
2. GAS がスプレッドシートを読んで集計 → Webアプリにデータ配信
3. 調整値（fcstAdjusted）は別シート「FCST調整」に保存
4. 週次スナップショットは「FCSTスナップショット」シートに保存（52週分保持）

### 主要シート名（Config.gs）
- `FCST` - 元データシート
- `FCST調整` - ユーザー入力の調整値
- `FCSTスナップショット` - 週次スナップショット
- `SFデータ更新` - Coefficient経由SF同期
- `SFユーザー` / `目標` / `サマリ` / `案件リスト (最新)`

### データ構造キー
- 担当者×期間のキー: `"name|period"` （例: `"山田|2026-05"`）
- 期間形式: `"2026-05"` (YYYY-MM)
- 対象年度: 2026年、対象月: 5・6・7月

### appsscript.json の webapp 設定（済）
```json
{
  "webapp": {
    "executeAs": "USER_DEPLOYING",
    "access": "DOMAIN"
  }
}
```

---

## 前セッションで実装した変更（2026-04-02〜04-03）

### FcstAdjusted.gs
- スケーリング処理（×10000 / ÷10000）を**削除**し値をそのまま保存
- ヘッダー行（1行目）を確保、データ読込・更新を2行目以降に統一
- `FCST_STATE_HEADERS = ['担当者', '期間', 'Net', 'New+Exp', 'Churn', '更新日時', 'Note']`

### FcstSnapshot.gs
- シート作成時にヘッダー行を追加
- `FcstSnapshot_getLatestMembers()` 追加（live view の先週比計算用）

### js.html（UI）
- 「-」表示を「未入力」に変更
- New+Exp / Churn セクションに**目標値トグル**と**差分バッジ**追加
- Net の target バッジ表示調整
- Key Deal 表示コンポーネント
- メモ（Note）入力・表示機能

---

## 未解決の課題・次のステップ

### 1. Webアプリデプロイメント問題（最重要）
- `clasp deploy` が正しく Web App 型のデプロイを作れていない疑い
- `appsscript.json` の webapp セクションは現在設定済み（確認済み）
- **確認手順**:
  1. `cd C:\Users\RyoAkimoto\Documents\FcstDashboard && clasp push`
  2. Apps Script エディタで手動デプロイ確認、または
  3. `clasp deploy --description "test"` 後に `clasp deployments` でWeb App URLを確認

### 2. アクセス制限
- `Code.gs:doGet()` でドメイン制限: `@sansan.com` のみアクセス可
- `appsscript.json`: `"access": "DOMAIN"` 設定済み

### 3. 前回セッション終了時の状態
- ファイル修正は Codex により完了
- clasp push 後のデプロイメントURL確認・動作確認が未実施の可能性あり
- 「ページが見つかりません」エラーへの対処中に会話が終了

---

## よく使うコマンド

```bash
# FcstDashboard ディレクトリへ
cd C:\Users\RyoAkimoto\Documents\FcstDashboard

# GASへプッシュ
clasp push

# デプロイ
clasp deploy --description "説明"

# デプロイ一覧確認
clasp deployments

# ログ確認
clasp logs
```

---

## Codex連携ルール（CLAUDE.md より）

- **Claudeの役割**: 要件定義・設計・プラン・レビュー・調査
- **Codexの役割**: 実装全般（コードの追加・修正・削除はすべてCodexに委譲）
- 例外なし：規模を問わず実装はCodexへ

---

## 参照ファイル

- `~/.claude/CLAUDE.md` - グローバル設定
- `~/.claude/codex-workflow.md` - Codex連携ワークフロー詳細
- `~/.claude/projects/C--Users-RyoAkimoto/memory/MEMORY.md` - 記憶インデックス
- `OPP_LIST_DESIGN.md` - 案件リスト設計仕様
