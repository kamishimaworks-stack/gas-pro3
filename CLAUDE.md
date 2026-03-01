# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## プロジェクト概要
建築業向けAI見積自動作成SPA (v10.0)。バックエンドはGAS、フロントエンドはVue.js 3 (CDN)。
DB=Spreadsheet、ストレージ=Drive、AI=Gemini API、通知=LINE Messaging API。
業務フロー: 見積作成 → 発注書作成 → 請求書受領/発行 → 入出金管理 → 出来高報告 → 会計仕訳エクスポート

## 開発コマンド
- デプロイ: `clasp push --force` / 取得: `clasp pull` / エディタ起動: `clasp open`
- E2Eテスト: `node test-e2e.js` (Playwright, headless chromium, ja-JP locale)
- Git: `kamishimaworks-stack/gas-pro3.git` (mainブランチ)

## ファイル構成とGAS制約
- `コード.js`: 唯一のバックエンド(~5000行)。`require`/`import`不可、全てグローバル関数。V8ランタイム。
- `index.html`: Vue 3 Composition APIのSPA(~8200行)。SFC不可、CDN読み込みで`<script>`直書き。
- `app_script.html`: `index.html`にincludeされるJS/CSS。
- `*_template.html`: PDF生成用テンプレート(GAS構文 `<? ?>` 使用)。6種: quote, order, bill, ledger, progress_report, logo。

## バックエンド (`コード.js`)

### API規約
- フロントから呼ぶ関数は `api` プレフィックス必須（例: `apiSaveEstimate`）。
- プライベートヘルパーは `_` サフィックス（例: `readProgressItems_`）。
- 全APIは `JSON.stringify(...)` で文字列を返却。フロント側で `JSON.parse()`。
- 書き込み系APIは `LockService.getScriptLock()` で排他制御。

### キャッシュ戦略
- `CacheService.getScriptCache()` 使用。TTL: マスタ=25分(`CACHE_TTL=1500`), PJ=2分(`CACHE_TTL_SHORT=120`), 発注=1分(`CACHE_TTL_ORDERS=60`)。
- 保存時: `invalidateDataCache_()` で汎用キャッシュ破棄。
- 出来高系: `invalidateProgressCache_(orderId)` で月±2ヶ月分の関連キーを破棄。
- 主要キャッシュキー: `projects_data`, `orders_data`, `masters_data`, `progress_data_{orderId}_{ym}`, `progress_report_list_{ym}`

### DB操作パターン
- Spreadsheetのカラム位置はインデックス直指定（マジックナンバー）。列順序変更時は全参照箇所を要確認。
- 出来高DBは `PROG_COL` 定数(1-based)で列管理: no:1, name:2, spec:3, totalQty:4, unit:5, price:6, estimateAmt:7, prevCumQty:8, currCumQty:9, progressAmt:10, progressRate:11, periodQty:12, periodPayment:13, estId:14, orderId:15, reportMonth:16
- 数式列(estimateAmt, progressAmt, progressRate, periodQty, periodPayment)は `setProgressFormulas_()` のARRAYFORMULAで自動計算。手動セット不可。
- `PROGRESS_HEADERS.length` で読み書き範囲が自動決定されるため、列追加時はHEADERS配列・PROG_COL・newRows・readItems返却全ての同期が必須。

### ScriptProperties
- `PROGRESS_REPORT_HEADERS`: 出来高ヘッダー情報(orderId → headerData のJSONマップ)
- `ADMIN_USERS`: 管理者メールリスト

### 主要APIセクション (コード.js内の位置)
| セクション | 主なAPI | 説明 |
|---|---|---|
| 認証・マスタ | apiGetAuthStatus, apiGetMasters, apiGetUnifiedProducts | 4ソース統合(基本/元請別/セット/材料) |
| 見積 | apiGetProjects, apiSaveUnifiedData, apiSaveAndCreateEstimatePdf | 保存時に発注を自動生成 |
| 発注 | apiGetOrders, apiCreateOrderPdf, apiGetOrderDetails | 発注先単位でPDF生成 |
| 請求書 | apiParseInvoiceFile, apiSaveInvoice, apiGetInvoices | OCR(Gemini)解析対応 |
| 入出金 | apiGetDeposits/Payments, apiSaveDeposit/Payment | 見積/発注に紐付け |
| 出来高 | apiProgressGetItems, apiProgressMonthlyCloseOrder, apiProgressGeneratePdf | 月別データ管理・工事単位月締め |
| AI | apiAiAssistantUnified, apiPredictUnitPrice, apiChatWithSystemBot | Gemini API統合 |
| 仕訳 | apiGenerateJournalData, apiGetAnalysisData | CSV出力・年度分析 |
| LINE | processLineEvent_, apiTestLineSend | Webhook + プッシュ通知 |

## フロントエンド (`index.html`)

### GAS通信
- `gas('api関数名', arg1, arg2, ...)` ラッパー使用。
- 非GAS環境(ローカル開発): `google.script` 未検出時、モックデータへフォールバック（setup()内のif/else ifチェーン）。
- モック更新時: funcName分岐内の対応する `else if` ブロックを修正。

### 画面構成 (`currentTab`)
| タブ値 | 画面名 | 概要 |
|---|---|---|
| `menu` | ホーム | メニューカード、リマインダー、統計 |
| `list` | 案件一覧 | 見積/発注リスト、フィルタ、サブタブ切替 |
| `edit` | 見積編集 | 3パネル: 左(マスタ/セット検索), 中央(明細), 右(発注作成) |
| `dedicated_order` | 発注書編集 | 独立した発注作成画面 |
| `progress` | 出来高報告一覧 | 年月・発注先・工事名フィルタ |
| `progress_edit` | 出来高報告編集 | 月セレクタ、品目編集、月締め、PDF生成 |
| `invoice` | 請求書受取処理 | OCR + 手動入力、Drive連携 |
| `admin` | 管理画面 | 承認・分析・設定(adminSubTab切替) |
| `master` | マスタデータ編集 | 各種マスタのCRUD |

### 主要状態オブジェクト
- `form` reactive: 見積ヘッダー+明細 (`form.header`, `form.items[]`)
- `dedicatedOrderForm` reactive: 発注書ヘッダー+明細
- `orderInPanel` reactive: 見積編集の右パネル発注
- `masters` reactive: clients[], sets[], vendors[]
- `progressItems` ref: 出来高品目リスト
- `progressEditYear/Month` ref + `progressEditYM` computed: 編集画面の報告月("YYYY-MM")
- `isDirty` ref: 未保存変更検知（ページ離脱警告）
- `progressDirtyRows` ref(Set): 出来高の未保存行追跡

### UIライブラリ
- Vue 3 (CDN global build), Tailwind CSS (CDN), Chart.js (defer), Material Icons

## 出来高報告の月別データモデル

### 月締めフロー (`apiProgressMonthlyCloseOrder`)
1. orderId + currentYM(または空reportMonth)の行を取得
2. reportMonth空の行にcurrentYMをタグ付け（既存行の報告月を確定）
3. 翌月行を新規作成: prevCumQty=旧currCumQty, reportMonth=nextYM
4. 二重締め防止: nextYMデータ存在チェック
5. フロント側: 成功時に自動的に翌月表示に切替

### フィルタの仕組み
- 一覧画面: `loadProgressReportList()` がAPIに"YYYY-MM"を渡し、バックエンドでフィルタ
- 編集画面: `loadProgressEditData(orderId, ym)` が月指定でデータ取得
- `!item.reportMonth || item.reportMonth === ym` パターンで空月(未タグ)も含める

## PDF生成フロー
1. バックエンドでデータを `{header, items, pages}` に構造化。
2. `HtmlService.createTemplateFromFile('テンプレート名')` で展開。
3. 生成HTMLをDriveへPDF保存し、URLをフロントへ返却。
4. `paginateItems(items, firstPageRows, otherPageRows)` でページ分割。

## データモデル関係図
```
見積(Estimate) ──1:N──→ 発注(Order) ──1:N──→ 出来高(Progress)
     │                      │                    │
     └──1:N──→ 入金         └──1:N──→ 出金       └── 報告月(YYYY-MM)で月別管理
                             └──1:N──→ 受取請求書
```

## ログ・出力の言語ルール
- **console.log / Logger.log などのログメッセージは日本語で記述する。**
- ただし、**ファイル名・関数名・変数名・パス**はそのまま英数字(コード上の表記)で出力する。
- 例: `console.log('保存完了: ' + fileName)` / `Logger.log('apiSaveEstimate: 見積データを保存しました')`
