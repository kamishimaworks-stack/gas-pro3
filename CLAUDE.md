# CLAUDE.md

## プロジェクト概要
建築業向けAI見積自動作成SPA。バックエンドはGAS、フロントエンドはVue.js 3。
DB=Spreadsheet、ストレージ=Drive、AI=Gemini API。
業務フロー: 見積作成 → 発注書作成 → 請求書受領/発行 → 入出金管理 → 会計仕訳エクスポート

## 開発コマンド
- デプロイ: `clasp push` / 取得: `clasp pull` / エディタ起動: `clasp open`
- E2Eテスト: `node test-e2e.js` (Playwright + ローカルモック)
- Git: `kamishimaworks-stack/gas-pro3.git` (mainブランチ)

## ファイル構成とGAS制約
- `コード.js`: 唯一のバックエンド(約3300行)。`require`/`import`不可、全てグローバル関数。
- `index.html`: Vue 3 Composition APIのSPA(約4300行)。SFC不可、CDN読み込みで`<script>`直書き。
- `app_script.html`: `index.html`にincludeされるJS/CSS。
- `*_template.html`: PDF生成用テンプレート(GAS構文 `<? ?>` 使用)。

## バックエンド (`コード.js`)
- **API関数**: フロントから呼ぶ関数は `api` プレフィックス必須。
- **排他制御**: 新規ID生成(`getNextSequenceId`)は `LockService` を使用。
- **キャッシュ**: `CacheService`活用(マスタ25分, PJ2分, 発注1分)。保存時に `invalidateDataCache_()` で破棄。
- **DB操作**: Spreadsheetのカラム位置はインデックス直指定(マジックナンバー)。列順序変更時は要注意。
- **データ統合**: マスタ追加時は `apiGetUnifiedProducts()` (4ソース統合)を更新。

## フロントエンド (`index.html`)
- **GAS通信**: `gas('api関数名', args)` ラッパー使用。非GAS環境はモックへフォールバック。
- **画面 (`currentTab`)**: `menu`(ホーム), `list`(一覧/統計), `edit`(見積編集), `admin`(管理), `invoice`(受取請求書OCR), `order`(発注)。
- **見積編集(`edit`)**: 3パネル構成(左:マスタ/セット検索, 中央:明細, 右:発注作成)。
- **状態管理**: `form`(見積), `orderInPanel`(発注), `masters`(マスタ), `isDirty`(未保存検知)。

## PDF生成フロー
1. バックエンドでデータを `{header, items, pages}` に構造化。
2. `HtmlService.createTemplateFromFile('テンプレート名')` で展開。
3. 生成HTMLをDriveへPDF保存し、URLをフロントへ返却。

## ログ・出力の言語ルール
- **console.log / Logger.log などのログメッセージは日本語で記述する。**
- ただし、**ファイル名・関数名・変数名・パス**はそのまま英数字(コード上の表記)で出力する。
- 例: `console.log('保存完了: ' + fileName)` / `Logger.log('apiSaveEstimate: 見積データを保存しました')`