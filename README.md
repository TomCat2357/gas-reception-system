# 📊 Google AppSheetスタイル 受付データ管理アプリ

Google Apps Script 製の Web アプリ。スプレッドシートに **多段（9 層）ヘッダー** で JSON を展開し、受付データを登録・更新できます。

## 🌟 概要

- **Web アプリ + スプレッドシート** 連携（JST バージョン表示付き）
- **STRUCTURE シートからのフォーム自動生成**（`?page=structure_form`）
- **受付フォーム（レガシー UI）**（`?page=reception`）

## ✅ 現在の機能

- 受付データの **新規登録 / 更新（編集）**（`saveReceptionData`）
- 受付データ **1件取得**（`getReceptionById`）
- 受付データ **一覧取得（要約列生成含む）**（`listReceptionIndex`）

> ※ **削除（Delete）API は未実装**。将来対応予定です（UI の「削除」記述は抑制）。
> 実装関数は Code.js のとおり：保存（新規/更新）、1件取得、一覧。

## 🧭 使い方

- **メイン**：デプロイ URL
- **受付フォーム（レガシー）**：`?page=reception`
- **デバッグ**：`?page=debug`
- **STRUCTURE フォーム（推奨）**：`?page=structure_form`（`STRUCTURE` シートから自動生成）
- **CSV 変換テスト**：`?page=test_csv`

> 旧 `?page=csv_converter` は削除しました（`?page=test_csv` を利用）。

## 🧪 段階テストルート（STRUCTURE→CSV→JSON→HTML / Form⇄DATA）

以下は tasks/001_taksk.md に基づく段階テスト用のページです。既存ページは保持されます。

- S1: `?page=test_structure_to_csv`：STRUCTURE を CSV で `<pre>` 表示
- S2: `?page=test_csv_to_json`：CSV→フィールド配列 JSON を `<pre>` 表示
- S3: `?page=test_json_to_form`：JSON→フォーム HTML を表示（空の初期値）
- ワンショット: `?page=reception_from_structure`：S1→S2→S3 を直列実行しフォーム表示
- 編集: `?page=reception_edit&row=2`：DATA 2行目→JSON 復元→フォーム初期値に反映

受け入れ確認（抜粋）
- S1: CSV が表示される
- S2: フィールド配列 JSON が表示される
- S3: input/textarea/select がDOMに出現
- F1: `reception_from_structure` から送信で `DATA` シートに1行追記
- F2: `reception_edit&row=N` で該当行の内容が初期値に反映

## 🗂️ ファイル構成（抜粋）

```
gas-reception-system/
├── Code.js
├── views/
│   ├── webapp.html
│   ├── reception_form.html
│   ├── structure_form.html   # STRUCTURE シート→フォーム生成
│   ├── debug.html
│   ├── simple_test.html      # 簡易テストページ
│   └── test_csv_converter.html
└── README.md
```

## 🧱 STRUCTURE シート仕様（要点）

- 10 行ヘッダー：1 行目 = 型、2–10 行目 = L1–L9（多段パス）
- 自動推定：`LIST`/`EXISTENSE` などの型を入力値から推定しヘッダー化
- `?page=structure_form` で HTML を生成して表示

## 🧰 スプレッドシート側メニュー

- メニューは **「📝 受付入力」→「🌐 Webアプリを開く」** を提供（GAS メニュー）

## 🔧 TODO（今後）

- 削除 API（行論理削除 or 物理削除）の追加
- 一覧 API の絞り込み条件・ページング（必要に応じて）

## 🧪 Monorepo（TypeScript）開発/CI

- ルートは npm workspaces を使用し、`packages/*` と `apps-script/` を管理
- コアロジック（Zod/構造解析）は `packages/core`、GAS 薄層は `packages/gas`
- 最小 UI 用の雛形は `packages/webui`

コマンド例:

```
npm install
npm run -w packages/core test
npm run -w packages/gas build
```

CI（GitHub Actions）は push/PR 時に lint/test/build を実行します。

## 🚀 GAS デプロイ（最小）

```
cd apps-script
clasp login
clasp push
clasp deploy -d "bootstrap"
```

apps-script には `dist/Code.js`（`packages/gas` からビルド）と `views/*.html` を配置します。
