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

