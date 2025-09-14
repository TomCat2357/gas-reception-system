# Architecture

- packages/core … 純 TypeScript のドメイン層（構造シート→ツリーモデル、検証、正規化）
- packages/gas  … GAS アダプタ（SpreadsheetApp, HtmlService）。core を呼ぶだけ
- packages/webui … Vite+TS。開発はローカル、公開は HTML Service に組み込み
- apps-script/   … clasp 用プロジェクト（appsscript.json, views/*.html, dist/Code.js）

## データモデル（概要）
- 9段ヘッダの A1: I… → ツリー化し、leaf にフィールド型/必須等のメタを付与
- `saveReceptionData(payload)` は core の Zod スキーマで検証→正規化→GASへ書込

## ビルド & デプロイ（最小）
- `npm run -w packages/gas build` → `packages/gas/dist/Code.js`
- `clasp push`（apps-script/ から）→ Web アプリ更新
