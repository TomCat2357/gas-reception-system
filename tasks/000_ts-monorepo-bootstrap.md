# Task 000: TSモノレポ化 & GAS薄層化の下地づくり

## 背景
gas-reception-system を将来 GAS で運用するが、開発段階は GAS 非依存で回したい。
ロジックは純 TypeScript（core）に集約し、GAS は I/O の薄いアダプタに徹する。

## 目的
- packages/core: 構造シート解析/検証/正規化の純TS
- packages/gas: GAS アダプタ（doGet/saveReceptionData 等）= core を呼ぶだけ
- packages/webui: Vite+TS の開発用UI（本番時は HTML Service にバンドル）
- apps-script/: clasp 用（dist/Code.js と views/*.html を同期）

## 受け入れ条件（Definition of Done）
- [ ] ルートを npm workspaces 化（packages/*, apps-script を管理）
- [ ] packages/core に Zod を使った RecordSchema と parseStructure（9段ヘッダ→ツリーモデル）
- [ ] packages/gas から core API を呼ぶ（`global.doGet`,`global.saveReceptionData` を公開）
- [ ] packages/webui から `saveReceptionData()` を呼べる（GAS/ローカル両対応）
- [ ] GitHub Actions で lint + test が緑
- [ ] README に開発/デプロイ手順を反映（`clasp push` の最小手順含む）

## 実装ヒント
- GAS 側ビルド: esbuild で iife 出力（GAS のグローバル関数要件）
- スキーマ: Zod で型/検証を一元化（GAS 層は I/O のみ）
- 競合対策: save 時に LockService を使う設計余地をコメントで残す

## 参考（人間向け）
- Codex でこのタスクを実施する際は、このファイルと docs/ARCHITECTURE.md を前提に作業してください。
