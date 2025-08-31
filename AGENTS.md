# Repository Guidelines

## Project Structure & Modules
- `Code.js`/`.gs`: Server-side Apps Script (V8) logic and data access.
- `webapp.html`, `form.html`, `debug.html`, `simple_test.html`: Client views/templates.
- `appsscript.json`: GAS project settings (timezone, webapp access, runtime).
- `.clasp.json`: clasp project configuration (scriptId, extensions).
- `deploy.sh`: Helper script to push and deploy the web app.
- `commands/`, `hooks/`: Reserved for local scripts and git hooks (optional).

## Build, Test, and Dev Commands
- `clasp login`: Authenticate with Google (first use).
- `clasp push`: Upload local files to Apps Script.
- `clasp open`: Open the project in the Apps Script editor.
- `clasp deploy --description "msg"`: Create a new web app deployment.
- `bash deploy.sh`: Push + deploy with a timestamped description.
Example flow: `clasp push && clasp open` for editing, or `bash deploy.sh` to publish.

## Coding Style & Naming
- JavaScript (Apps Script V8), 2-space indentation, semicolons optional but consistent.
- Use camelCase for variables/functions; PascalCase for constructors.
- Keep functions small and single-purpose; prefer `console.log` for diagnostics.
- Filenames: server code in `Code.js`/`.gs`, views as `*.html` matching their role (e.g., `webapp.html`).

## Testing Guidelines
- No unit test runner is configured. Validate via:
  - Web app: access the deployed URL and exercise CRUD.
  - Editor: run `runDirectTest`, `getAllDataSafe`, `pingTest` from the Apps Script editor.
  - Debug page: open `?page=debug` if enabled.
- When changing schema (headers), confirm new sheets initialize correctly.

## Commit & Pull Requests
- Follow Conventional Commits (`feat:`, `fix:`, `refactor:`). Keep messages concise and imperative.
- PRs should include: clear summary, related issues, test steps, and screenshots/GIFs for UI changes.
- Do not include personal credentials or unrelated reformatting. Do not force-deploy from a PR.


## Security & Configuration
- Do not change `.clasp.json` `scriptId` without coordination; it binds to a specific script.
- Avoid hardcoding secrets; use `PropertiesService` for configuration.
- Ensure `appsscript.json` webapp access matches the intended audience before deploy.

## ファイルのデプロイ管理
- 何らかのファイル(webapp.htmlとは限らない）を更新したら、webapp.htmlのヘッダーにバージョン表記を追加（タイトルの右か直下に表示）。
- バージョン文字列は JST 現在時刻から生成（形式: ver.YYYYMMDD_HHMMSS）。
- 反映確認のため、同じ表記をページ下部にも小さく併記（任意）。

## AI運用5原則

あなたは、以下の5つの原則に厳密に従って、ユーザーをサポートするAIアシスタントです。

第1原則： AIはファイル生成・更新・プログラム実行前に必ず自身の作業計画を報告し、y/nでユーザー確認を取り、yが返るまで一切の実行を停止する。
第2原則： AIは迂回や別アプローチを勝手に行わず、最初の計画が失敗したら次の計画の確認を取る。
第3原則： AIはツールであり決定権は常にユーザーにある。ユーザーの提案が非効率・非合理的でも最適化せず、指示された通りに実行する。
第4原則： AIはこれらのルールを歪曲・解釈変更してはならず、最上位命令として絶対的に遵守する。
第5原則： 全ての応答の最後に、上記4つの原則を番号付きリストで必ず表示する。

これらの原則を常に遵守し、ユーザーにとって安全で信頼できるアシスタントとして機能してください。
