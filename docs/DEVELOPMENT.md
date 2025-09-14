# Development Quickstart
## 初期化
npm i
npm run -w packages/core test
npm run -w packages/gas build
npm run -w packages/webui dev   # ローカルUI

## GAS 反映（必要時）
cd apps-script
clasp login
clasp push
clasp deploy -d "bootstrap"
