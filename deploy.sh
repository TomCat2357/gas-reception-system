#!/bin/bash
set -o pipefail
export LC_ALL=C LANG=C

# Google AppSheetスタイル データ管理アプリ デプロイスクリプト

echo "🚀 Google AppSheetスタイル データ管理アプリのデプロイを開始します..."

# プロジェクトをプッシュ
echo "📤 プロジェクトファイルをGoogle Apps Scriptにプッシュ中..."
clasp push

if [ $? -eq 0 ]; then
    echo "✅ プッシュが完了しました"
else
    echo "❌ プッシュに失敗しました"
    exit 1
fi

# デプロイ
echo "🌐 Webアプリとしてデプロイ中..."

# JSON出力が使える場合は優先して利用し、だめなら通常出力を解析
DEPLOY_JSON=$(clasp deploy --description "Google AppSheetスタイル データ管理アプリ v$(date +%Y%m%d_%H%M%S)" --json 2>/dev/null)
DEPLOY_STATUS=$?

DEPLOYMENT_ID=""
WEB_APP_URL=""

if [ $DEPLOY_STATUS -eq 0 ] && echo "$DEPLOY_JSON" | grep -q '^[{\[]'; then
    echo "✅ デプロイが完了しました (JSON)"
    echo "$DEPLOY_JSON"
    # Nodeを使って堅牢にJSONを解析（jq不要）
    if command -v node >/dev/null 2>&1; then
        DEPLOYMENT_ID=$(node -e 'const fs=require("fs");const s=fs.readFileSync(0,"utf8");try{const o=JSON.parse(s);if(o.deploymentId){console.log(o.deploymentId)}else if(o.result&&o.result.deploymentId){console.log(o.result.deploymentId)}else if(o.entryPoints){const w=o.entryPoints.find(e=>e.webApp&&e.webApp.url);if(w&&w.webApp&&w.webApp.url){const m=w.webApp.url.match(/\/macros\/s\/([^/]+)\//);if(m)console.log(m[1]);}}}catch(e){}' <<< "$DEPLOY_JSON" || true)
        WEB_APP_URL=$(node -e 'const fs=require("fs");const s=fs.readFileSync(0,"utf8");try{const o=JSON.parse(s);if(o.webAppUrl){console.log(o.webAppUrl)}else if(o.entryPoints){const w=o.entryPoints.find(e=>e.webApp&&e.webApp.url);if(w&&w.webApp&&w.webApp.url)console.log(w.webApp.url);} }catch(e){}' <<< "$DEPLOY_JSON" || true)
    fi
else
    # JSON出力が使えないclaspの場合のフォールバック
    DEPLOY_OUTPUT=$(clasp deploy --description "Google AppSheetスタイル データ管理アプリ v$(date +%Y%m%d_%H%M%S)")
    if [ $? -ne 0 ]; then
        echo "❌ デプロイに失敗しました"
        exit 1
    fi
    echo "✅ デプロイが完了しました"
    echo "$DEPLOY_OUTPUT"
    # WebApp URLを出力から抽出
    WEB_APP_URL=$(echo "$DEPLOY_OUTPUT" | tr -d '\r' | grep -Eo 'https://script.google.com/macros/s/[^[:space:]]+' | head -n1)
    # URLからdeploymentIdを抽出
    if [ -n "$WEB_APP_URL" ]; then
        DEPLOYMENT_ID=$(echo "$WEB_APP_URL" | sed -n 's|.*/macros/s/\([^/]*\)/.*|\1|p')
    fi
    # それでも取れない場合はdeploymentId行から抽出
    if [ -z "$DEPLOYMENT_ID" ]; then
        DEPLOYMENT_ID=$(echo "$DEPLOY_OUTPUT" | tr -d '\r' | grep -Ei 'deployment id|deploymentId' | grep -Eo 'AKf[[:alnum:]_\-]+' | head -n1)
    fi
fi

# Script IDは参照リンク用に取得
SCRIPT_ID=$(grep '"scriptId"' .clasp.json | cut -d '"' -f4 | tr -d '\r')

echo ""
echo "🌟 Webアプリケーションの情報:"
if [ -n "$DEPLOYMENT_ID" ]; then
    printf '%s %s\n' "🆔 Deployment ID:" "$DEPLOYMENT_ID"
fi
if [ -n "$WEB_APP_URL" ]; then
    printf '%s %s\n' "🌐 Web App URL:" "$WEB_APP_URL"
elif [ -n "$DEPLOYMENT_ID" ]; then
    ADMIN_WEB_URL="https://script.google.com/macros/s/$DEPLOYMENT_ID/exec"
    printf '%s %s\n' "🌐 Web App URL:" "$ADMIN_WEB_URL"
fi
if [ -n "$SCRIPT_ID" ]; then
    printf '%s %s\n' "📋 Script ID:" "$SCRIPT_ID"
    ADMIN_EDIT_URL="https://script.google.com/home/projects/$SCRIPT_ID/edit"
    # 念のため既知の固定パスの二重文字を正規化（コピー時の異常対策）
    #ADMIN_EDIT_URL=${ADMIN_EDIT_URL//projects/projects}
    printf '%s %s\n' "⚙️  管理画面:" "$ADMIN_EDIT_URL"
fi

echo ""
echo "📖 次のステップ:"
echo "1. 管理画面でデプロイ設定を確認"
echo "2. アクセス権限を設定（全員 または 組織内のユーザー）"
echo "3. Web App URLを共有してアプリを使用開始"

echo ""
echo "🎉 デプロイが正常に完了しました！"
echo "📚 詳細な使用方法はREADME.mdを参照してください"
