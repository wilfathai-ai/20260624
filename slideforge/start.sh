#!/bin/bash
set -e
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "🎨   SlideForge — AIスライド自動生成     "
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

# Python依存ライブラリ確認
echo "📦 依存ライブラリを確認中…"
pip install -r "$SCRIPT_DIR/requirements.txt" -q --break-system-packages 2>/dev/null \
  || pip install -r "$SCRIPT_DIR/requirements.txt" -q

# pptxgenjs確認
if ! node -e "require('pptxgenjs')" &>/dev/null; then
  echo "📦 pptxgenjs をインストール中…"
  npm install -g pptxgenjs
fi

# 既存バックエンドを停止
pkill -f "python.*app.py" 2>/dev/null || true
sleep 0.5

# バックエンド起動
echo "🚀 バックエンド起動中 (port 5050)…"
cd "$SCRIPT_DIR/backend"
python app.py &
BACKEND_PID=$!
sleep 1.5

# ブラウザを開く
FRONTEND="$SCRIPT_DIR/frontend/index.html"
echo "🌐 ブラウザを起動中…"
if command -v xdg-open &>/dev/null; then
  xdg-open "$FRONTEND"
elif command -v open &>/dev/null; then
  open "$FRONTEND"
else
  echo "   ブラウザで開いてください: file://$FRONTEND"
fi

echo ""
echo "✅ SlideForge を起動しました！"
echo "   バックエンド: http://localhost:5050"
echo "   Ctrl+C で終了"
echo ""

trap "kill $BACKEND_PID 2>/dev/null; echo '👋 終了しました'; exit 0" INT TERM
wait $BACKEND_PID
