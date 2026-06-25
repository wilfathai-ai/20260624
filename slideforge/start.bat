@echo off
chcp 65001 > nul
echo.
echo ==========================================
echo   SlideForge - AI Slide Auto Generator
echo ==========================================
echo.

echo 依存ライブラリを確認中...
pip install -r requirements.txt -q

echo pptxgenjsを確認中...
node -e "require('pptxgenjs')" 2>nul || call npm install -g pptxgenjs

echo バックエンドを起動中...
cd backend
start /B python app.py
cd ..
timeout /t 2 /nobreak > nul

echo ブラウザを起動中...
start "" "%~dp0frontend\index.html"

echo.
echo SlideForge を起動しました！
echo バックエンド: http://localhost:5050
echo このウィンドウを閉じると終了します
echo.
pause
