@echo off
REM LibreOffice Windows自動インストールスクリプト

echo === PowerPoint スクリプトスライド変換ツール セットアップ ===
echo.

REM LibreOfficeダウンロード
echo 1. LibreOfficeをダウンロードしています...
cd /d "%TEMP%"
powershell -Command "Invoke-WebRequest -Uri 'https://download.documentfoundation.org/libreoffice/stable/7.6.4/win/x86_64/LibreOffice_7.6.4_Win_x86-64.msi' -OutFile 'LibreOffice.msi'"

REM 静默インストール
echo 2. LibreOfficeをインストールしています...
msiexec /i LibreOffice.msi /quiet /norestart

REM Python依存関係インストール
echo 3. Python依存関係をインストールしています...
cd /d "%~dp0"
pip install -r requirements_web.txt

REM クリーンアップ
echo 4. クリーンアップしています...
del "%TEMP%\LibreOffice.msi"

echo.
echo ✅ セットアップ完了！
echo.
echo 起動方法：
echo 1. このフォルダでコマンドプロンプトを開く
echo 2. python web_app.py
echo 3. ブラウザで http://localhost:8000
echo.
pause
