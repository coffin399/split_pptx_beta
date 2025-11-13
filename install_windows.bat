@echo off
chcp 65001 >nul
echo === PowerPoint Script Slide Converter Setup ===
echo.

echo 1. Downloading LibreOffice...
cd /d "%TEMP%"
powershell -Command "Invoke-WebRequest -Uri 'https://download.documentfoundation.org/libreoffice/stable/7.6.4/win/x86_64/LibreOffice_7.6.4_Win_x86-64.msi' -OutFile 'LibreOffice.msi'"

echo 2. Installing LibreOffice...
msiexec /i LibreOffice.msi /quiet /norestart

echo 3. Installing Python dependencies...
cd /d "%~dp0"
pip install -r requirements_web.txt

echo 4. Cleaning up...
del "%TEMP%\LibreOffice.msi"

echo.
echo Setup complete!
echo.
echo To start:
echo 1. Open command prompt in this folder
echo 2. Run: python web_app.py
echo 3. Open browser: http://localhost:8000
echo.
pause
