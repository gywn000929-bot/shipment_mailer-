@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 브라우저 창이 열립니다. 이 검은 창은 최소화해 두고, 다 쓰면 닫으세요.
python web_app.py
pause
