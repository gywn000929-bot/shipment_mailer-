@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [1/3] 패키지 설치 중...
python -m pip install -r requirements.txt
if errorlevel 1 ( echo 설치 실패. Python이 설치되어 있는지 확인하세요. & pause & exit /b )
echo.
echo [2/3] 샘플로 테스트 실행 (Outlook 사용 안 함)...
python make_drafts.py --ship "sample\ship_sample.xlsx" --air "sample\air_sample.xlsx" --dry-run -y --week-monday 2026-11-23
echo.
echo [3/3] 바탕화면에 바로가기 만들기...
powershell -NoProfile -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\MailMaker.lnk'); $s.TargetPath='%~dp00_run_app.bat'; $s.WorkingDirectory='%~dp0'; $s.IconLocation='%SystemRoot%\System32\shell32.dll,156'; $s.Save()"
powershell -NoProfile -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\MailMakerWeb.lnk'); $s.TargetPath='%~dp00_run_web.bat'; $s.WorkingDirectory='%~dp0'; $s.IconLocation='%SystemRoot%\System32\shell32.dll,14'; $s.Save()"
echo.
echo 끝. 바탕화면의 "메일 작성기"를 더블클릭하면 창이 열립니다.
pause
