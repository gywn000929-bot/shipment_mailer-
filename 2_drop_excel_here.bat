@echo off
chcp 65001 >nul
cd /d "%~dp0"
if "%~1"=="" (
  echo 선박리스트.xlsx 와 에어리스트.xlsx 를 이 파일 위에 함께 끌어다 놓으세요.
  echo 파일명에 "선박" 또는 "에어" 글자가 있어야 자동으로 구분됩니다.
  pause
  exit /b
)
set SHIP=
set AIR=
for %%F in (%*) do (
  echo %%~nF | findstr /i "에어 air" >nul && set AIR=%%~F
  echo %%~nF | findstr /i "선박 ship vessel" >nul && set SHIP=%%~F
)
if "%SHIP%%AIR%"=="" ( set SHIP=%~1 & set AIR=%~2 )
set ARGS=
if not "%SHIP%"=="" set ARGS=%ARGS% --ship "%SHIP%"
if not "%AIR%"=="" set ARGS=%ARGS% --air "%AIR%"
python make_drafts.py %ARGS%
pause
