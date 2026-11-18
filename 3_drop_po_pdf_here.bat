@echo off
chcp 65001 >nul
cd /d "%~dp0"
if "%~1"=="" (
  echo ERP에서 뽑은 주문서 PDF를 이 파일 위에 끌어다 놓으세요. 여러 개도 됩니다.
  pause
  exit /b
)
python make_po_mails.py %*
pause
