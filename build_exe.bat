@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Python 없는 PC에서도 쓸 수 있는 exe 묶음을 만듭니다. (이 PC에 Python과 인터넷 필요, 5~10분)
python -m pip install -q pyinstaller
if errorlevel 1 ( echo pyinstaller 설치 실패 & pause & exit /b )
rmdir /s /q build dist 2>nul
python -m PyInstaller --noconfirm --noconsole --name "MailMaker" ^
  --hidden-import pandas --hidden-import openpyxl --hidden-import pdfplumber --hidden-import pypdf ^
  --hidden-import win32com --hidden-import win32com.client --hidden-import pythoncom --hidden-import pywintypes ^
  --collect-all pdfplumber --collect-all pypdfium2 --collect-all pdfminer ^
  app.py
if errorlevel 1 ( echo 빌드 실패 & pause & exit /b )
rem 실행에 필요한 파일을 exe 옆에 복사
copy /y make_drafts.py "dist\MailMaker\" >nul
copy /y make_po_mails.py "dist\MailMaker\" >nul
copy /y learn_from_outlook.py "dist\MailMaker\" >nul
copy /y config_tools.py "dist\MailMaker\" >nul
copy /y web_app.py "dist\MailMaker\" >nul
> "dist\MailMaker\MailMakerWeb.bat" echo @start "" "%%~dp0MailMaker.exe" --web
copy /y config.xlsx "dist\MailMaker\" >nul
copy /y config_kim.xlsx "dist\MailMaker\" >nul
copy /y GUIDE.txt "dist\MailMaker\" >nul
xcopy /e /i /y sample "dist\MailMaker\sample" >nul
echo.
echo 완료: dist\MailMaker 폴더를 통째로 복사해서 쓰세요. 안의 "MailMaker.exe"(창 방식) 또는 "MailMakerWeb.bat"(브라우저 방식)를 더블클릭하면 됩니다.
echo (계장님 PC: config_kim.xlsx 를 config.xlsx 로 이름 바꾸기)
pause
