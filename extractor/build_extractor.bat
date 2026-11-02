@echo off
REM Builds olkextract.exe (onedir) that the WinForms app bundles under /extractor.
REM Run on Windows with Python 3.11+ installed.
setlocal
cd /d "%~dp0"

python -m venv .venv
call .venv\Scripts\activate.bat

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

pyinstaller --noconfirm --onedir --name olkextract olkextract.py

echo.
echo Fertig: dist\olkextract\olkextract.exe
echo Beim Build von OLK2IMAP.App wird dieser Ordner automatisch nach bin\...\extractor kopiert.
endlocal
