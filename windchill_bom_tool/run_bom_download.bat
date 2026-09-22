@echo off
setlocal
cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    echo Python wurde nicht gefunden. Bitte Python installieren und erneut starten.
    pause
    exit /b 1
)

python -m pip show playwright >nul 2>&1
if errorlevel 1 (
    echo Installiere benoetigte Bibliothek "playwright" ...
    python -m pip install --quiet -r requirements.txt
)

python bom_download.py %*
pause
