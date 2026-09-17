@echo off
setlocal
cd /d "%~dp0"
if not exist "build\venv\Scripts\python.exe" (
    py -3.12 -m venv build\venv
    if errorlevel 1 exit /b 1
)
"build\venv\Scripts\python.exe" -m pip install -r requirements-build.lock
if errorlevel 1 exit /b 1
"build\venv\Scripts\python.exe" -B -m unittest discover -s tests -p "*_regression.py"
if errorlevel 1 exit /b 1
"build\venv\Scripts\python.exe" -m PyInstaller --noconfirm --distpath dist\release --workpath build\release PRMakerWidget.spec
if errorlevel 1 exit /b 1
"dist\release\PRMakerWidget.exe" --self-check
if errorlevel 1 exit /b 1
echo Build verified: dist\release\PRMakerWidget.exe
echo Install with: powershell -ExecutionPolicy Bypass -File install.ps1
