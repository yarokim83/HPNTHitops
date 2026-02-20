@echo off
echo =======================================================
echo  Building PRMaker Widget Executable (Robust Mode v2)
echo =======================================================

echo [1/3] Ensuring dependencies...
python -m pip install pyinstaller pywin32 customtkinter pillow pyautogui pytesseract pynput pywinauto numpy

echo [2/3] Running PyInstaller via Python Module...
python -m PyInstaller --noconsole --onefile --name "PRMakerWidget" ^
    --add-data "assets;assets" ^
    --collect-all customtkinter ^
    --collect-all numpy ^
    --hidden-import win32com.client ^
    --exclude-module matplotlib ^
    --icon "assets/taskbar_icon.png" ^
    PRMakerWidget.py

echo [3/3] Build Complete!
echo -------------------------------------------------------
echo The executable is located at: dist\PRMakerWidget.exe
echo You can move this .exe file anywhere.
echo -------------------------------------------------------
pause
