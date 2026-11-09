@echo off
REM Сборка PDF2DWG.exe на Windows (нужен Python 3.10+)
cd /d %~dp0
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name PDF2DWG --icon icon.ico --add-data "icon.ico;." ^
  --collect-all customtkinter --collect-all tkinterdnd2 ^
  app.py
echo.
echo Готово: dist\PDF2DWG.exe
pause
