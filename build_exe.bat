@echo off
REM Bouwt dist\Waardering\Waardering.exe met het template erbij.
python -m pip install -r requirements.txt pyinstaller || exit /b 1
python -m PyInstaller --noconfirm --windowed --name Waardering ^
  --add-data "template;template" ^
  --collect-submodules anthropic ^
  waardering\__main__.py || exit /b 1
echo.
echo Klaar: dist\Waardering\Waardering.exe
