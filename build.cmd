@echo off
REM Build the frozen Windows .exe — identical to the tagged CI build.
REM Output: dist\com_monitor.exe
REM Uses the committed icon.ico; run `uv run python make_icon.py` by hand to
REM regenerate it after changing the design.
setlocal
cd /d "%~dp0"

uv run pyinstaller --onefile --windowed --name com_monitor --icon icon.ico --add-data "icon.ico;." monitor.py || exit /b 1

echo.
echo Built dist\com_monitor.exe