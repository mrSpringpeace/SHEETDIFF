@echo off
rem SHEETDIFF - okno pro porovnani dvou sesitu Excel.
rem Na tento soubor lze pretahnout dva sesity (stary, novy), predvyplni se.
setlocal
set "SD_HOME=%~dp0"

where pyw >nul 2>nul
if %errorlevel%==0 (
    start "" pyw -3 "%SD_HOME%sheetdiff_gui.pyw" %*
    exit /b 0
)
where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "%SD_HOME%sheetdiff_gui.pyw" %*
    exit /b 0
)
echo Python nebyl nalezen. Nainstalujte Python 3.11+ a knihovny: python -m pip install lxml openpyxl
pause
exit /b 1
