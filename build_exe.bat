@echo off
rem SHEETDIFF - sestaveni jednoho SHEETDIFF.exe pomoci Nuitka (kompilator Zig se stahne sam).
rem Vysledek: dist\SHEETDIFF_v<verze>.exe (verze ze sheetdiff\__init__.py)
setlocal
cd /d "%~dp0"

python -c "import nuitka" 2>nul || (
    echo Nuitka neni nainstalovana: python -m pip install nuitka
    exit /b 1
)
for /f %%v in ('python -c "import sheetdiff; print(sheetdiff.__version__)"') do set VER=%%v
if not exist assets\sheetdiff.ico python tools\make_icon.py

python -m nuitka sheetdiff_main.py ^
    --mode=onefile ^
    --zig ^
    --assume-yes-for-downloads ^
    --enable-plugins=tk-inter ^
    --windows-console-mode=attach ^
    --windows-icon-from-ico=assets\sheetdiff.ico ^
    --onefile-tempdir-spec="{CACHE_DIR}/SHEETDIFF/%VER%" ^
    --nofollow-import-to=numpy,PIL,pandas,pytest,IPython,matplotlib,openpyxl.tests ^
    --company-name=SHEETDIFF ^
    --product-name=SHEETDIFF ^
    --file-description="SHEETDIFF - porovnani sesitu Excel" ^
    --file-version=%VER%.0 ^
    --product-version=%VER%.0 ^
    --output-dir=build ^
    --output-filename=SHEETDIFF_v%VER%.exe ^
    --remove-output
if errorlevel 1 exit /b 1

if not exist dist mkdir dist
move /y build\SHEETDIFF_v%VER%.exe dist\SHEETDIFF_v%VER%.exe >nul
echo.
echo Hotovo: %CD%\dist\SHEETDIFF_v%VER%.exe
