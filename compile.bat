@echo off
REM Lastrada Report Processor - Compilation Script
REM Run this from inside the project folder to produce standalone .exe files

echo ========================================
echo  Lastrada Report Processor - Build
echo ========================================
echo.

REM Use the venv Python/pip so all packages are available
set VENV=C:\Users\tcbit\PyCharmMiscProject\.venv\Scripts

echo Installing / updating dependencies...
"%VENV%\pip.exe" install pdfplumber openpyxl --quiet --prefer-binary
if errorlevel 1 (
    echo ERROR: pip install failed
    pause & exit /b 1
)

echo Installing PyInstaller...
"%VENV%\pip.exe" install pyinstaller --quiet
if errorlevel 1 (
    echo ERROR: PyInstaller install failed
    pause & exit /b 1
)

echo.
echo Step 1 of 2: Compiling pdf_filler.exe...
echo.

"%VENV%\pyinstaller.exe" ^
    --onefile ^
    --noconsole ^
    --name "pdf_filler" ^
    --hidden-import "pdfplumber" ^
    --hidden-import "openpyxl" ^
    --hidden-import "openpyxl.utils" ^
    pdf_filler.py

if errorlevel 1 (
    echo.
    echo ERROR: pdf_filler compilation failed - check messages above
    pause & exit /b 1
)

echo.
echo Step 2 of 2: Compiling Lastrada_Report_Processor.exe ^(with pdf_filler bundled^)...
echo.

"%VENV%\pyinstaller.exe" ^
    --onefile ^
    --windowed ^
    --name "Lastrada_Report_Processor" ^
    --hidden-import "pystray._win32" ^
    --hidden-import "PIL._tkinter_finder" ^
    --add-data "dist\pdf_filler.exe;." ^
    fpc_processor.py

if errorlevel 1 (
    echo.
    echo ERROR: Compilation failed - check messages above
    pause & exit /b 1
)

echo.
echo ========================================
echo  BUILD SUCCESSFUL
echo ========================================
echo.
echo Executables:
echo   dist\Lastrada_Report_Processor.exe  (distribute this one)
echo   dist\pdf_filler.exe                 (bundled inside the above)
echo.
pause
